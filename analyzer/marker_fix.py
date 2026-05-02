def is_digit(ch):
    return ch >= '0' and ch <= '9'


def is_non_ascii_letter(ch):
    return bool(ch) and ord(ch) > 127


def _is_chinese_or_letter(ch):
    if not ch:
        return False
    if '\u4e00' <= ch <= '\u9fff':
        return True
    return ch.isalpha()


def _prev_non_space(text, start):
    i = start - 1
    while i >= 0 and text[i] in {' ', '\u00a0', '\u3000', '\t'}:
        i -= 1
    return text[i] if i >= 0 else ''


def _next_non_space(text, end):
    i = end
    while i < len(text) and text[i] in {' ', '\u00a0', '\u3000', '\t'}:
        i += 1
    return text[i] if i < len(text) else ''


def _window(text, start, end, radius=8):
    return text[max(0, start - radius):min(len(text), end + radius)]


def _is_numeric_context(text, start, end):
    left = text[start - 1] if start > 0 else ''
    right = text[end] if end < len(text) else ''
    prev_ch = _prev_non_space(text, start)
    next_ch = _next_non_space(text, end)
    nearby = _window(text, start, end)

    if left in '.．/' or right in '.．/%％':
        return True
    if left.isdigit() or right.isdigit():
        return True
    if prev_ch in '第表图式章节年月日页-' or next_ch in '章节年月日页个条项点倍%％':
        return True
    if any(unit in nearby for unit in ['年', '月', '日', '小时', '分钟', '秒', '页']):
        if prev_ch in '年月日第表图式章节' or next_ch in '年月日页':
            return True
    return False


def plain_spans(text, valid_numbers=None):
    """Find citation-looking bare numbers such as “Spirtes等人32提出”.

    The matcher is deliberately conservative: it only accepts bare numbers in
    author/work-style citation contexts and rejects date, section, page, percent,
    and list-number contexts even when the number exists in the bibliography.
    """
    valid = set(valid_numbers or [])
    out = []
    i = 0
    total = len(text)
    while i < total:
        if not is_digit(text[i]):
            i += 1
            continue
        j = i
        while j < total and is_digit(text[j]):
            j += 1
        raw = text[i:j]
        if len(raw) <= 3:
            number = int(raw)
            if (not valid or number in valid) and looks_like_citation(text, i, j):
                out.append((i, j, number))
        i = j
    return out


def looks_like_citation(text, start, end):
    if _is_numeric_context(text, start, end):
        return False

    left = text[max(0, start - 24):start].lower().rstrip()
    right = text[end:min(len(text), end + 18)].lower().lstrip()
    if not left or not right:
        return False
    if left.endswith('[') or left.endswith('(') or left.endswith('（'):
        return False

    english_left_keys = ['et al', 'ref', 'reference']
    english_right_keys = ['proposed', 'reported', 'showed', 'found', 'used', 'introduced']
    chinese_left_keys = ['等人', '学者', '研究', '工作']
    chinese_right_keys = [
        '提出', '指出', '认为', '表明', '证明', '发现', '验证', '说明', '显示', '采用',
        '引入', '构建', '设计', '进一步', '最早', '首次'
    ]

    left_ok = (
        any(key in left for key in english_left_keys)
        or any(key in left for key in chinese_left_keys)
    )
    right_ok = (
        any(right.startswith(key) for key in english_right_keys)
        or any(right.startswith(key) for key in chinese_right_keys)
    )

    if right.startswith(('）', ')', '、', '.', '．')):
        return False

    return left_ok and right_ok
