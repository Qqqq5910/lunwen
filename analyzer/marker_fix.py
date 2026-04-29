def is_digit(ch):
    return ch >= '0' and ch <= '9'


def is_non_ascii_letter(ch):
    return bool(ch) and ord(ch) > 127


def plain_spans(text, valid_numbers=None):
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
    left = text[max(0, start - 20):start].lower().rstrip()
    right = text[end:min(len(text), end + 14)].lower().lstrip()
    if not left or not right:
        return False
    if left.endswith('[') or left.endswith('('):
        return False
    left_keys = ['et al', 'ref', 'reference']
    right_keys = ['proposed', 'reported', 'showed', 'found', 'used', 'introduced']
    left_ok = any(key in left for key in left_keys) or is_non_ascii_letter(left[-1])
    right_ok = any(right.startswith(key) for key in right_keys) or is_non_ascii_letter(right[0])
    return left_ok and right_ok
