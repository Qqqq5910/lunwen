import re

NUMBER_LIST = r"\d+(?:\s*[-–—~～,，;；、]\s*\d+)*"
MARKER = r"(?:\[\s*" + NUMBER_LIST + r"\s*\]|［\s*" + NUMBER_LIST + r"\s*］|【\s*" + NUMBER_LIST + r"\s*】|〔\s*" + NUMBER_LIST + r"\s*〕)"
CITATION_PATTERN = re.compile(r"(" + MARKER + r")")
SEPARATOR = r"(?:\s*(?:、|,|，|;|；)\s*|\s*)"
CITATION_SEQUENCE_PATTERN = re.compile(r"(" + MARKER + r"(?:" + SEPARATOR + MARKER + r")+)")
SINGLE_MARKER_PATTERN = re.compile(MARKER)


def normalize_marker(raw):
    value = raw.strip()
    for left in ["［", "【", "〔"]:
        value = value.replace(left, "[")
    for right in ["］", "】", "〕"]:
        value = value.replace(right, "]")
    return value.replace("，", ",").replace("；", ",").replace("、", ",").replace(";", ",")


def expand_citation_numbers(raw):
    value = normalize_marker(raw).strip("[]").replace(" ", "")
    value = value.replace("–", "-").replace("—", "-").replace("～", "~")
    numbers = []
    for part in value.split(","):
        if not part:
            continue
        sep = "~" if "~" in part else "-" if "-" in part else None
        if sep:
            left, right = part.split(sep, 1)
            if left.isdigit() and right.isdigit():
                start = int(left)
                end = int(right)
                if start <= end and start > 0:
                    numbers.extend(range(start, end + 1))
        elif part.isdigit():
            number = int(part)
            if number > 0:
                numbers.append(number)
    return numbers


def marker_numbers_from_sequence(text):
    numbers = []
    for marker in SINGLE_MARKER_PATTERN.findall(text):
        numbers.extend(expand_citation_numbers(marker))
    return numbers


def compact_ranges(numbers):
    ordered = sorted(set(numbers))
    if not ordered:
        return []
    groups = []
    start = ordered[0]
    prev = ordered[0]
    for number in ordered[1:]:
        if number == prev + 1:
            prev = number
        else:
            groups.append((start, prev))
            start = number
            prev = number
    groups.append((start, prev))
    return groups


def bracket_pair(style):
    if style == "【】":
        return "【", "】"
    if style == "〔〕":
        return "〔", "〕"
    if style == "［］":
        return "［", "］"
    return "[", "]"


def format_numbers(numbers, citation_rule=None):
    rule = citation_rule or {}
    left, right = bracket_pair(rule.get("bracket_style", "[]"))
    range_sep = rule.get("range_separator", "~")
    list_sep = rule.get("list_separator", ",")
    parts = []
    for start, end in compact_ranges(numbers):
        if start == end:
            parts.append(str(start))
        else:
            parts.append(f"{start}{range_sep}{end}")
    return f"{left}{list_sep.join(parts)}{right}"


def compact_sequence_text(text, citation_rule=None):
    numbers = marker_numbers_from_sequence(text)
    if not numbers:
        return text
    return format_numbers(numbers, citation_rule)


def normalize_single_marker_text(text, citation_rule=None):
    numbers = expand_citation_numbers(text)
    if not numbers:
        return text
    return format_numbers(numbers, citation_rule)
