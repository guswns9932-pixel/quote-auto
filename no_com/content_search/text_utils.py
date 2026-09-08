"""검색 패턴, 스니펫, 확장자 필터 등 파일 형식과 무관한 순수 텍스트 로직."""

import re

MAX_HITS_PER_FILE = 30


def normalize_text(value):
    """검색 및 결과 표시용으로 문자열을 정리한다."""
    if value is None:
        return ""
    text = str(value)
    return re.sub(r"\s+", " ", text).strip()


def make_snippet(text, start, end, radius=70):
    """일치 위치 전후 문맥을 짧게 잘라 반환한다."""
    flat = normalize_text(text)
    if not flat:
        return ""
    start = max(0, min(start, len(flat)))
    end = max(start, min(end, len(flat)))
    left = max(0, start - radius)
    right = min(len(flat), end + radius)
    prefix = "…" if left > 0 else ""
    suffix = "…" if right < len(flat) else ""
    return prefix + flat[left:right] + suffix


def find_matches(text, pattern):
    """
    text 안에서 패턴을 찾아 (시작, 끝, 일치한 키워드, 스니펫) 목록을 반환한다.
    pattern은 re.Pattern 또는 SearchPattern이다. 일치한 키워드를 알 수 없으면
    (예: build_pattern을 거치지 않은 단순 re.Pattern) None이 된다.
    """
    if text is None:
        return []
    text = str(text)

    results = []
    for match in pattern.finditer(text):
        keyword = _matched_keyword(match, pattern)
        results.append((
            match.start(), match.end(), keyword,
            make_snippet(text, match.start(), match.end()),
        ))
        if len(results) >= MAX_HITS_PER_FILE:
            break
    return results


def _matched_keyword(match, pattern):
    keywords = getattr(pattern, "keywords", None)
    if not keywords or not match.lastgroup:
        return None
    try:
        index = int(match.lastgroup[len("kw"):])
    except ValueError:
        return None
    if 0 <= index < len(keywords):
        return keywords[index]
    return None


def _wildcard_to_regex(keyword):
    """
    '*'(임의 길이 문자열)와 '?'(문자 1개)만 지원하는 와일드카드 표현을
    부분 문자열 검색용 정규식 조각으로 바꾼다(전체 일치가 아닌 finditer용).
    """
    parts = []
    for char in keyword:
        if char == "*":
            parts.append(".*")
        elif char == "?":
            parts.append(".")
        else:
            parts.append(re.escape(char))
    return "".join(parts)


def parse_keywords(text):
    """
    쉼표/세미콜론/줄바꿈으로 구분된 여러 키워드를 리스트로 반환한다.
    공백은 구분자로 쓰지 않는다 ("2024 보고서"처럼 키워드 자체에 공백이 들어갈 수 있어서다).
    """
    if not text:
        return []
    tokens = re.split(r"[,;\n]+", text)
    return [token.strip() for token in tokens if token.strip()]


class SearchPattern:
    """
    컴파일된 정규식과 '그룹 이름 → 원본 키워드' 매핑을 함께 들고 다니는 래퍼.
    re.Pattern은 속성을 추가로 붙일 수 없어서(부착 시 AttributeError) 대신 이 얇은
    래퍼로 감싼다. .search()/.finditer() 등 나머지 메서드는 내부 regex로 그대로
    위임되므로, 다른 코드는 이 래퍼를 일반 컴파일 패턴처럼 넘겨 쓰면 된다.
    """

    __slots__ = ("regex", "keywords")

    def __init__(self, regex, keywords):
        self.regex = regex
        self.keywords = keywords

    def __getattr__(self, name):
        return getattr(self.regex, name)


def build_pattern(keywords, use_wildcard, case_sensitive):
    """
    keywords는 문자열 하나 또는 여러 키워드의 리스트일 수 있다.
    여러 개면 그중 하나라도 일치하면 매치되는(OR) 패턴 하나로 합치고, 어떤
    매치가 어느 키워드에서 왔는지 find_matches()가 알아낼 수 있도록 각 키워드를
    이름 붙은 그룹(kw0, kw1, ...)으로 감싼다.
    """
    if isinstance(keywords, str):
        keywords = [keywords]

    flags = 0 if case_sensitive else re.IGNORECASE
    fragments = []
    for index, keyword in enumerate(keywords):
        expression = _wildcard_to_regex(keyword) if use_wildcard else re.escape(keyword)
        fragments.append(f"(?P<kw{index}>{expression})")
    expression = "|".join(fragments)
    regex = re.compile(expression, flags)
    return SearchPattern(regex, list(keywords))


def parse_extension_filter(text):
    """
    "txt, .pdf *.docx" 같은 입력을 {'.txt', '.pdf', '.docx'} 형태로 정규화한다.
    비어 있으면 빈 집합(필터 없음, 전체 검색)을 반환한다.
    """
    if not text:
        return set()

    tokens = re.split(r"[,;\s]+", text.strip())
    extensions = set()
    for token in tokens:
        token = token.strip().lstrip("*")
        if not token:
            continue
        if not token.startswith("."):
            token = "." + token
        extensions.add(token.lower())
    return extensions
