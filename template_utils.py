import re
from datetime import datetime
from typing import Any, Iterable, Optional

from dateutil.relativedelta import relativedelta

try:
    from pymorphy3 import MorphAnalyzer
except ImportError:  # pragma: no cover
    MorphAnalyzer = None


_DATE_FORMATS = (
    "%d.%m.%Y",
    "%Y-%m-%d",
    "%d.%m.%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
)
_RELATIVE_MONTHS_PATTERN = re.compile(
    r"^(?P<base>.+?)\s*\+\s*(?P<months>\d+)\s*месяц(?:а|ев)?$",
    re.IGNORECASE,
)
_GENITIVE_CASE_PATTERN = re.compile(r"^(?P<base>.+?)\s*\+\s*РП$", re.IGNORECASE)
_DATIVE_CASE_PATTERN = re.compile(r"^(?P<base>.+?)\s*\+\s*ДП$", re.IGNORECASE)
_ANGLE_BRACKET_PATTERN = re.compile(r"<[^<>]+>")
_TOKEN_PATTERN = re.compile(r"^([^A-Za-zА-Яа-яЁё0-9-]*)(.*?)([^A-Za-zА-Яа-яЁё0-9-]*)$")
_QUOTE_MARKERS = ('"', "'", "«", "»")
_GENITIVE_SKIP_WORDS = {
    "и", "или", "с", "со", "по", "на", "от", "до", "из", "для", "при", "под", "над",
}
_INDECLINABLE_WORDS = {
    "хостес",
}
_MORPH = MorphAnalyzer() if MorphAnalyzer is not None else None
_PERSON_NAME_KEY_HINTS = (
    "фио",
    "полное имя",
    "руководитель",
    "председател",
    "член комиссии",
    "работник",
)
_FULL_ORG_NAME_KEY_HINTS = (
    "полное наименование",
    "полное наименование+рп",
    "полное наименование+дп",
)
_IP_FULL_NAME_PREFIX = "Индивидуальный предприниматель "
_CASE_BASE_NAME_ALIASES = {
    "руководителя": "Руководитель",
    "должности руководителя": "Должность руководителя",
    "основания подписи": "Основание подписи",
}


def _parse_date(value: Any) -> Optional[datetime]:
    if value in (None, ""):
        return None

    text = str(value).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _format_date(value: datetime) -> str:
    return value.strftime("%d.%m.%Y")


def _candidate_base_keys(key: str) -> list[str]:
    stripped_key = key.strip()
    if stripped_key.startswith("<") and stripped_key.endswith(">"):
        bare_key = stripped_key[1:-1].strip()
        return [stripped_key, bare_key]
    return [stripped_key, f"<{stripped_key}>"]


def _candidate_base_names(base_name: str) -> list[str]:
    normalized_name = _normalize_key_name(base_name)
    candidates: list[str] = []

    canonical_name = _CASE_BASE_NAME_ALIASES.get(normalized_name)
    if canonical_name:
        candidates.append(canonical_name)

    candidates.append(base_name.strip())

    deduped: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        normalized_candidate = _normalize_key_name(candidate)
        if normalized_candidate in seen:
            continue
        seen.add(normalized_candidate)
        deduped.append(candidate)

    return deduped


def _normalize_key_name(value: Any) -> str:
    text = str(value).strip()
    if text.startswith("<") and text.endswith(">"):
        text = text[1:-1].strip()
    return " ".join(text.split()).lower()


def _normalize_for_inflection_compare(value: Any) -> str:
    return " ".join(str(value).strip().lower().replace("ё", "е").split())


def _inflect_word_to_genitive(word: str) -> str:
    if not word:
        return word

    match = _TOKEN_PATTERN.match(word)
    if not match:
        return word

    prefix, core, suffix = match.groups()
    if not core:
        return word

    lower_core = core.lower()
    normalized_core = lower_core.replace("ё", "е")
    if normalized_core == "хостес":
        return word
    if (
        len(core) <= 2
        or
        any(marker in core for marker in _QUOTE_MARKERS)
        or any(char.isdigit() for char in core)
        or core.isupper()
        or lower_core in _GENITIVE_SKIP_WORDS
        or normalized_core in _INDECLINABLE_WORDS
        or lower_core == "ответственностью"
    ):
        return word

    if _MORPH is not None:
        parsed = _choose_best_genitive_parse(core)
        if parsed is not None:
            inflected = parsed.inflect({"gent"})
            if inflected is not None:
                inflected_word = inflected.word
            else:
                inflected_word = _fallback_inflect_word_to_genitive(core)
        else:
            inflected_word = _fallback_inflect_word_to_genitive(core)
    else:
        inflected_word = _fallback_inflect_word_to_genitive(core)

    if inflected_word == core:
        return word

    if core[:1].isupper():
        inflected_word = inflected_word.capitalize()

    return f"{prefix}{inflected_word}{suffix}"


def _choose_best_genitive_parse(word: str):
    parses = _MORPH.parse(word)
    if not parses:
        return None

    def priority(parse) -> tuple[int, float]:
        pos = getattr(parse.tag, "POS", None)
        score = getattr(parse, "score", 0.0)

        if pos == "NOUN":
            return (0, -score)
        if pos in {"NPRO", "NUMR"}:
            return (1, -score)
        if pos in {"ADJF", "ADJS", "PRTF", "PRTS"}:
            return (3, -score)
        return (2, -score)

    return sorted(parses, key=priority)[0]


def _guess_person_gender(words: list[str]) -> Optional[str]:
    for word in words:
        lower_word = word.lower()
        if lower_word.endswith(("овна", "евна", "ична", "инична")):
            return "femn"
        if lower_word.endswith(("ович", "евич", "ич")) and not lower_word.endswith(("оглы", "улы", "кызы")):
            return "masc"

    for word in words:
        lower_word = word.lower()
        if lower_word.endswith(("ова", "ева", "ина", "ына", "ая")):
            return "femn"

    return None


def _looks_like_person_name(value: Any, base_name: Optional[str] = None) -> bool:
    text = str(value).strip()
    if not text:
        return False

    lowered_base_name = (base_name or "").strip("<>").lower()
    if any(hint in lowered_base_name for hint in _PERSON_NAME_KEY_HINTS):
        return True

    if any(marker in text for marker in _QUOTE_MARKERS) or any(char.isdigit() for char in text):
        return False

    words = [word for word in text.split() if word]
    if not 2 <= len(words) <= 4:
        return False

    cyrillic_words = 0
    titled_words = 0
    for word in words:
        match = _TOKEN_PATTERN.match(word)
        if not match:
            continue
        _, core, _ = match.groups()
        if not core:
            continue
        if re.fullmatch(r"[А-Яа-яЁё-]+", core):
            cyrillic_words += 1
        if core[:1].isupper():
            titled_words += 1

    return cyrillic_words == len(words) and titled_words >= max(2, len(words) - 1)


def _looks_like_full_org_name(base_name: Optional[str]) -> bool:
    normalized_base_name = _normalize_key_name(base_name or "")
    return any(hint in normalized_base_name for hint in _FULL_ORG_NAME_KEY_HINTS)


def _split_ip_full_name(value: Any) -> Optional[tuple[str, str]]:
    text = str(value or "").strip()
    if not text.startswith(_IP_FULL_NAME_PREFIX):
        return None

    person_name = text[len(_IP_FULL_NAME_PREFIX):].strip()
    if not _looks_like_person_name(person_name, "Руководитель"):
        return None

    return _IP_FULL_NAME_PREFIX.strip(), person_name


def _choose_best_person_name_parse(word: str, gender_hint: Optional[str], expected_kind: Optional[str] = None):
    parses = _MORPH.parse(word)
    if not parses:
        return None

    def priority(parse) -> tuple[int, int, int, int, float]:
        tag = parse.tag
        score = getattr(parse, "score", 0.0)

        if "Surn" in tag:
            kind = 0
        elif "Name" in tag:
            kind = 1
        elif "Patr" in tag:
            kind = 2
        elif getattr(tag, "POS", None) == "NOUN":
            kind = 3
        else:
            kind = 4

        expected_rank = 1
        if expected_kind is not None:
            expected_rank = 0 if expected_kind in tag else 1

        number_rank = 0 if getattr(tag, "number", None) in (None, "sing") else 1
        case_rank = 0 if getattr(tag, "case", None) in (None, "nomn", "accs") else 1

        gender_rank = 1
        if gender_hint:
            gender_rank = 0 if gender_hint in tag else 1

        return (expected_rank, number_rank, case_rank, kind, gender_rank, -score)

    return sorted(parses, key=priority)[0]


def _fallback_inflect_person_name_word_to_genitive(word: str) -> str:
    lower_word = word.lower()

    if lower_word.endswith(("овна", "евна", "ична", "инична")):
        return word[:-1] + "ы"
    if lower_word.endswith(("ович", "евич")):
        return word + "а"
    if lower_word.endswith(("оглы", "улы", "кызы")):
        return word
    if lower_word.endswith(("ова", "ева", "ина", "ына")):
        return word[:-1] + "ой"
    if lower_word.endswith(("ская", "цкая")):
        return word[:-2] + "ой"
    if lower_word.endswith("ая"):
        return word[:-2] + "ой"
    if lower_word.endswith("ия"):
        return word[:-1] + "и"
    if lower_word.endswith("ья"):
        return word[:-1] + "и"
    if lower_word.endswith("я"):
        return word[:-1] + "и"
    if lower_word.endswith("а"):
        preceding = lower_word[-2:-1]
        return word[:-1] + ("и" if preceding in "гкхжчшщц" else "ы")
    if lower_word.endswith("й"):
        return word[:-1] + "я"
    if lower_word.endswith("ь"):
        return word[:-1] + "я"
    if lower_word and lower_word[-1] in "бвгджзклмнпрстфхцчшщ":
        return word + "а"

    return word


def _inflect_person_name_word_to_genitive(word: str, gender_hint: Optional[str], expected_kind: Optional[str] = None) -> str:
    if not word:
        return word

    match = _TOKEN_PATTERN.match(word)
    if not match:
        return word

    prefix, core, suffix = match.groups()
    if not core or any(char.isdigit() for char in core):
        return word

    inflected_word = core
    if _MORPH is not None:
        parsed = _choose_best_person_name_parse(core, gender_hint, expected_kind)
        if parsed is not None:
            inflected = parsed.inflect({"gent"})
            if inflected is not None:
                inflected_word = inflected.word
            if inflected_word == core:
                inflected_word = _fallback_inflect_person_name_word_to_genitive(core)
        else:
            inflected_word = _fallback_inflect_person_name_word_to_genitive(core)
    else:
        inflected_word = _fallback_inflect_person_name_word_to_genitive(core)

    if core[:1].isupper():
        inflected_word = inflected_word.capitalize()

    return f"{prefix}{inflected_word}{suffix}"


def _fallback_inflect_word_to_genitive(word: str) -> str:
    lower_word = word.lower()

    if lower_word.endswith(("ова", "ева", "ина", "ына")):
        return word[:-1] + "ой"
    if lower_word.endswith(("ская", "цкая")):
        return word[:-2] + "ой"
    if lower_word.endswith("ая"):
        return word[:-2] + "ой"
    if lower_word.endswith("яя"):
        return word[:-2] + "ей"
    if lower_word.endswith("ой"):
        if not word[:1].isupper():
            return word
        return word[:-2] + "ого"
    if lower_word.endswith("ей"):
        if not word[:1].isupper():
            return word
        return word[:-2] + "ея"
    if lower_word.endswith(("ий", "ый")):
        suffix = "его" if lower_word.endswith(("ний", "чий", "щий", "жий", "ший")) else "ого"
        return word[:-2] + suffix
    if lower_word.endswith("ия"):
        return word[:-2] + "ии"
    if lower_word.endswith("ья"):
        return word[:-2] + "ьи"
    if lower_word.endswith("я"):
        return word[:-1] + "и"
    if lower_word.endswith("а"):
        preceding = lower_word[-2:-1]
        return word[:-1] + ("и" if preceding in "гкхжчшщц" else "ы")
    if lower_word.endswith("ь"):
        if lower_word.endswith(("ость", "есть", "асть", "сть")):
            return word[:-1] + "и"
        return word[:-1] + "я"
    if lower_word.endswith("й"):
        return word[:-1] + "я"
    if lower_word.endswith("ие"):
        return word[:-2] + "их"
    if lower_word.endswith("ые"):
        return word[:-2] + "ых"
    if lower_word and lower_word[-1] in "бвгджзклмнпрстфхцчшщ":
        return word + "а"

    return word


def _inflect_word_to_dative(word: str) -> str:
    if not word:
        return word

    match = _TOKEN_PATTERN.match(word)
    if not match:
        return word

    prefix, core, suffix = match.groups()
    if not core:
        return word

    lower_core = core.lower()
    normalized_core = lower_core.replace("ё", "е")
    if normalized_core == "хостес":
        return word
    if (
        len(core) <= 2
        or any(marker in core for marker in _QUOTE_MARKERS)
        or any(char.isdigit() for char in core)
        or core.isupper()
        or lower_core in _GENITIVE_SKIP_WORDS
        or normalized_core in _INDECLINABLE_WORDS
        or lower_core == "ответственностью"
    ):
        return word

    if _MORPH is not None:
        parsed = _choose_best_genitive_parse(core)
        if parsed is not None:
            inflected = parsed.inflect({"datv"})
            if inflected is not None:
                inflected_word = inflected.word
            else:
                inflected_word = _fallback_inflect_word_to_dative(core)
        else:
            inflected_word = _fallback_inflect_word_to_dative(core)
    else:
        inflected_word = _fallback_inflect_word_to_dative(core)

    if inflected_word == core:
        return word

    if core[:1].isupper():
        inflected_word = inflected_word.capitalize()

    return f"{prefix}{inflected_word}{suffix}"


def _fallback_inflect_person_name_word_to_dative(word: str) -> str:
    lower_word = word.lower()

    if lower_word.endswith(("овна", "евна", "ична", "инична")):
        return word[:-1] + "е"
    if lower_word.endswith(("ович", "евич")):
        return word + "у"
    if lower_word.endswith(("оглы", "улы", "кызы")):
        return word
    if lower_word.endswith(("ова", "ева", "ина", "ына")):
        return word[:-1] + "ой"
    if lower_word.endswith(("ская", "цкая")):
        return word[:-2] + "ой"
    if lower_word.endswith("ая"):
        return word[:-2] + "ой"
    if lower_word.endswith("ия"):
        return word[:-1] + "и"
    if lower_word.endswith("ья"):
        return word[:-1] + "е"
    if lower_word.endswith("я"):
        return word[:-1] + "е"
    if lower_word.endswith("а"):
        return word[:-1] + "е"
    if lower_word.endswith("й"):
        return word[:-1] + "ю"
    if lower_word.endswith("ь"):
        return word[:-1] + "ю"
    if lower_word and lower_word[-1] in "бвгджзклмнпрстфхцчшщ":
        return word + "у"

    return word


def _inflect_person_name_word_to_dative(word: str, gender_hint: Optional[str], expected_kind: Optional[str] = None) -> str:
    if not word:
        return word

    match = _TOKEN_PATTERN.match(word)
    if not match:
        return word

    prefix, core, suffix = match.groups()
    if not core or any(char.isdigit() for char in core):
        return word

    inflected_word = core
    lower_core = core.lower()
    prefer_fallback = gender_hint == "femn" and lower_core.endswith(
        (
            "ова",
            "ева",
            "ина",
            "ына",
            "ская",
            "цкая",
            "ая",
            "ия",
            "ья",
            "я",
            "а",
            "овна",
            "евна",
            "ична",
            "инична",
        )
    )

    if prefer_fallback:
        inflected_word = _fallback_inflect_person_name_word_to_dative(core)
    elif _MORPH is not None:
        parsed = _choose_best_person_name_parse(core, gender_hint, expected_kind)
        if parsed is not None:
            inflected = parsed.inflect({"datv"})
            if inflected is not None:
                inflected_word = inflected.word
            if inflected_word == core:
                inflected_word = _fallback_inflect_person_name_word_to_dative(core)
        else:
            inflected_word = _fallback_inflect_person_name_word_to_dative(core)
    else:
        inflected_word = _fallback_inflect_person_name_word_to_dative(core)

    if core[:1].isupper():
        inflected_word = inflected_word.capitalize()

    return f"{prefix}{inflected_word}{suffix}"


def _fallback_inflect_word_to_dative(word: str) -> str:
    lower_word = word.lower()

    if lower_word.endswith(("ова", "ева", "ина", "ына")):
        return word[:-1] + "ой"
    if lower_word.endswith(("ская", "цкая")):
        return word[:-2] + "ой"
    if lower_word.endswith("ая"):
        return word[:-2] + "ой"
    if lower_word.endswith("яя"):
        return word[:-2] + "ей"
    if lower_word.endswith(("ий", "ый", "ой")):
        return word[:-2] + "ому"
    if lower_word.endswith("ия"):
        return word[:-2] + "ии"
    if lower_word.endswith("ья"):
        return word[:-1] + "е"
    if lower_word.endswith("я"):
        return word[:-1] + "е"
    if lower_word.endswith("а"):
        return word[:-1] + "е"
    if lower_word.endswith("ь"):
        return word[:-1] + "ю"
    if lower_word.endswith("й"):
        return word[:-1] + "ю"
    if lower_word.endswith("ие"):
        return word[:-2] + "им"
    if lower_word.endswith("ые"):
        return word[:-2] + "ым"
    if lower_word and lower_word[-1] in "бвгджзклмнпрстфхцчшщ":
        return word + "у"

    return word


def _to_genitive_phrase(value: Any, base_name: Optional[str] = None) -> str:
    if value in (None, ""):
        return ""

    if _normalize_for_inflection_compare(value) in _INDECLINABLE_WORDS:
        return str(value)

    if _looks_like_full_org_name(base_name):
        ip_full_name_parts = _split_ip_full_name(value)
        if ip_full_name_parts is not None:
            org_prefix, person_name = ip_full_name_parts
            return (
                f"{_to_genitive_phrase(org_prefix)} "
                f"{_to_genitive_phrase(person_name, 'Руководитель')}"
            )

    words = str(value).split()
    if _looks_like_person_name(value, base_name):
        gender_hint = _guess_person_gender(words)
        expected_kinds = ("Surn", "Name", "Patr")
        return " ".join(
            _inflect_person_name_word_to_genitive(
                word,
                gender_hint,
                expected_kinds[index] if index < len(expected_kinds) else None,
            )
            for index, word in enumerate(words)
        )

    result = " ".join(_inflect_word_to_genitive(word) for word in words)
    return _preserve_indeclinable_tokens_from_source(str(value), result)


def _to_dative_phrase(value: Any, base_name: Optional[str] = None) -> str:
    if value in (None, ""):
        return ""

    if _normalize_for_inflection_compare(value) in _INDECLINABLE_WORDS:
        return str(value)

    if _looks_like_full_org_name(base_name):
        ip_full_name_parts = _split_ip_full_name(value)
        if ip_full_name_parts is not None:
            org_prefix, person_name = ip_full_name_parts
            return (
                f"{_to_dative_phrase(org_prefix)} "
                f"{_to_dative_phrase(person_name, 'Руководитель')}"
            )

    words = str(value).split()
    if _looks_like_person_name(value, base_name):
        gender_hint = _guess_person_gender(words)
        expected_kinds = ("Surn", "Name", "Patr")
        return " ".join(
            _inflect_person_name_word_to_dative(
                word,
                gender_hint,
                expected_kinds[index] if index < len(expected_kinds) else None,
            )
            for index, word in enumerate(words)
        )

    result = " ".join(_inflect_word_to_dative(word) for word in words)
    return _preserve_indeclinable_tokens_from_source(str(value), result)


def _preserve_indeclinable_tokens_from_source(source: str, result: str) -> str:
    source_words = source.split()
    result_words = result.split()
    if len(source_words) != len(result_words):
        return result

    normalized_indeclinables = set(_INDECLINABLE_WORDS) | {"хостес"}
    restored_words: list[str] = []
    for source_word, result_word in zip(source_words, result_words):
        normalized_source = _normalize_for_inflection_compare(source_word)
        normalized_result = _normalize_for_inflection_compare(result_word)
        if normalized_source in normalized_indeclinables:
            restored_words.append(source_word)
            continue
        if normalized_result in normalized_indeclinables:
            restored_words.append(source_word)
            continue
        restored_words.append(result_word)

    return " ".join(restored_words)


def inflect_to_genitive_phrase(value: Any, base_name: Optional[str] = None) -> str:
    return _to_genitive_phrase(value, base_name)


def inflect_to_dative_phrase(value: Any, base_name: Optional[str] = None) -> str:
    return _to_dative_phrase(value, base_name)


def _resolve_relative_month_value(key: str, replacements: dict[str, Any]) -> Optional[str]:
    stripped_key = key.strip()
    match = _RELATIVE_MONTHS_PATTERN.match(stripped_key.strip("<>"))
    if not match:
        return None

    base_name = match.group("base").strip()
    months = int(match.group("months"))

    base_value = None
    for candidate_key in _candidate_base_keys(base_name):
        if candidate_key in replacements:
            base_value = replacements[candidate_key]
            break

    base_date = _parse_date(base_value)
    if base_date is None:
        return None

    return _format_date(base_date + relativedelta(months=months))


def _resolve_genitive_case_value(key: str, replacements: dict[str, Any]) -> Optional[str]:
    stripped_key = key.strip()
    match = _GENITIVE_CASE_PATTERN.match(stripped_key.strip("<>"))
    if not match:
        return None

    base_name = match.group("base").strip()

    for candidate_base_name in _candidate_base_names(base_name):
        for candidate_key in _candidate_base_keys(candidate_base_name):
            if candidate_key in replacements:
                return _to_genitive_phrase(replacements[candidate_key], candidate_base_name)

    normalized_base_names = {_normalize_key_name(name) for name in _candidate_base_names(base_name)}
    for replacement_key, replacement_value in replacements.items():
        if _normalize_key_name(replacement_key) in normalized_base_names:
            return _to_genitive_phrase(replacement_value, base_name)

    return None


def _resolve_dative_case_value(key: str, replacements: dict[str, Any]) -> Optional[str]:
    stripped_key = key.strip()
    match = _DATIVE_CASE_PATTERN.match(stripped_key.strip("<>"))
    if not match:
        return None

    base_name = match.group("base").strip()

    for candidate_base_name in _candidate_base_names(base_name):
        for candidate_key in _candidate_base_keys(candidate_base_name):
            if candidate_key in replacements:
                return _to_dative_phrase(replacements[candidate_key], candidate_base_name)

    normalized_base_names = {_normalize_key_name(name) for name in _candidate_base_names(base_name)}
    for replacement_key, replacement_value in replacements.items():
        if _normalize_key_name(replacement_key) in normalized_base_names:
            return _to_dative_phrase(replacement_value, base_name)

    return None


def expand_relative_month_keys(
    replacements: dict[str, Any], candidate_keys: Optional[Iterable[str]] = None
) -> dict[str, Any]:
    expanded = dict(replacements)
    keys_to_process = list(candidate_keys or ())

    if not keys_to_process:
        keys_to_process = list(expanded.keys())

    for key in keys_to_process:
        if key in expanded and expanded[key] not in (None, ""):
            continue

        resolved_value = _resolve_relative_month_value(key, expanded)
        if resolved_value is None:
            resolved_value = _resolve_genitive_case_value(key, expanded)
        if resolved_value is None:
            resolved_value = _resolve_dative_case_value(key, expanded)

        if resolved_value is not None:
            expanded[key] = resolved_value

    return expanded


def replace_text_with_relative_dates(text: str, replacements: dict[str, Any]) -> str:
    if not text:
        return text

    candidate_keys = _ANGLE_BRACKET_PATTERN.findall(text)
    expanded = expand_relative_month_keys(replacements, candidate_keys)

    result = text
    for key in candidate_keys:
        if key in expanded:
            result = result.replace(key, str(expanded[key]))

    return result
