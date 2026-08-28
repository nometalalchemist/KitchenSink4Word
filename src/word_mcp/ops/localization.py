"""International language support: localized built-in style names + CJK-aware
word counting.

Word's built-in style IDs (``Heading1``, ``TOC1``, ``Caption`` ...) are the
same in every language version — only the DISPLAY names localize ("Heading 1"
becomes 제목 1 on a Korean install, Überschrift 1 on a German one). Any code
that matches styles by display name (COM ``Style.NameLocal``, or the w:name
element in styles.xml) therefore misfires on non-English installs. This module
holds the alias table and the matching helpers those sites route through.

Alias sources (verification status is marked per entry below):
- English:  canonical names, from Word itself.
- German / French: DocTools "List of Built-in Style Names — English, Danish,
  German, French" (Lene Fredborg, thedoctools.com, Word 2016) — VERIFIED.
- Japanese: antenna.co.jp built-in style reference — VERIFIED for the styles
  listed there.
- Korean:   제목 n / 목차 제목 / 캡션 / 인용 confirmed on an actual Korean
  Word install; 제목 n additionally corroborated by Microsoft ko-kr support
  pages. Remaining Korean names are the widely documented gallery names but
  are marked UNVERIFIED where no authoritative source was found.
- Spanish / Italian: heading, caption and quote names corroborated by
  Spanish/Italian Word documentation pages; the rest marked UNVERIFIED.
- Chinese (Simplified) / Portuguese: partially corroborated; entries that
  could not be confirmed against a source are marked UNVERIFIED.

An UNVERIFIED alias is a marked best-effort entry, never a silent guess: a
wrong alias only matters if that exact string is also a real style name for a
DIFFERENT style in some language, which none of these are known to be.

Matching is case-insensitive and whitespace-insensitive (all whitespace is
stripped before comparison), so "TOC Heading" == "tocheading" and
"참고 문헌" == "참고문헌".
"""

from __future__ import annotations

import re

# --------------------------------------------------------------- alias table
# canonical key -> localized display names (English always included).
# Languages: en, ko, de, fr, es, ja, zh-CN, pt, it.

_HEADING_WORD = {
    # per-language word used for "Heading {n}"; expanded below for n = 1..9
    "en": "Heading",        # Heading 1
    "ko": "제목",           # 제목 1        (verified: Korean install + MS ko-kr)
    "de": "Überschrift",    # Überschrift 1 (verified: DocTools)
    "fr": "Titre",          # Titre 1       (verified: DocTools)
    "es": "Título",         # Título 1      (verified: es Word docs)
    "ja": "見出し",         # 見出し 1      (verified: antenna.co.jp)
    "zh": "标题",           # 标题 1        (verified: zh Word docs)
    "pt": "Título",         # Título 1      # UNVERIFIED (same string as es)
    "it": "Titolo",         # Titolo 1      (verified: it Word docs)
}

STYLE_ALIASES: dict[str, set[str]] = {
    # Heading 1..9 entries are generated from _HEADING_WORD below.
    **{f"heading{n}": set() for n in range(1, 10)},
    "toc_heading": {
        "TOC Heading",
        "목차 제목",                       # ko (verified: Korean install)
        "Inhaltsverzeichnisüberschrift",   # de (verified: DocTools)
        "En-tête de table des matières",   # fr (verified: DocTools)
        "Título TDC",                      # es  # UNVERIFIED
        "目次の見出し",                    # ja (verified: antenna.co.jp)
        "TOC 标题",                        # zh  # UNVERIFIED
        "Título do Sumário",               # pt  # UNVERIFIED
        "Titolo sommario",                 # it  # UNVERIFIED
    },
    "caption": {
        "Caption",
        "캡션",             # ko (verified: Korean install)
        "Beschriftung",     # de (verified: DocTools)
        "Légende",          # fr (verified: DocTools)
        "Epígrafe",         # es (verified: es Word docs)
        "Descripción",      # es, newer builds  # UNVERIFIED
        "図表番号",         # ja (verified: antenna.co.jp)
        "题注",             # zh (verified: zh Word docs)
        "Legenda",          # pt  # UNVERIFIED
        "Didascalia",       # it (verified: it Word docs)
    },
    "quote": {
        "Quote",
        "인용",               # ko (verified: Korean install)
        "Anführungszeichen",  # de, Word 2016 era (verified: DocTools)
        "Zitat",              # de, newer builds  # UNVERIFIED
        "Citation",           # fr (verified: DocTools)
        "Cita",               # es (verified: es Word docs)
        "引用文",             # ja (verified: antenna.co.jp)
        "引用",               # zh  # UNVERIFIED (same string as ko — same key)
        "Citação",            # pt  # UNVERIFIED
        "Citazione",          # it (verified: it Word docs)
    },
    "block_text": {
        "Block Text",
        "블록 텍스트",       # ko  # UNVERIFIED
        "Blocktext",         # de (verified: DocTools)
        "Normal centré",     # fr (verified: DocTools — yes, really)
        "ブロック テキスト", # ja  # UNVERIFIED (antenna lists "ブロック")
        "ブロック",          # ja  # UNVERIFIED
    },
    "footnote_text": {
        "Footnote Text",
        "각주 텍스트",              # ko  # UNVERIFIED
        "Fußnotentext",             # de (verified: DocTools)
        "Note de bas de page",      # fr (verified: DocTools)
        "Texto nota pie",           # es  # UNVERIFIED
        "脚注文字列",               # ja (verified: antenna.co.jp)
        "脚注文本",                 # zh  # UNVERIFIED
        "Texto de nota de rodapé",  # pt  # UNVERIFIED
        "Testo nota a piè di pagina",  # it  # UNVERIFIED
    },
    "title": {
        "Title",
        "제목",     # ko (verified: Korean install; collides with the heading
        #             word alone — exact match only, "제목 1" is heading1)
        "Titel",    # de (verified: DocTools)
        "Titre",    # fr (verified: DocTools)
        "Título",   # es/pt (verified: es Word docs)
        "表題",     # ja (verified: antenna.co.jp)
        "标题",     # zh (verified: zh Word docs)
        "Titolo",   # it (verified: it Word docs)
    },
    "subtitle": {
        "Subtitle",
        "부제",         # ko  # UNVERIFIED
        "Untertitel",   # de (verified: DocTools)
        "Sous-titre",   # fr (verified: DocTools)
        "Subtítulo",    # es/pt  # UNVERIFIED
        "副題",         # ja (verified: antenna.co.jp)
        "副标题",       # zh  # UNVERIFIED
        "Sottotitolo",  # it  # UNVERIFIED
    },
    "normal": {
        "Normal",     # en/fr/es/pt (fr/es verified: DocTools & es Word docs)
        "표준",       # ko  # UNVERIFIED
        "Standard",   # de (verified: DocTools)
        "標準",       # ja (verified: antenna.co.jp)
        "正文",       # zh  # UNVERIFIED
        "Normale",    # it (verified: it Word docs)
    },
    "list_paragraph": {
        "List Paragraph",
        "목록 단락",           # ko  # UNVERIFIED
        "Listenabsatz",        # de (verified: DocTools)
        "Paragraphe de liste", # fr (verified: DocTools)
        "Párrafo de lista",    # es  # UNVERIFIED
        "リスト段落",          # ja (verified: antenna.co.jp)
        "列表段落",            # zh  # UNVERIFIED
        "Parágrafo da Lista",  # pt  # UNVERIFIED
        "Paragrafo elenco",    # it  # UNVERIFIED
    },
    # TOC 1..3 generated below from per-language patterns.
    **{f"toc{n}": set() for n in range(1, 4)},
}

# Expand Heading 1..9 (space and no-space number forms both normalize the
# same, so one spelling per language suffices).
for _n in range(1, 10):
    STYLE_ALIASES[f"heading{_n}"] = {
        f"{word} {_n}" for word in _HEADING_WORD.values()
    }

_TOC_WORD = {
    "en": "TOC",           # TOC 1        (verified)
    "ko": "목차",          # 목차 1       # UNVERIFIED
    "de": "Verzeichnis",   # Verzeichnis 1 (verified: DocTools)
    "fr": "TM",            # TM 1         (verified: DocTools)
    "es": "TDC",           # TDC 1        # UNVERIFIED
    "ja": "目次",          # 目次 1       (verified: antenna.co.jp)
    "zh": "目录",          # 目录 1       (verified: zh Word docs)
    "pt": "Sumário",       # Sumário 1    # UNVERIFIED
    "it": "Sommario",      # Sommario 1   # UNVERIFIED
}
for _n in range(1, 4):
    STYLE_ALIASES[f"toc{_n}"] = {f"{word} {_n}" for word in _TOC_WORD.values()}


# ------------------------------------------------------------------ matching


def _norm(name: str | None) -> str:
    """Casefold and strip ALL whitespace: matching is case- and
    whitespace-insensitive across every language."""
    return re.sub(r"\s+", "", (name or "")).casefold()


_CANONICAL_BY_NORM: dict[str, str] = {}
for _key, _names in STYLE_ALIASES.items():
    for _name in _names:
        _prev = _CANONICAL_BY_NORM.setdefault(_norm(_name), _key)
        if _prev != _key:  # pragma: no cover — table integrity guard
            raise AssertionError(
                f"alias {_name!r} maps to both {_prev!r} and {_key!r}"
            )


def localized_names(key: str) -> set[str]:
    """All known display names (every language) for a canonical style key."""
    if key not in STYLE_ALIASES:
        raise KeyError(
            f"unknown canonical style key {key!r}; "
            f"known: {sorted(STYLE_ALIASES)}"
        )
    return set(STYLE_ALIASES[key])


_NORMS_BY_KEY: dict[str, set[str]] = {
    key: {_norm(n) for n in names} for key, names in STYLE_ALIASES.items()
}


def style_name_matches(name: str | None, canonical_key: str) -> bool:
    """True when `name` is the display name of `canonical_key` in ANY
    supported language (case- and whitespace-insensitive)."""
    key_norms = _NORMS_BY_KEY.get(canonical_key)
    if key_norms is None:
        raise KeyError(
            f"unknown canonical style key {canonical_key!r}; "
            f"known: {sorted(STYLE_ALIASES)}"
        )
    return _norm(name) in key_norms


def canonical_for_name(name: str | None) -> str | None:
    """The canonical style key a display name belongs to, or None."""
    return _CANONICAL_BY_NORM.get(_norm(name))


# -------------------------------------------------- CJK-aware word counting
#
# Space-tokenized counting (\S+) badly under-counts Japanese and Chinese,
# which do not separate words with spaces. The academic convention for zh/ja
# is a CHARACTER count, so each CJK ideograph or kana character counts as one
# "word". Korean hangul DOES use spaces between words, so hangul text keeps
# ordinary token counting.
#
# Counted one-word-per-character:
#   U+3040–U+309F   Hiragana
#   U+30A0–U+30FF   Katakana
#   U+31F0–U+31FF   Katakana Phonetic Extensions
#   U+3400–U+4DBF   CJK Unified Ideographs Extension A
#   U+4E00–U+9FFF   CJK Unified Ideographs
#   U+F900–U+FAFF   CJK Compatibility Ideographs
#   U+FF66–U+FF9D   Halfwidth Katakana
#   U+20000–U+2FFFF CJK Unified Ideographs Extensions B–F + Compatibility
#                   Ideographs Supplement
#
# Stripped but NOT counted (punctuation is not a word in either convention):
#   U+3000–U+303F   CJK Symbols and Punctuation (。、「」 …)
#   U+FE30–U+FE4F   CJK Compatibility Forms
#   U+FF01–U+FF0F, U+FF1A–U+FF20, U+FF3B–U+FF40, U+FF5B–U+FF65
#                   Fullwidth punctuation (fullwidth digits/letters are NOT
#                   stripped — they remain part of ordinary tokens)
#
# Hangul (U+AC00–U+D7AF etc.) is deliberately NOT in the per-character set:
# Korean separates words with spaces, so hangul flows through token counting.

_CJK_CHAR = re.compile(
    "["
    "぀-ヿ"  # hiragana + katakana (one contiguous run of blocks)
    "ㇰ-ㇿ"  # katakana phonetic extensions
    "㐀-䶿"  # CJK unified ideographs extension A
    "一-鿿"  # CJK unified ideographs
    "豈-﫿"  # CJK compatibility ideographs
    "ｦ-ﾝ"  # halfwidth katakana
    "𠀀-𯿿"  # CJK ext B-F + compat ideographs supplement
    "]"
)
_CJK_PUNCT = re.compile(
    "["
    "　-〿"  # CJK symbols and punctuation
    "︰-﹏"  # CJK compatibility forms
    "！-／"  # fullwidth ! " # $ % & ' ( ) * + , - . /
    "：-＠"  # fullwidth : ; < = > ? @
    "［-｀"  # fullwidth [ \ ] ^ _ `
    "｛-･"  # fullwidth { | } ~ + halfwidth CJK punctuation
    "]"
)
_TOKEN = re.compile(r"\S+")


def cjk_aware_word_count(text: str) -> dict:
    """Word count that handles CJK text correctly.

    Returns {"words", "cjk_chars", "counting"} where words = space-delimited
    tokens (after removing CJK characters) + one word per CJK character;
    cjk_chars is the CJK character count alone; counting is "spaces" (no CJK
    present — identical to the ``\\S+`` token count), "cjk" (only CJK), or
    "mixed". Pure-space text (English, Korean, ...) returns exactly the same
    number as plain ``\\S+`` tokenization — the invariance the rest of the
    stats/journalcount code relies on."""
    text = text or ""
    cjk_chars = len(_CJK_CHAR.findall(text))
    if cjk_chars:
        residue = _CJK_PUNCT.sub(" ", _CJK_CHAR.sub(" ", text))
        space_tokens = len(_TOKEN.findall(residue))
    else:
        # No CJK: count the raw text so the \S+ invariance is exact (CJK
        # punctuation alone, with no CJK words, is left to ordinary
        # tokenization rather than silently dropped).
        space_tokens = len(_TOKEN.findall(text))
    if cjk_chars and space_tokens:
        counting = "mixed"
    elif cjk_chars:
        counting = "cjk"
    else:
        counting = "spaces"
    return {
        "words": space_tokens + cjk_chars,
        "cjk_chars": cjk_chars,
        "counting": counting,
    }
