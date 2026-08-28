"""International language support: localized built-in style names + CJK-aware
word counting.

Word's built-in style IDs (``Heading1``, ``TOC1``, ``Caption`` ...) are the
same in every language version — only the DISPLAY names localize ("Heading 1"
becomes 제목 1 on a Korean install, Überschrift 1 on a German one). Any code
that matches styles by display name (COM ``Style.NameLocal``, or the w:name
element in styles.xml) therefore misfires on non-English installs. This module
holds the alias table and the matching helpers those sites route through.

Alias sources (verification status is marked per entry below; the formerly
UNVERIFIED entries were researched and corrected in the v1.6 sweep —
2026-08-28):
- English:  canonical names, from Word itself.
- German / French: DocTools "List of Built-in Style Names — English, Danish,
  German, French" (Lene Fredborg, thedoctools.com, Word 2016) — VERIFIED.
- Japanese: antenna.co.jp built-in style reference (incl. the Word 2021
  full paragraph-style list) — VERIFIED for the styles listed there.
- Korean:   제목 n / 목차 제목 / 캡션 / 인용 confirmed on an actual Korean
  Word install; further names confirmed from real Korean-Word docx
  styles.xml dumps (linked "... Char" styles) and Microsoft ko-kr Q&A.
- Spanish / Italian / Chinese (Simplified) / Portuguese: confirmed against
  authentic localized-Word output found in public repos — mso-style-name /
  mso-style-link strings in Word-exported HTML/CSS, styles.xml dumps, and
  localized styleIds (Word derives styleIds from the display name, e.g.
  ``TtuloTDC`` <- "Título TDC") — plus vendor/MVP documentation.

CAUTION for future maintenance: localized support.microsoft.com pages are
machine-translated and have shown provably wrong style names (the es-es
captions page calls the Caption style "Título", colliding with Heading 1);
never source aliases from them. Aliases that could not be confirmed from at
least one credible source were REMOVED, not kept: a wrong alias silently
mis-detects styles, which is worse than no alias (removals are commented in
the table below).

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
    "pt": "Título",         # Título 1      (verified: MS support pt-BR
    #                         "Numerar os títulos" + real pt-BR Word export
    #                         with "Título 1 Char".."Título 9" — same string
    #                         as Spanish, and that is correct)
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
        "Título TDC",                      # es (verified: real es docx dump,
        #                                    styleId TtuloTDC)
        "目次の見出し",                    # ja (verified: antenna.co.jp)
        "TOC 标题",                        # zh (verified: zh Word UI reports
        #                                    on Baidu Zhidao / CSDN)
        "Cabeçalho do Sumário",            # pt-BR (verified: real pt-BR docx
        #                                    styleId CabealhodoSumrio; docx4j
        #                                    TocSdtUtils). The former guess
        #                                    "Título do Sumário" was WRONG
        #                                    and was removed (v1.6 sweep).
        "Titolo sommario",                 # it (verified: Italian styleId
        #                                    contract Titolosommario, utf8dok)
    },
    "caption": {
        "Caption",
        "캡션",             # ko (verified: Korean install)
        "Beschriftung",     # de (verified: DocTools)
        "Légende",          # fr (verified: DocTools)
        "Epígrafe",         # es, older builds (verified: es Word docs +
        #                     wordexperto.com)
        "Descripción",      # es, newer builds (verified: wordexperto.com
        #                     [Spanish Word MVP] + real es docx styleId
        #                     Descripcin)
        "図表番号",         # ja (verified: antenna.co.jp)
        "题注",             # zh (verified: zh Word docs)
        "Legenda",          # pt (verified: MS support pt-BR captions page +
        #                     real pt-BR docx styleId Legenda; also pt-PT
        #                     per sfm.pt)
        "Didascalia",       # it (verified: it Word docs)
    },
    "quote": {
        "Quote",
        "인용",               # ko (verified: Korean install)
        "Anführungszeichen",  # de, Word 2010/2016 era (verified: DocTools)
        "Zitat",              # de, newer builds (verified: real German Word
        #                       RTF char style "Zitat Zchn" in LibreOffice
        #                       core test data; German Word tutorials)
        "Citation",           # fr (verified: DocTools)
        "Cita",               # es (verified: es Word docs)
        "引用文",             # ja (verified: antenna.co.jp)
        "引用",               # zh (verified: real zh Word export,
        #                       mso-style-name "引用 字符"; same string as
        #                       ko — same canonical key, no collision)
        "Citação",            # pt (verified: real pt-BR Word export,
        #                       mso-style-link "Citação Char")
        "Citazione",          # it (verified: it Word docs)
    },
    "block_text": {
        "Block Text",
        # ko "블록 텍스트" REMOVED (v1.6 sweep): no credible source found,
        # and the Japanese precedent (Word shortens this style to just
        # "ブロック") shows the literal-translation guess cannot be trusted.
        # Re-add only after confirmation on a Korean Word install.
        "Blocktext",         # de (verified: DocTools)
        "Normal centré",     # fr (verified: DocTools — yes, really)
        # ja "ブロック テキスト" REMOVED (v1.6 sweep): WRONG — the Antenna
        # House Word 2021 built-in style list has plain "ブロック".
        "ブロック",          # ja (verified: antenna.co.jp Word 2021 list)
    },
    "footnote_text": {
        "Footnote Text",
        "각주 텍스트",              # ko (verified: real Korean docx
        #                             styles.xml, "각주 텍스트 Char")
        "Fußnotentext",             # de (verified: DocTools)
        "Note de bas de page",      # fr (verified: DocTools)
        "Texto nota pie",           # es (verified: real es Word export CSS
        #                             "Texto nota pie Car" + ENI Word manual)
        "脚注文字列",               # ja (verified: antenna.co.jp)
        "脚注文本",                 # zh (verified: real zh Word export CSS
        #                             "脚注文本 Char")
        "Texto de nota de rodapé",  # pt (verified: real pt-BR Word export,
        #                             mso-style-link)
        "Testo nota piè di pagina", # it (verified: MS Answers it-it walking
        #                             through Gestisci Stili). The former
        #                             guess "Testo nota A piè di pagina"
        #                             (extra "a") was WRONG and was removed
        #                             (v1.6 sweep).
    },
    "title": {
        "Title",
        "제목",     # ko (verified: Korean install; collides with the heading
        #             word alone — exact match only, "제목 1" is heading1)
        "Titel",    # de (verified: DocTools)
        "Titre",    # fr (verified: DocTools)
        "Título",   # pt, and es on some builds (the same collide-with-
        #             heading pattern as Korean 제목 — exact match only,
        #             "Título 1" is heading1)
        "Puesto",   # es (verified: real es Word export CSS/styleIds — the
        #             famous "Title"-as-job-title translation; found while
        #             validating the table, v1.6 sweep)
        "表題",     # ja (verified: antenna.co.jp)
        "标题",     # zh (verified: zh Word docs)
        "Titolo",   # it (verified: it Word docs)
    },
    "subtitle": {
        "Subtitle",
        "부제",         # ko (verified: real Korean docx styles.xml,
        #                 "부제 Char")
        "Untertitel",   # de (verified: DocTools)
        "Sous-titre",   # fr (verified: DocTools)
        "Subtítulo",    # es/pt (verified: real es Word export "Subtítulo
        #                 Car" + pt-BR docx styleId Subttulo)
        "副題",         # ja (verified: antenna.co.jp)
        "副标题",       # zh (verified: real zh Word export CSS
        #                 "副标题 Char")
        "Sottotitolo",  # it (verified: Italian Word export CSS
        #                 "Sottotitolo Carattere")
    },
    "normal": {
        "Normal",     # en/fr/es/pt (fr/es verified: DocTools & es Word docs)
        "표준",       # ko (verified: Korean MS Q&A + Korean Word books
        #               describing 표준 as the base style)
        "Standard",   # de (verified: DocTools)
        "標準",       # ja (verified: antenna.co.jp)
        "正文",       # zh (verified: zh style guides — 样式基准 "正文";
        #               real exports carry derived 正文文本 built-ins)
        "Normale",    # it (verified: it Word docs)
    },
    "list_paragraph": {
        "List Paragraph",
        # ko "목록 단락" REMOVED (v1.6 sweep): only third-party app resource
        # files use it; no Word-origin evidence found. Re-add only after
        # confirmation on a Korean Word install.
        "Listenabsatz",        # de (verified: DocTools)
        "Paragraphe de liste", # fr (verified: DocTools)
        "Párrafo de lista",    # es (verified: real es Word export
        #                        mso-style-parent + styleId Prrafodelista)
        "リスト段落",          # ja (verified: antenna.co.jp)
        "列出段落",            # zh (verified: real zh Word 2016 exports —
        #                        collision names "列出段落1/2" and CKEditor
        #                        paste-from-office corpus; this is the MS
        #                        Word UI string)
        "列表段落",            # zh, secondary (blogs + WPS-suspect exports
        #                        only; kept as a harmless extra alias since
        #                        it collides with nothing)
        "Parágrafo da Lista",  # pt (verified: CKEditor real Word 2016 paste
        #                        corpus, "Parágrafo da Lista1")
        "Paragrafo elenco",    # it (verified: Italian styleId contract
        #                        Paragrafoelenco, utf8dok)
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
    "ko": "목차",          # 목차 1       (verified: MS Answers ko-kr thread
    #                        on modifying 목차 1..4 level styles)
    "de": "Verzeichnis",   # Verzeichnis 1 (verified: DocTools)
    "fr": "TM",            # TM 1         (verified: DocTools)
    "es": "TDC",           # TDC 1        (verified: real es docx styleIds
    #                        TDC1/2/3 + aulaClic / libguides)
    "ja": "目次",          # 目次 1       (verified: antenna.co.jp)
    "zh": "目录",          # 目录 1       (verified: zh Word docs)
    "pt": "Sumário",       # Sumário 1    (verified pt-BR: real docx styleId
    #                        Sumrio1 = "toc 1", docx4j TocStyles; pt-PT
    #                        likely uses "Índice 1" but that is UNCONFIRMED
    #                        and deliberately NOT added)
    "it": "Sommario",      # Sommario 1   (verified: Italian TOC guides +
    #                        utf8dok Sommario1..9)
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
