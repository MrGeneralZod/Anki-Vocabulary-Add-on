import html
import hashlib
import json
import os
import re
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser
from typing import Any, Dict, List, Optional, Tuple

from anki.notes import Note
from aqt import mw
from aqt import gui_hooks
from aqt.browser import Browser
from aqt.editor import Editor
from aqt.qt import QAction, QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLineEdit, QMenu, Qt, QVBoxLayout, QWidget
from aqt.utils import askUser, showInfo, tooltip


MENU_LABEL = "Enrich Vocabulary Fields"
LAST_API_ERROR = ""
DEFAULT_CONFIG = {
    "source_field": "Word",
    "field_map": {
        "ipa": "IPA",
        "definition": "Definition",
        "examples": "Examples",
        "synonyms": "Synonyms",
        "antonyms": "Antonyms",
        "image": "Image",
    },
    "separator": "; ",
    "overwrite_existing": False,
    "max_definitions": 3,
    "max_examples": 3,
    "data_source": "custom",
    "api_keys": {
        "wordnik": "",
        "merriam_webster": "",
        "merriam_collegiate": "",
        "merriam_sd3": "",
    },
    "cambridge_usage_tags_enabled": True,
    "cambridge_usage_tag_map": {},
}


def _read_config() -> Dict[str, Any]:
    cfg = mw.addonManager.getConfig(__name__) or {}
    merged = DEFAULT_CONFIG.copy()
    merged.update(cfg)
    field_map = DEFAULT_CONFIG["field_map"].copy()
    field_map.update(merged.get("field_map", {}))
    merged["field_map"] = field_map
    api_keys = DEFAULT_CONFIG["api_keys"].copy()
    api_keys.update(merged.get("api_keys", {}))
    merged["api_keys"] = api_keys
    usage_tag_map = DEFAULT_CONFIG["cambridge_usage_tag_map"].copy()
    usage_tag_map.update(merged.get("cambridge_usage_tag_map", {}))
    merged["cambridge_usage_tag_map"] = usage_tag_map
    # Safety for older/broken config states where source_field was overwritten by data source id.
    if merged.get("source_field") in ("custom", "dictionaryapi", "wordnik", "merriam_webster", "cambridge"):
        merged["source_field"] = DEFAULT_CONFIG["source_field"]
    return merged


def _clean(value: str) -> str:
    return " ".join(value.split()).strip()


def _plain_text(value: str) -> str:
    unescaped = html.unescape(value or "")
    no_tags = re.sub(r"<[^>]+>", " ", unescaped)
    return _clean(no_tags)


def _resolve_note_field_name(note: Note, configured_name: str) -> Optional[str]:
    if configured_name in note:
        return configured_name
    target = _clean(configured_name).lower()
    for actual_name in note.keys():
        if _clean(actual_name).lower() == target:
            return actual_name
    # Friendly fallback names for common note models.
    fallback_candidates = ("learnable", "word", "expression", "front", "term")
    note_fields_normalized = { _clean(name).lower(): name for name in note.keys() }
    for candidate in fallback_candidates:
        if candidate in note_fields_normalized:
            return note_fields_normalized[candidate]
    return None


def _auto_heal_source_field(cfg: Dict[str, Any], note: Note) -> str:
    configured = cfg.get("source_field", "")
    if configured in note:
        return configured
    normalized = {_clean(name).lower(): name for name in note.keys()}
    if _clean(configured).lower() in normalized:
        return normalized[_clean(configured).lower()]
    # If stale default is kept, prefer Learnable when present.
    if _clean(configured).lower() == "word" and "learnable" in normalized:
        healed = normalized["learnable"]
        cfg["source_field"] = healed
        mw.addonManager.writeConfig(__name__, cfg)
        return healed
    return configured


def _find_first_nonempty_field(note: Note) -> Optional[str]:
    for field_name in note.keys():
        value = _plain_text(note[field_name])
        if value:
            return field_name
    return None


def _request_dictionary_data(word: str) -> Optional[List[Dict[str, Any]]]:
    url = "https://api.dictionaryapi.dev/api/v2/entries/en/" + urllib.parse.quote(word)
    req = urllib.request.Request(url, headers={"User-Agent": "AnkiAddonVocabularyEnricher/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception:
        return None


def _cambridge_dictionary_url_slug(word: str) -> str:
    """Build the path segment Cambridge uses for multi-word entries (e.g. break-out).

    A literal space (break%20out) redirects to the wrong article (e.g. break-out-in-a-cold-sweat).
    Hyphenated lowercase slugs match the canonical dictionary URLs.
    """
    s = _clean(word).lower().replace("-", " ")
    tokens = re.findall(r"[a-z0-9']+", s)
    if tokens:
        return "-".join(tokens)
    return re.sub(r"\s+", "-", _clean(word).lower())


def _request_cambridge_html(word: str) -> Optional[str]:
    encoded = urllib.parse.quote(_cambridge_dictionary_url_slug(word), safe="-")
    urls = [
        "https://dictionary.cambridge.org/dictionary/english/" + encoded,
        "https://dictionary.cambridge.org/us/dictionary/english/" + encoded,
    ]
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "identity",
        "Referer": "https://dictionary.cambridge.org/",
    }
    for url in urls:
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=12) as response:
                body = response.read().decode("utf-8", errors="ignore")
        except Exception:
            continue
        if not body:
            continue
        lower = body.lower()
        # Skip anti-bot/login stubs and keep real dictionary pages.
        if ("verify you are human" in lower) or ("cf-chl" in lower):
            continue
        if ("ddef_d" in lower) or ("def-block" in lower) or ("meaning of " in lower):
            return body
    return None


_CAMBRIDGE_THES_BUTTON_SYN = re.compile(
    r'<span[^>]*\bclass=["\'][^"\']*\bdthesButton\b[^"\']*\bsynonym\b[^"\']*["\'][^>]*>(.*?)</span>',
    re.IGNORECASE | re.DOTALL,
)
_CAMBRIDGE_THES_BUTTON_ANT = re.compile(
    r'<span[^>]*\bclass=["\'][^"\']*\bdthesButton\b[^"\']*\bopposite\b[^"\']*["\'][^>]*>(.*?)</span>',
    re.IGNORECASE | re.DOTALL,
)
# Cambridge appends a second block ("word | American Thesaurus"); skip it so lists stay short.
_CAMBRIDGE_THES_US_SECTION_HEAD = " | American Thesaurus</h2>"
_CAMBRIDGE_THES_MAX_WORDS = 3


def _request_cambridge_thesaurus_html(word: str) -> Optional[str]:
    """Fetch Cambridge Thesaurus HTML (dedicated /thesaurus/… page, not dictionary xref blocks)."""
    encoded = urllib.parse.quote(_cambridge_dictionary_url_slug(word), safe="-")
    urls = [
        "https://dictionary.cambridge.org/thesaurus/" + encoded,
        "https://dictionary.cambridge.org/us/thesaurus/" + encoded,
    ]
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "identity",
        "Referer": "https://dictionary.cambridge.org/",
    }
    for url in urls:
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=12) as response:
                body = response.read().decode("utf-8", errors="ignore")
        except Exception:
            continue
        if not body:
            continue
        lower = body.lower()
        if ("verify you are human" in lower) or ("cf-chl" in lower):
            continue
        if "dthesbutton" in lower:
            return body
    return None


def _extract_cambridge_thesaurus_synonyms_antonyms(html_text: str) -> Tuple[List[str], List[str]]:
    """Parse synonym / antonym headwords from a Cambridge thesaurus page."""
    split_at = html_text.find(_CAMBRIDGE_THES_US_SECTION_HEAD)
    if split_at >= 0:
        html_text = html_text[:split_at]

    def _button_words(pattern) -> List[str]:
        out: List[str] = []
        for m in pattern.finditer(html_text):
            raw = re.sub(r"<[^>]+>", " ", m.group(1))
            text = _clean(html.unescape(raw))
            if text:
                out.append(text)
        return out

    def _unique(items: List[str]) -> List[str]:
        seen: set[str] = set()
        uniq: List[str] = []
        for item in items:
            key = item.lower()
            if key not in seen:
                seen.add(key)
                uniq.append(item)
        return uniq

    syn = _unique(_button_words(_CAMBRIDGE_THES_BUTTON_SYN))[:_CAMBRIDGE_THES_MAX_WORDS]
    ant = _unique(_button_words(_CAMBRIDGE_THES_BUTTON_ANT))[:_CAMBRIDGE_THES_MAX_WORDS]
    return syn, ant


def _request_wordnik_data(word: str, api_key: str) -> Optional[Dict[str, Any]]:
    if not api_key:
        return None
    base = "https://api.wordnik.com/v4/word.json/" + urllib.parse.quote(word)
    definitions_url = (
        f"{base}/definitions?limit=5&includeRelated=false&useCanonical=true&api_key={urllib.parse.quote(api_key)}"
    )
    examples_url = f"{base}/examples?limit=5&api_key={urllib.parse.quote(api_key)}"
    related_url = f"{base}/relatedWords?useCanonical=true&limitPerRelationshipType=10&api_key={urllib.parse.quote(api_key)}"
    try:
        req_defs = urllib.request.Request(definitions_url, headers={"User-Agent": "AnkiAddonVocabularyEnricher/1.0"})
        req_examples = urllib.request.Request(examples_url, headers={"User-Agent": "AnkiAddonVocabularyEnricher/1.0"})
        req_related = urllib.request.Request(related_url, headers={"User-Agent": "AnkiAddonVocabularyEnricher/1.0"})
        with urllib.request.urlopen(req_defs, timeout=10) as r_defs:
            definitions = json.loads(r_defs.read().decode("utf-8"))
        with urllib.request.urlopen(req_examples, timeout=10) as r_examples:
            examples = json.loads(r_examples.read().decode("utf-8"))
        with urllib.request.urlopen(req_related, timeout=10) as r_related:
            related = json.loads(r_related.read().decode("utf-8"))
        return {"definitions": definitions, "examples": examples, "related": related}
    except Exception:
        return None


def _request_merriam_data(word: str, api_keys: Dict[str, str]) -> Optional[List[Dict[str, Any]]]:
    global LAST_API_ERROR
    endpoint_keys = (
        ("collegiate", _clean(api_keys.get("merriam_collegiate", ""))),
        ("sd3", _clean(api_keys.get("merriam_sd3", ""))),
        ("learners", _clean(api_keys.get("merriam_webster", ""))),
        ("collegiate", _clean(api_keys.get("merriam_webster", ""))),  # backward compatibility
    )
    if not any(k for _endpoint, k in endpoint_keys):
        LAST_API_ERROR = "Merriam API key is empty."
        return None
    last_error = ""
    for endpoint, api_key in endpoint_keys:
        if not api_key:
            continue
        url = (
            f"https://www.dictionaryapi.com/api/v3/references/{endpoint}/json/"
            + urllib.parse.quote(word)
            + "?key="
            + urllib.parse.quote(api_key)
        )
        req = urllib.request.Request(url, headers={"User-Agent": "AnkiAddonVocabularyEnricher/1.0"})
        try:
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode("utf-8"))
            if isinstance(data, list) and data and isinstance(data[0], dict):
                LAST_API_ERROR = ""
                return data
            if isinstance(data, list) and data and isinstance(data[0], str):
                last_error = f"Merriam {endpoint}: word not found (suggestions returned)."
                continue
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                last_error = f"Merriam {endpoint}: API key rejected (HTTP {e.code})."
                break
            last_error = f"Merriam {endpoint}: HTTP {e.code}."
        except Exception as e:
            last_error = f"Merriam {endpoint}: {e.__class__.__name__}."
    LAST_API_ERROR = last_error or "Merriam returned no usable entries."
    return None


def _request_datamuse_data(word: str) -> Dict[str, List[str]]:
    syn_url = "https://api.datamuse.com/words?ml=" + urllib.parse.quote(word) + "&max=20"
    ant_url = "https://api.datamuse.com/words?rel_ant=" + urllib.parse.quote(word) + "&max=20"
    out = {"synonyms": [], "antonyms": []}
    try:
        syn_req = urllib.request.Request(syn_url, headers={"User-Agent": "AnkiAddonVocabularyEnricher/1.0"})
        with urllib.request.urlopen(syn_req, timeout=10) as r_syn:
            syn_data = json.loads(r_syn.read().decode("utf-8"))
        out["synonyms"] = [_clean(item.get("word", "")) for item in syn_data if _clean(item.get("word", ""))]
    except Exception:
        pass
    try:
        ant_req = urllib.request.Request(ant_url, headers={"User-Agent": "AnkiAddonVocabularyEnricher/1.0"})
        with urllib.request.urlopen(ant_req, timeout=10) as r_ant:
            ant_data = json.loads(r_ant.read().decode("utf-8"))
        out["antonyms"] = [_clean(item.get("word", "")) for item in ant_data if _clean(item.get("word", ""))]
    except Exception:
        pass
    return out


def _extract_details(
    payload: List[Dict[str, Any]],
    word: str,
    separator: str,
    max_definitions: int,
    max_examples: int,
    selected_definition: Optional[str] = None,
) -> Dict[str, str]:
    ipa = ""
    definitions: List[str] = []
    examples: List[str] = []
    synonyms: List[str] = []
    antonyms: List[str] = []

    for entry in payload:
        if not ipa:
            ipa = _clean(entry.get("phonetic", ""))
            if not ipa:
                for p in entry.get("phonetics", []):
                    text = _clean(p.get("text", ""))
                    if text:
                        ipa = text
                        break

        for meaning in entry.get("meanings", []):
            synonyms.extend(_clean(item) for item in meaning.get("synonyms", []) if _clean(item))
            antonyms.extend(_clean(item) for item in meaning.get("antonyms", []) if _clean(item))
            for d in meaning.get("definitions", []):
                definition = _clean(d.get("definition", ""))
                if definition:
                    definitions.append(definition)
                example = _clean(d.get("example", ""))
                if example:
                    examples.append(example)
                synonyms.extend(_clean(item) for item in d.get("synonyms", []) if _clean(item))
                antonyms.extend(_clean(item) for item in d.get("antonyms", []) if _clean(item))

    def unique(items: List[str]) -> List[str]:
        out: List[str] = []
        seen = set()
        for item in items:
            if item not in seen:
                out.append(item)
                seen.add(item)
        return out

    uniq_defs_all = unique(definitions)
    if selected_definition:
        uniq_defs = [selected_definition]
    else:
        uniq_defs = uniq_defs_all[:1]
    uniq_examples = unique(examples)[:max_examples]
    uniq_syn = unique(synonyms)
    uniq_ant = unique(antonyms)
    examples_html = ""
    if uniq_examples:
        escaped_word = re.escape(word)
        pattern = re.compile(escaped_word, flags=re.IGNORECASE) if escaped_word else None
        rendered: List[str] = []
        for example in uniq_examples:
            safe_example = html.escape(example)
            if pattern:
                safe_example = pattern.sub(lambda m: f"<u>{html.escape(m.group(0))}</u>", safe_example)
            rendered.append(f"<li>{safe_example}</li>")
        examples_html = "<ul>" + "".join(rendered) + "</ul>"

    return {
        "ipa": ipa,
        "definition": separator.join(uniq_defs),
        "examples": examples_html,
        "synonyms": separator.join(uniq_syn),
        "antonyms": separator.join(uniq_ant),
    }


def _extract_from_wordnik(payload: Dict[str, Any], word: str, separator: str, max_examples: int) -> Dict[str, str]:
    definitions = []
    for item in payload.get("definitions", []):
        text = _clean(item.get("text", ""))
        if text:
            definitions.append(text)

    examples = []
    examples_data = payload.get("examples", {})
    for item in examples_data.get("examples", []):
        text = _clean(item.get("text", ""))
        if text:
            examples.append(text)

    syn = []
    ant = []
    for rel in payload.get("related", []):
        rel_type = _clean(rel.get("relationshipType", ""))
        words = [_clean(w) for w in rel.get("words", []) if _clean(w)]
        if rel_type in ("synonym", "same-context"):
            syn.extend(words)
        if rel_type in ("antonym",):
            ant.extend(words)

    return _extract_details(
        [{"phonetic": "", "meanings": []}],
        word,
        separator,
        1,
        max_examples,
        selected_definition=(definitions[0] if definitions else None),
    ) | {
        "examples": _extract_details(
            [{"phonetic": "", "meanings": [{"definitions": [{"example": e} for e in examples[:max_examples]]}]}],
            word,
            separator,
            1,
            max_examples,
        )["examples"],
        "synonyms": separator.join(syn),
        "antonyms": separator.join(ant),
    }


def _extract_from_merriam(payload: List[Dict[str, Any]], separator: str) -> Dict[str, str]:
    ipa = ""
    definition = ""
    for entry in payload:
        if not ipa:
            hwi = entry.get("hwi", {}) if isinstance(entry.get("hwi"), dict) else {}
            prs = hwi.get("prs", [])
            if isinstance(prs, list):
                for pr_item in prs:
                    if isinstance(pr_item, dict):
                        candidate = _clean(pr_item.get("mw", ""))
                        if candidate:
                            ipa = candidate
                            break
        if not definition:
            shortdef = entry.get("shortdef", [])
            if isinstance(shortdef, list) and shortdef:
                definition = _clean(str(shortdef[0]))
        if definition:
            break
    return {
        "ipa": ipa,
        "definition": definition,
        "examples": "",
        "synonyms": "",
        "antonyms": "",
    }


def _extract_cambridge_ipa(html_text: str) -> str:
    ipa = ""
    # Prefer UK IPA, then fallback to first IPA on page.
    uk_match = re.search(
        r'<span class="[^"]*\buk dpron-i\b[^"]*">.*?<span class="[^"]*\bipa\b[^"]*">\s*([^<]+?)\s*</span>',
        html_text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    first_match = re.search(r'<span class="[^"]*\bipa\b[^"]*">\s*([^<]+?)\s*</span>', html_text, flags=re.IGNORECASE)
    match = uk_match or first_match
    if match:
        ipa = _clean(html.unescape(match.group(1)))
    if ipa and not ipa.startswith("/"):
        ipa = "/" + ipa
    if ipa and not ipa.endswith("/"):
        ipa = ipa + "/"
    return ipa


_DCAMBRIDGE_PHRASE_BODY_OPEN = re.compile(
    r'<div\b[^>]*(?:\bphrase-body\b[^>]*\bdphrase_b\b|\bdphrase_b\b[^>]*\bphrase-body\b)[^>]*>',
    re.IGNORECASE,
)


def _consume_cambridge_div_from(html_text: str, start: int) -> int:
    """Index after the </div> that closes the opening <div> at start (balanced, shallow scan)."""
    if start >= len(html_text) or html_text[start : start + 4].lower() != "<div":
        return start
    depth = 1
    pos = start + 4
    lower = html_text.lower()
    while pos < len(html_text) and depth > 0:
        i_open = lower.find("<div", pos)
        i_close = lower.find("</div>", pos)
        if i_close < 0:
            return len(html_text)
        if i_open >= 0 and i_open < i_close:
            depth += 1
            pos = i_open + 4
        else:
            depth -= 1
            pos = i_close + len("</div>")
    return pos


def _strip_cambridge_nested_phrase_blocks(html_text: str) -> str:
    """Drop phrase-body sections so nested expressions (e.g. break out in a sweat) are not main senses."""
    out = html_text
    while True:
        m = _DCAMBRIDGE_PHRASE_BODY_OPEN.search(out)
        if not m:
            break
        start = m.start()
        end = _consume_cambridge_div_from(out, start)
        out = out[:start] + out[end:]
    return out


def _extract_cambridge_senses(html_text: str) -> List[Dict[str, Any]]:
    class _CambridgeBlockParser(HTMLParser):
        def __init__(self) -> None:
            super().__init__()
            self.definition = ""
            self.examples: List[str] = []
            self.synonyms: List[str] = []
            self.antonyms: List[str] = []
            self.usage_labels: List[str] = []
            self.image_url = ""
            self._in_synonyms_block = False
            self._synonyms_block_depth = 0
            self._in_antonyms_block = False
            self._antonyms_block_depth = 0
            self._capture_kind: Optional[str] = None
            self._capture_depth = 0
            self._capture_parts: List[str] = []

        def _class_list(self, attrs: List[tuple]) -> List[str]:
            class_attr = ""
            for k, v in attrs:
                if k == "class" and v:
                    class_attr = v
                    break
            return [c.strip().lower() for c in class_attr.split() if c.strip()]

        def handle_starttag(self, tag: str, attrs: List[tuple]) -> None:
            classes = self._class_list(attrs)
            is_definition = "ddef_d" in classes
            is_example = ("dexamp" in classes) or ("examp" in classes)
            is_synonyms_block = ("xref" in classes) and ("synonyms" in classes)
            is_antonyms_block = ("xref" in classes) and ("opposite" in classes)
            is_synonym_item = ("x-h" in classes) or ("dx-h" in classes)
            is_usage_label = ("usage" in classes) or ("dusage" in classes)
            is_image_tag = tag.lower() in ("img", "amp-img")

            if is_image_tag and not self.image_url:
                src = ""
                class_text = " ".join(classes)
                for k, v in attrs:
                    key = str(k).lower()
                    if key == "src" and v:
                        src = str(v).strip()
                        break
                    if key in ("data-src", "data-src-mp", "data-image"):
                        src = str(v).strip()
                # Skip decorative/media-control images and keep definition pictures.
                if src and ("speaker" in src.lower() or "audio" in src.lower()):
                    src = ""
                if src and ("dimg" not in class_text) and ("/images/full/" not in src.lower()):
                    src = ""
                if src:
                    if src.startswith("//"):
                        src = "https:" + src
                    elif src.startswith("/"):
                        src = "https://dictionary.cambridge.org" + src
                    self.image_url = src

            if is_synonyms_block:
                self._in_synonyms_block = True
                self._synonyms_block_depth = 1
            elif self._in_synonyms_block:
                self._synonyms_block_depth += 1

            if is_antonyms_block:
                self._in_antonyms_block = True
                self._antonyms_block_depth = 1
            elif self._in_antonyms_block:
                self._antonyms_block_depth += 1

            if self._capture_kind:
                self._capture_depth += 1
                return

            if is_definition and not self.definition:
                self._capture_kind = "definition"
                self._capture_depth = 1
                self._capture_parts = []
            elif is_example:
                self._capture_kind = "example"
                self._capture_depth = 1
                self._capture_parts = []
            elif self._in_synonyms_block and is_synonym_item:
                self._capture_kind = "synonym"
                self._capture_depth = 1
                self._capture_parts = []
            elif self._in_antonyms_block and is_synonym_item:
                self._capture_kind = "antonym"
                self._capture_depth = 1
                self._capture_parts = []
            elif (not self._in_synonyms_block) and (not self._in_antonyms_block) and is_usage_label:
                self._capture_kind = "usage"
                self._capture_depth = 1
                self._capture_parts = []

        def handle_endtag(self, _tag: str) -> None:
            if self._capture_kind:
                self._capture_depth -= 1
                if self._capture_depth == 0:
                    text = _clean(html.unescape("".join(self._capture_parts)))
                    if text:
                        if self._capture_kind == "definition":
                            self.definition = text
                        elif self._capture_kind == "example":
                            self.examples.append(text)
                        elif self._capture_kind == "synonym":
                            self.synonyms.append(text)
                        elif self._capture_kind == "antonym":
                            self.antonyms.append(text)
                        elif self._capture_kind == "usage":
                            self.usage_labels.append(text)
                    self._capture_kind = None
                    self._capture_parts = []
                return

            if self._in_synonyms_block:
                self._synonyms_block_depth -= 1
                if self._synonyms_block_depth <= 0:
                    self._in_synonyms_block = False
                    self._synonyms_block_depth = 0

            if self._in_antonyms_block:
                self._antonyms_block_depth -= 1
                if self._antonyms_block_depth <= 0:
                    self._in_antonyms_block = False
                    self._antonyms_block_depth = 0

        def handle_data(self, data: str) -> None:
            if self._capture_kind:
                self._capture_parts.append(data)

    html_text = _strip_cambridge_nested_phrase_blocks(html_text)

    senses: List[Dict[str, Any]] = []
    parts = re.split(
        r'(?=<div[^>]*class=["\'][^"\']*\bdef-block\b[^"\']*["\'][^>]*>)',
        html_text,
        flags=re.IGNORECASE,
    )
    seen_definitions = set()
    for part in parts:
        if "def-block" not in part:
            continue
        parser = _CambridgeBlockParser()
        parser.feed(part)
        definition = _clean(parser.definition)
        if not definition or definition in seen_definitions:
            continue
        seen_definitions.add(definition)
        senses.append(
            {
                "definition": definition,
                "examples": [e for e in parser.examples if _clean(e)],
                "synonyms": [s for s in parser.synonyms if _clean(s)],
                "antonyms": [a for a in parser.antonyms if _clean(a)],
                "labels": [l for l in parser.usage_labels if _clean(l)],
                "image_url": parser.image_url,
            }
        )

    if senses:
        return senses

    # Fallback when def-blocks are not present in HTML response.
    fallback: List[Dict[str, Any]] = []
    parser = _CambridgeBlockParser()
    parser.feed(html_text)
    if parser.definition:
        fallback.append(
            {
                "definition": parser.definition,
                "examples": parser.examples,
                "synonyms": parser.synonyms,
                "antonyms": parser.antonyms,
                "labels": parser.usage_labels,
                "image_url": parser.image_url,
            }
        )
    if fallback:
        return fallback

    # Last-resort fallback for pages where only ddef_d fragments are available.
    seen = set()
    for m in re.finditer(
        r'<(?:div|span)[^>]*class=["\'][^"\']*\bddef_d\b[^"\']*["\'][^>]*>(.*?)</(?:div|span)>',
        html_text,
        flags=re.IGNORECASE | re.DOTALL,
    ):
        raw = re.sub(r"<[^>]+>", " ", m.group(1))
        definition = _clean(html.unescape(raw))
        if definition and definition not in seen:
            fallback.append(
                {
                    "definition": definition,
                    "examples": [],
                    "synonyms": [],
                    "antonyms": [],
                    "labels": [],
                    "image_url": "",
                }
            )
            seen.add(definition)
    return fallback


def _extract_cambridge_global_usage_labels(html_text: str) -> List[str]:
    first_def_block_idx = html_text.lower().find("def-block")
    scope = html_text if first_def_block_idx < 0 else html_text[:first_def_block_idx]
    labels: List[str] = []
    for match in re.finditer(
        r'<span[^>]*class=["\'][^"\']*\b(?:usage|dusage)\b[^"\']*["\'][^>]*>(.*?)</span>',
        scope,
        flags=re.IGNORECASE | re.DOTALL,
    ):
        raw = re.sub(r"<[^>]+>", " ", match.group(1))
        text = _clean(html.unescape(raw))
        if text:
            labels.append(text)
    out: List[str] = []
    seen = set()
    for label in labels:
        key = label.lower()
        if key not in seen:
            out.append(label)
            seen.add(key)
    return out


def _extract_from_cambridge(
    html_text: str,
    word: str,
    separator: str,
    max_definitions: int,
    max_examples: int,
    selected_definition: Optional[str] = None,
) -> Dict[str, str]:
    ipa = _extract_cambridge_ipa(html_text)
    senses = _extract_cambridge_senses(html_text)

    def unique(items: List[str]) -> List[str]:
        out: List[str] = []
        seen = set()
        for item in items:
            if item not in seen:
                out.append(item)
                seen.add(item)
        return out

    chosen_definitions: List[str] = []
    chosen_examples: List[str] = []
    chosen_labels: List[str] = []
    chosen_image_url = ""
    global_labels = _extract_cambridge_global_usage_labels(html_text)
    if selected_definition:
        for sense in senses:
            if sense.get("definition") == selected_definition:
                chosen_definitions = [selected_definition]
                chosen_examples = list(sense.get("examples", []))
                chosen_labels = list(sense.get("labels", []))
                chosen_image_url = _clean(str(sense.get("image_url", "")))
                break
    if not chosen_definitions:
        chosen_definitions = [s.get("definition", "") for s in senses if s.get("definition")]
        chosen_definitions = unique(chosen_definitions)[:max_definitions]
        if senses:
            chosen_examples = list(senses[0].get("examples", []))
            chosen_labels = list(senses[0].get("labels", []))
            chosen_image_url = _clean(str(senses[0].get("image_url", "")))

    uniq_examples = unique(chosen_examples)[:max_examples]
    thes_html = _request_cambridge_thesaurus_html(word)
    if thes_html:
        uniq_synonyms, uniq_antonyms = _extract_cambridge_thesaurus_synonyms_antonyms(thes_html)
    else:
        uniq_synonyms, uniq_antonyms = [], []
    uniq_labels = unique(chosen_labels + global_labels)

    examples_html = ""
    if uniq_examples:
        escaped_word = re.escape(word)
        pattern = re.compile(escaped_word, flags=re.IGNORECASE) if escaped_word else None
        rendered: List[str] = []
        for example in uniq_examples:
            safe_example = html.escape(example)
            if pattern:
                safe_example = pattern.sub(lambda m: f"<u>{html.escape(m.group(0))}</u>", safe_example)
            rendered.append(f"<li>{safe_example}</li>")
        examples_html = "<ul>" + "".join(rendered) + "</ul>"

    return {
        "ipa": ipa,
        "definition": separator.join(chosen_definitions),
        "examples": examples_html,
        "synonyms": separator.join(uniq_synonyms),
        "antonyms": separator.join(uniq_antonyms),
        "image": (f'<img src="{html.escape(chosen_image_url, quote=True)}">' if chosen_image_url else ""),
        "usage_labels": separator.join(uniq_labels),
    }


def _apply_tags(note: Note, tags: List[str]) -> bool:
    changed = False
    for raw_tag in tags:
        tag = _clean(raw_tag).replace(" ", "_")
        if not tag:
            continue
        try:
            if hasattr(note, "has_tag") and note.has_tag(tag):
                continue
            note.add_tag(tag)
            changed = True
        except Exception:
            continue
    return changed


def _download_cambridge_image_to_media(image_url: str, word: str) -> Optional[str]:
    if mw.col is None or not image_url:
        return None
    req = urllib.request.Request(
        image_url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0.0.0 Safari/537.36"
            ),
            "Accept": "image/*,*/*;q=0.8",
            "Referer": "https://dictionary.cambridge.org/",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as response:
            data = response.read()
            content_type = str(response.headers.get("Content-Type", "")).lower()
    except Exception:
        return None
    if not data:
        return None
    if content_type and ("image/" not in content_type):
        return None

    parsed = urllib.parse.urlparse(image_url)
    ext = os.path.splitext(parsed.path)[1].lower()
    if ext not in (".jpg", ".jpeg", ".png", ".webp", ".gif"):
        if "png" in content_type:
            ext = ".png"
        elif "webp" in content_type:
            ext = ".webp"
        elif "gif" in content_type:
            ext = ".gif"
        else:
            ext = ".jpg"

    safe_word = re.sub(r"[^a-zA-Z0-9_-]+", "_", word).strip("_") or "word"
    digest = hashlib.md5(image_url.encode("utf-8")).hexdigest()[:10]
    filename = f"cambridge_{safe_word}_{digest}{ext}"

    media = mw.col.media
    try:
        if hasattr(media, "write_data"):
            media.write_data(filename, data)
            return filename
    except Exception:
        pass

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(data)
            temp_path = tmp.name
        added = media.add_file(temp_path)
        return os.path.basename(added) if added else None
    except Exception:
        return None
    finally:
        try:
            if "temp_path" in locals() and temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
        except Exception:
            pass


def _localize_cambridge_image_html(image_html: str, word: str) -> str:
    if not image_html:
        return image_html
    match = re.search(r'<img[^>]*\bsrc=["\']([^"\']+)["\']', image_html, flags=re.IGNORECASE)
    if not match:
        return image_html
    remote_url = _clean(html.unescape(match.group(1)))
    if not remote_url or not remote_url.startswith(("http://", "https://")):
        return image_html
    local_filename = _download_cambridge_image_to_media(remote_url, word)
    if not local_filename:
        return image_html
    return f'<img src="{html.escape(local_filename, quote=True)}">'


def _find_existing_emoji_variant(base_label_tag: str) -> Optional[str]:
    if mw.col is None:
        return None
    try:
        existing_tags = mw.col.tags.all()
    except Exception:
        return None
    if not existing_tags:
        return None

    normalized = _clean(base_label_tag).replace(" ", "_")
    if not normalized:
        return None

    emoji_suffix_re = re.compile(
        rf"^{re.escape(normalized)}[\u2600-\u27BF\U0001F300-\U0001FAFF\U0001F1E6-\U0001F1FF\ufe0f]+$",
        flags=re.IGNORECASE,
    )
    for tag in existing_tags:
        candidate = _clean(str(tag)).replace(" ", "_")
        if emoji_suffix_re.match(candidate):
            return candidate
    return None


def _merge_non_empty(base: Dict[str, str], extra: Dict[str, str]) -> Dict[str, str]:
    merged = base.copy()
    for key, value in extra.items():
        if value and not merged.get(key):
            merged[key] = value
    return merged


def _extract_definition_candidates(payload: List[Dict[str, Any]]) -> List[str]:
    definitions: List[str] = []
    for entry in payload:
        for meaning in entry.get("meanings", []):
            for d in meaning.get("definitions", []):
                definition = _clean(d.get("definition", ""))
                if definition:
                    definitions.append(definition)

    unique_defs: List[str] = []
    seen = set()
    for definition in definitions:
        if definition not in seen:
            unique_defs.append(definition)
            seen.add(definition)
    return unique_defs


def _choose_definition_dialog(
    definitions: List[str],
    parent: Optional[QWidget] = None,
) -> Optional[str]:
    if not definitions:
        return None
    if len(definitions) == 1:
        return definitions[0]

    dialog = QDialog(parent or mw)
    dialog.setWindowTitle("Choose definition")
    dialog.setWindowModality(Qt.WindowModality.WindowModal)
    layout = QVBoxLayout(dialog)

    combo = QComboBox(dialog)
    combo.addItems(definitions)
    layout.addWidget(combo)

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, parent=dialog)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    if dialog.exec() != QDialog.DialogCode.Accepted:
        return None
    return combo.currentText() or None


def _set_field(note: Any, field_name: str, value: str, overwrite_existing: bool) -> bool:
    if not field_name or field_name not in note:
        return False
    existing = _clean(html.unescape(note[field_name]))
    if existing and not overwrite_existing:
        return False
    note[field_name] = value
    return True


def _show_field_mapping_dialog(
    available_fields: List[str],
    cfg: Dict[str, Any],
    parent: Optional[QWidget] = None,
) -> Optional[Dict[str, Any]]:
    if not available_fields:
        showInfo("Current note type has no fields.")
        return None

    dialog = QDialog(parent or mw)
    dialog.setWindowTitle("Choose fields for enrichment")
    dialog.setWindowModality(Qt.WindowModality.WindowModal)
    layout = QVBoxLayout(dialog)
    form = QFormLayout()
    layout.addLayout(form)

    source_field_combo = QComboBox(dialog)
    source_field_combo.addItems(available_fields)
    if cfg["source_field"] in available_fields:
        source_field_combo.setCurrentText(cfg["source_field"])
    form.addRow("Source word field:", source_field_combo)

    map_labels = {
        "ipa": "IPA field:",
        "definition": "Definition field:",
        "examples": "Examples field:",
        "synonyms": "Synonyms field:",
        "antonyms": "Antonyms field:",
        "image": "Image field:",
    }
    target_combos: Dict[str, QComboBox] = {}
    for key, label in map_labels.items():
        combo = QComboBox(dialog)
        combo.addItems(available_fields)
        default_value = cfg["field_map"].get(key, "")
        if default_value in available_fields:
            combo.setCurrentText(default_value)
        target_combos[key] = combo
        form.addRow(label, combo)

    overwrite_checkbox = QCheckBox("Overwrite existing target fields", dialog)
    overwrite_checkbox.setChecked(bool(cfg.get("overwrite_existing", False)))
    form.addRow("Options:", overwrite_checkbox)

    data_source_combo = QComboBox(dialog)
    data_source_combo.addItems(["custom", "dictionaryapi", "cambridge", "wordnik", "merriam_webster"])
    if cfg.get("data_source", "custom") in ["custom", "dictionaryapi", "cambridge", "wordnik", "merriam_webster"]:
        data_source_combo.setCurrentText(cfg.get("data_source", "custom"))
    form.addRow("Dictionary source:", data_source_combo)

    wordnik_key = QLineEdit(dialog)
    wordnik_key.setEchoMode(QLineEdit.EchoMode.Password)
    wordnik_key.setText(cfg.get("api_keys", {}).get("wordnik", ""))
    form.addRow("Wordnik API key:", wordnik_key)

    merriam_collegiate_key = QLineEdit(dialog)
    merriam_collegiate_key.setEchoMode(QLineEdit.EchoMode.Password)
    merriam_collegiate_key.setText(cfg.get("api_keys", {}).get("merriam_collegiate", ""))
    form.addRow("Merriam Collegiate key:", merriam_collegiate_key)

    merriam_sd3_key = QLineEdit(dialog)
    merriam_sd3_key.setEchoMode(QLineEdit.EchoMode.Password)
    merriam_sd3_key.setText(cfg.get("api_keys", {}).get("merriam_sd3", ""))
    form.addRow("Merriam SD3 key:", merriam_sd3_key)

    merriam_legacy_key = QLineEdit(dialog)
    merriam_legacy_key.setEchoMode(QLineEdit.EchoMode.Password)
    merriam_legacy_key.setText(cfg.get("api_keys", {}).get("merriam_webster", ""))
    form.addRow("Merriam legacy key (optional):", merriam_legacy_key)

    def update_api_visibility() -> None:
        selected = data_source_combo.currentText()
        wordnik_key.setVisible(selected in ("wordnik", "custom"))
        form.labelForField(wordnik_key).setVisible(selected in ("wordnik", "custom"))
        show_merriam = selected in ("merriam_webster", "custom")
        merriam_collegiate_key.setVisible(show_merriam)
        form.labelForField(merriam_collegiate_key).setVisible(show_merriam)
        merriam_sd3_key.setVisible(show_merriam)
        form.labelForField(merriam_sd3_key).setVisible(show_merriam)
        merriam_legacy_key.setVisible(show_merriam)
        form.labelForField(merriam_legacy_key).setVisible(show_merriam)

    data_source_combo.currentIndexChanged.connect(lambda _i: update_api_visibility())
    update_api_visibility()

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, parent=dialog)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    if dialog.exec() != QDialog.DialogCode.Accepted:
        return None

    field_map = {key: combo.currentText() for key, combo in target_combos.items()}
    source_field = source_field_combo.currentText()
    if not source_field:
        showInfo("Source field cannot be empty.")
        return None
    return {
        "source_field": source_field,
        "field_map": field_map,
        "overwrite_existing": overwrite_checkbox.isChecked(),
        "data_source": data_source_combo.currentText(),
        "api_keys": {
            "wordnik": _clean(wordnik_key.text()),
            "merriam_webster": _clean(merriam_legacy_key.text()),
            "merriam_collegiate": _clean(merriam_collegiate_key.text()),
            "merriam_sd3": _clean(merriam_sd3_key.text()),
        },
    }


def _persist_field_choices(
    cfg: Dict[str, Any],
    source_field: str,
    field_map: Dict[str, str],
    overwrite_existing: bool,
    data_source: str,
    api_keys: Dict[str, str],
) -> None:
    cfg["source_field"] = source_field
    cfg["field_map"] = field_map
    cfg["overwrite_existing"] = overwrite_existing
    cfg["data_source"] = data_source
    cfg["api_keys"] = api_keys
    mw.addonManager.writeConfig(__name__, cfg)


def _enrich_note(
    note: Note,
    source_field: str,
    cfg: Dict[str, Any],
    payload: Optional[Any] = None,
    selected_definition: Optional[str] = None,
) -> str:
    resolved_source_field = _resolve_note_field_name(note, source_field)
    if not resolved_source_field:
        return "missing_source"
    word = _plain_text(note[resolved_source_field])
    if not word:
        return "missing_source"

    source = cfg.get("data_source", "custom")
    details = {"ipa": "", "definition": "", "examples": "", "synonyms": "", "antonyms": "", "image": ""}

    if source == "dictionaryapi":
        actual_payload = payload or _request_dictionary_data(word)
        if not actual_payload:
            return "no_api_result"
        details = _extract_details(
            actual_payload,
            word,
            cfg["separator"],
            int(cfg["max_definitions"]),
            int(cfg["max_examples"]),
            selected_definition=selected_definition,
        )
    elif source == "cambridge":
        cambridge_html = payload if isinstance(payload, str) else _request_cambridge_html(word)
        if not cambridge_html:
            return "no_api_result"
        details = _extract_from_cambridge(
            cambridge_html,
            word,
            cfg["separator"],
            int(cfg["max_definitions"]),
            int(cfg["max_examples"]),
            selected_definition=selected_definition,
        )
        details["image"] = _localize_cambridge_image_html(details.get("image", ""), word)
        if not any(details.values()):
            return "no_api_result"
    elif source == "wordnik":
        wk = _request_wordnik_data(word, cfg.get("api_keys", {}).get("wordnik", ""))
        if not wk:
            return "no_api_result"
        details = _extract_from_wordnik(wk, word, cfg["separator"], int(cfg["max_examples"]))
    elif source == "merriam_webster":
        mw_data = _request_merriam_data(word, cfg.get("api_keys", {}))
        if not mw_data:
            return "no_api_result"
        details = _extract_from_merriam(mw_data, cfg["separator"])
        if details.get("examples"):
            details["examples"] = _extract_details(
                [{"phonetic": "", "meanings": [{"definitions": [{"example": details["examples"]}]}]}],
                word,
                cfg["separator"],
                1,
                int(cfg["max_examples"]),
            )["examples"]
    else:
        # Custom cascade: dictionaryapi -> Wordnik/Merriam fallback, Datamuse for related words.
        primary = payload or _request_dictionary_data(word)
        if primary:
            details = _extract_details(
                primary,
                word,
                cfg["separator"],
                int(cfg["max_definitions"]),
                int(cfg["max_examples"]),
                selected_definition=selected_definition,
            )
        wk = _request_wordnik_data(word, cfg.get("api_keys", {}).get("wordnik", ""))
        if wk:
            details = _merge_non_empty(details, _extract_from_wordnik(wk, word, cfg["separator"], int(cfg["max_examples"])))
        mw_data = _request_merriam_data(word, cfg.get("api_keys", {}))
        if mw_data:
            details = _merge_non_empty(details, _extract_from_merriam(mw_data, cfg["separator"]))
        dm = _request_datamuse_data(word)
        if dm.get("synonyms") and not details.get("synonyms"):
            details["synonyms"] = cfg["separator"].join(dm["synonyms"])
        if dm.get("antonyms") and not details.get("antonyms"):
            details["antonyms"] = cfg["separator"].join(dm["antonyms"])
        if not any(details.values()):
            return "no_api_result"
    changed = False
    for key, target_field in cfg["field_map"].items():
        value = details.get(key, "")
        if not value:
            continue
        field_overwrite = bool(cfg["overwrite_existing"]) or (source == "cambridge" and key == "image")
        if _set_field(note, target_field, value, field_overwrite):
            changed = True

    if source == "cambridge" and bool(cfg.get("cambridge_usage_tags_enabled", True)):
        labels_raw = details.get("usage_labels", "")
        labels = [_clean(x).lower() for x in labels_raw.split(cfg["separator"]) if _clean(x)]
        tags_to_add: List[str] = []
        for label in labels:
            emoji_variant = _find_existing_emoji_variant(label)
            tags_to_add.append(emoji_variant if emoji_variant else label)
        if _apply_tags(note, tags_to_add):
            changed = True

    if changed:
        note.flush()
        return "updated"
    return "skipped"


def _enrich_note_ids(note_ids: List[int], source_field: str, cfg: Dict[str, Any]) -> None:
    if mw.col is None:
        return

    updated = 0
    skipped = 0
    missing_source = 0
    no_api_result = 0

    mw.checkpoint("Enrich vocabulary fields")
    for nid in note_ids:
        note = mw.col.get_note(nid)
        result = _enrich_note(note, source_field, cfg)
        if result == "updated":
            updated += 1
        elif result == "skipped":
            skipped += 1
        elif result == "missing_source":
            missing_source += 1
        elif result == "no_api_result":
            no_api_result += 1

    mw.reset()
    tooltip(
        "Vocabulary enrichment finished. "
        f"Updated: {updated} | Skipped: {skipped} | "
        f"Missing source: {missing_source} | No API result: {no_api_result}"
    )


def _collect_same_notetype_note_ids(note_ids: List[int]) -> (List[int], int, List[str]):
    if mw.col is None or not note_ids:
        return [], 0, []

    first_note = mw.col.get_note(note_ids[0])
    target_mid = first_note.mid
    same_type_ids: List[int] = []
    skipped_type_count = 0

    for nid in note_ids:
        note = mw.col.get_note(nid)
        if note.mid == target_mid:
            same_type_ids.append(nid)
        else:
            skipped_type_count += 1

    return same_type_ids, skipped_type_count, list(first_note.keys())


def enrich_selected_notes() -> None:
    if mw.col is None:
        return
    cfg = _read_config()
    browser = mw.app.activeWindow()
    if not hasattr(browser, "selectedNotes"):
        showInfo("Open the Browser and select notes first.")
        return

    note_ids = browser.selectedNotes()
    if not note_ids:
        showInfo("Select at least one note in the Browser.")
        return

    proceed = askUser(
        f"Update {len(note_ids)} notes with dictionary fields?\n"
        "This will call the selected dictionary source for each word."
    )
    if not proceed:
        return

    same_type_ids, skipped_type_count, available_fields = _collect_same_notetype_note_ids(note_ids)
    if not same_type_ids:
        showInfo("No notes of a single note type selected.")
        return

    cfg["source_field"] = _auto_heal_source_field(cfg, mw.col.get_note(same_type_ids[0]))
    if cfg["source_field"] not in available_fields:
        showInfo(
            f"Configured source field '{cfg['source_field']}' is not in this note type. "
            "Open Browse -> losev -> Settings and choose valid fields."
        )
        return

    _enrich_note_ids(same_type_ids, cfg["source_field"], cfg)
    if skipped_type_count:
        showInfo(f"Skipped {skipped_type_count} selected notes with a different note type.")


def enrich_current_browser_note(editor: Editor) -> None:
    global LAST_API_ERROR
    if mw.col is None:
        return
    parent = editor.parentWindow
    if not isinstance(parent, Browser):
        showInfo("This button is intended for Browser note view.")
        return
    if editor.note is None:
        showInfo("Select a note in Browser first.")
        return

    cfg = _read_config()
    try:
        editor.saveNow(lambda: None)
    except Exception:
        pass

    working_note = editor.note
    if working_note is None:
        showInfo("Select a note in Browser first.")
        return
    try:
        db_note = mw.col.get_note(working_note.id)
        if db_note is not None:
            working_note = db_note
    except Exception:
        pass

    source_field = _auto_heal_source_field(cfg, working_note)

    resolved_source_field = _resolve_note_field_name(working_note, source_field)
    if not resolved_source_field:
        available = ", ".join(list(working_note.keys()))
        showInfo(
            f"Source field '{source_field}' is missing or empty in this note. "
            "Open Browse -> losev -> Settings to change it.\n\n"
            f"Available fields in note: {available}"
        )
        return

    word = _plain_text(working_note[resolved_source_field])
    if not word:
        fallback_field = _find_first_nonempty_field(working_note)
        if fallback_field:
            word = _plain_text(working_note[fallback_field])
            resolved_source_field = fallback_field
        else:
            available = ", ".join(list(working_note.keys()))
            showInfo(
                f"Configured source field: '{source_field}'\n"
                f"Resolved field: '{resolved_source_field}'\n"
                "Resolved field value is empty.\n\n"
                f"Available fields in note: {available}"
            )
            return

    if not word:
        showInfo(
            f"Source field '{source_field}' is missing or empty in this note. "
            "Open Browse -> losev -> Settings to change it."
        )
        return

    selected_definition = None
    payload = None
    if cfg.get("data_source", "custom") in ("dictionaryapi", "custom"):
        payload = _request_dictionary_data(word)
        if payload:
            definitions = _extract_definition_candidates(payload)
            selected_definition = _choose_definition_dialog(definitions, parent=parent)
            if definitions and not selected_definition:
                return
    elif cfg.get("data_source") == "cambridge":
        cambridge_html = _request_cambridge_html(word)
        if cambridge_html:
            senses = _extract_cambridge_senses(cambridge_html)
            definitions = [s.get("definition", "") for s in senses if s.get("definition")]
            selected_definition = _choose_definition_dialog(definitions, parent=parent)
            if definitions and not selected_definition:
                return
            payload = cambridge_html

    result = _enrich_note(
        working_note,
        source_field,
        cfg,
        payload=payload,
        selected_definition=selected_definition,
    )
    if result == "updated":
        mw.reset()
        editor.loadNoteKeepingFocus()
        tooltip("Current note updated.")
    elif result == "skipped":
        showInfo("Nothing changed (fields already filled or no values found).")
    elif result == "missing_source":
        showInfo(
            f"Source field '{cfg['source_field']}' is missing or empty in this note. "
            "Open Browse -> losev -> Settings to change it."
        )
    else:
        msg = "No dictionary result for this word."
        if cfg.get("data_source") == "merriam_webster" and LAST_API_ERROR:
            msg += f"\n\nDetails: {LAST_API_ERROR}"
        showInfo(msg)


def copy_word_and_definition(editor: Editor) -> None:
    if mw.col is None:
        return
    parent = editor.parentWindow
    if not isinstance(parent, Browser):
        showInfo("This button is intended for Browser note view.")
        return
    if editor.note is None:
        showInfo("Select a note in Browser first.")
        return

    cfg = _read_config()
    working_note = editor.note
    if working_note is None:
        showInfo("Select a note in Browser first.")
        return

    source_field = _auto_heal_source_field(cfg, working_note)
    source_field = _resolve_note_field_name(working_note, source_field) or source_field
    definition_field = cfg.get("field_map", {}).get("definition", "Definition")
    definition_field = _resolve_note_field_name(working_note, definition_field) or definition_field

    word = _plain_text(working_note[source_field]) if source_field in working_note else ""
    definition = _plain_text(working_note[definition_field]) if definition_field in working_note else ""

    if not word and not definition:
        showInfo("Word and definition are empty for the current note.")
        return

    text_to_copy = f"{word}\n{definition}"
    QApplication.clipboard().setText(text_to_copy)
    tooltip("Copied word + definition to clipboard.")


def open_browser_settings(browser: Browser) -> None:
    cfg = _read_config()

    note_for_fields = browser.editor.note
    if note_for_fields is None:
        note_ids = browser.selectedNotes()
        if not note_ids:
            showInfo("Select a note in Browser first to configure fields.")
            return
        note_for_fields = mw.col.get_note(note_ids[0])

    choices = _show_field_mapping_dialog(list(note_for_fields.keys()), cfg, parent=browser)
    if not choices:
        return

    _persist_field_choices(
        cfg,
        choices["source_field"],
        choices["field_map"],
        bool(choices["overwrite_existing"]),
        choices["data_source"],
        choices["api_keys"],
    )
    tooltip("Settings saved.")


def _add_editor_button(buttons: List[str], editor: Editor) -> List[str]:
    if not isinstance(editor.parentWindow, Browser):
        return buttons
    button = editor.addButton(
        icon=None,
        cmd="vocab_enrich_current_note",
        func=lambda ed=editor: enrich_current_browser_note(ed),
        tip="Auto-fill dictionary fields for current note",
        label="Enrich Note",
    )
    buttons.append(button)
    copy_button = editor.addButton(
        icon=None,
        cmd="copy_word_definition",
        func=lambda ed=editor: copy_word_and_definition(ed),
        tip="Copy word and definition to clipboard",
        label="📋 Copy W+D",
    )
    buttons.append(copy_button)
    return buttons


def _add_browser_menu(browser: Browser) -> None:
    losev_menu = QMenu("losev", browser.form.menubar)
    settings_action = QAction("Settings", losev_menu)
    settings_action.triggered.connect(lambda _checked=False, b=browser: open_browser_settings(b))
    losev_menu.addAction(settings_action)
    browser.form.menubar.addMenu(losev_menu)


def _add_menu() -> None:
    action = QAction(MENU_LABEL, mw)
    action.triggered.connect(enrich_selected_notes)
    mw.form.menuTools.addAction(action)


_add_menu()
gui_hooks.editor_did_init_buttons.append(_add_editor_button)
gui_hooks.browser_menus_did_init.append(_add_browser_menu)
