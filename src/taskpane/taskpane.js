/* =============================================================================
 * BiblioAudit Assistant — core task pane logic
 * -----------------------------------------------------------------------------
 * Bidirectional audit of GOST-style numeric citations:
 *   body text  ->  reference list   (dead sources)
 *   reference list  ->  body text   (phantom references)
 *
 * The file is organised in four layers:
 *   1. Configuration + regular expressions
 *   2. Pure parsing / audit core (no Office.js dependency — unit-testable)
 *   3. Report serialisers (Markdown / JSON / plain lines)
 *   4. Rendering + Office.js integration
 *
 * global Office, Word
 * ===========================================================================*/

(function () {
  "use strict";

  /* ===========================================================================
   * 1. CONFIGURATION
   * =========================================================================*/

  var CONFIG = {
    /** Default "bulk citation" threshold: more than N sources in one bracket. */
    defaultBulkThreshold: 7,
    /** A range wider than this is treated as malformed, not expanded. */
    maxRangeExpansion: 500,
    /** Highest plausible reference number; larger values are not citations. */
    maxSourceNumber: 2000,
    /** Minimum length of a trailing numbered block to accept as a reference list. */
    minTrailingBlock: 3,
    /**
     * How many consecutive non-entry paragraphs may be folded into the previous
     * bibliographic record before the reference list is considered finished.
     * Guards against swallowing whatever follows the list (a second-language
     * version of the article, appendices, author details).
     */
    maxContinuationParagraphs: 2,
    /** Characters of context captured around each citation. */
    snippetRadius: 70,
    /** Maximum reference description length kept in the report. */
    maxDescriptionLength: 400,

    /* --- Word export formatting ------------------------------------------ */
    /** Prefix of the generated report file: "АудитЛит <имя исходного файла>". */
    exportFilePrefix: "АудитЛит",
    exportFontName: "Times New Roman",
    exportFontSize: 12,
    /**
     * Line spacing in points. Word's UI divides this value by 12, so 12pt is
     * exactly single spacing.
     */
    exportLineSpacing: 12
  };

  /** All dash characters that may separate the ends of a range. */
  var DASHES = "\\-\\u2010\\u2011\\u2012\\u2013\\u2014\\u2015\\u2212";

  var PATTERNS = {
    /**
     * Reference-section heading, matched against the folded form of the line
     * (lower case, ё→е, Tajik letters folded to their Russian base), so one
     * pattern covers Russian, English and Tajik headings.
     */
    referenceHeading: new RegExp(
      "^(?:\\d+(?:\\.\\d+)*\\s*[.)]?\\s*)?" +
      "(?:" +
        // --- русский ---
        "список\\s+использованн(?:ой|ых|ы)\\s+(?:литератур[ыа]|источников)|" +
        "список\\s+(?:литератур[ыа]|источников|использованных\\s+источников)|" +
        "библиографическ(?:ий\\s+список|ий\\s+перечень|ая\\s+справка)|" +
        "использованн(?:ая\\s+литература|ые\\s+источники)|" +
        "литература|библиография|источники|" +
        // --- английский ---
        "references?(?:\\s+list)?|list\\s+of\\s+references|bibliography|" +
        "works\\s+cited|literature(?:\\s+cited)?|" +
        // --- таджикский (в свёрнутом виде: ӯ→у, ҳ→х, қ→к, ҷ→ч, ғ→г, ӣ→и) ---
        "адабиет(?:и\\s+истифодашуда)?|руйхати\\s+адабиет|фехристи\\s+адабиет|" +
        "русйхати\\s+адабиет|сарчашмахо|манбахо|адабиетхо" +
      ")\\s*[:.]?$"
    ),

    /** A numbered bibliography entry: "1. Ivanov…", "1) Ivanov…", "[1] Ivanov…". */
    referenceEntry: new RegExp("^\\[?(\\d{1,4})\\]?\\s*[.)\\]]?[\\s\\u00A0]+(\\S[\\s\\S]*)$"),

    /** Any square-bracketed span that could carry a citation. */
    bracket: /\[([^[\]\r\n]{1,160})\]/g,

    /** Single source number. */
    number: /^\d{1,4}$/,

    /** Range of sources, e.g. "3–12". */
    range: new RegExp("^(\\d{1,4})\\s*[" + DASHES + "]\\s*(\\d{1,4})$"),

    /** Page locator inside a citation: "с. 25", "С.25-30", "p. 7", "pp. 7–9". */
    pageLocator: new RegExp(
      "^(?:с|c|стр|ст|p|pp|s|seite)\\s*\\.?\\s*\\d{1,5}(?:\\s*[" + DASHES + "]\\s*\\d{1,5})?$",
      "i"
    ),

    /** Whitespace runs, including non-breaking and vertical tab. */
    whitespace: /[\s \r\n\t]+/g
  };

  /* ===========================================================================
   * 2. GENERIC UTILITIES
   * =========================================================================*/

  /** HTML-escape untrusted document text before injecting it into the pane. */
  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  /** Collapse every whitespace variant Word emits into single spaces. */
  function collapseWhitespace(value) {
    return String(value == null ? "" : value).replace(PATTERNS.whitespace, " ").trim();
  }

  /** Normalise all Unicode dashes to a plain hyphen for range parsing. */
  function normalizeDashes(value) {
    return String(value).replace(
      /[‐‑‒–—―−]/g,
      "-"
    );
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  /* --------------------------------------------------------------------------
   * Языки
   * ------------------------------------------------------------------------ */

  /**
   * Шесть букв таджикской кириллицы, отсутствующих в русском алфавите,
   * и их русские основы. Ключ — код символа (строчная и прописная формы).
   *   Ғғ U+0492/0493 · Ққ U+049A/049B · Ҳҳ U+04B2/04B3
   *   Ҷҷ U+04B6/04B7 · Ӣӣ U+04E2/04E3 · Ӯӯ U+04EE/04EF
   */
  var TAJIK_LETTERS = {
    1170: "г", 1171: "г",
    1178: "к", 1179: "к",
    1202: "х", 1203: "х",
    1206: "ч", 1207: "ч",
    1250: "и", 1251: "и",
    1262: "у", 1263: "у"
  };

  var LANGUAGE_LABELS = {
    ru: "русский",
    en: "английский",
    tj: "таджикский",
    "": "не определён"
  };

  /**
   * Свёртка строки к сравнимому виду: нижний регистр, ё→е и таджикские буквы
   * к русской основе. Позволяет держать один список ключевых слов на три языка.
   */
  function foldKey(value) {
    var source = String(value == null ? "" : value).toLowerCase();
    var out = "";

    for (var i = 0; i < source.length; i++) {
      var code = source.charCodeAt(i);
      if (code === 1105) out += "е";                       // ё
      else if (TAJIK_LETTERS[code]) out += TAJIK_LETTERS[code];
      else out += source.charAt(i);
    }
    return out;
  }

  /** Подсчёт букв по алфавитам. */
  function countScripts(text) {
    var latin = 0, cyrillic = 0, tajik = 0;
    var source = String(text == null ? "" : text);

    for (var i = 0; i < source.length; i++) {
      var code = source.charCodeAt(i);

      if ((code >= 65 && code <= 90) || (code >= 97 && code <= 122)) {
        latin++;
      } else if (code >= 0x0400 && code <= 0x04FF) {
        cyrillic++;
        if (TAJIK_LETTERS[code]) tajik++;
      }
    }
    return { latin: latin, cyrillic: cyrillic, tajik: tajik };
  }

  /**
   * Язык фрагмента: "tj" | "ru" | "en" | "".
   * Таджикский определяется по доле специфических букв — в связном тексте их
   * 3-8 %, тогда как русский текст может содержать одну-две в именах собственных.
   */
  function detectLanguage(text) {
    var counts = countScripts(text);
    var letters = counts.latin + counts.cyrillic;

    if (letters < 12) return "";

    if (counts.cyrillic > counts.latin) {
      if (counts.tajik >= 3 && counts.tajik / counts.cyrillic >= 0.012) return "tj";
      return "ru";
    }
    return "en";
  }

  function languageLabel(code) {
    return LANGUAGE_LABELS[code] || code || LANGUAGE_LABELS[""];
  }

  /** Russian plural selector: plural(5, "источник", "источника", "источников"). */
  function plural(count, one, few, many) {
    var n = Math.abs(count) % 100;
    var n1 = n % 10;
    if (n > 10 && n < 20) return many;
    if (n1 > 1 && n1 < 5) return few;
    if (n1 === 1) return one;
    return many;
  }

  /** Ascending numeric sort of a Set of integers. */
  function sortedNumbers(set) {
    var out = [];
    set.forEach(function (n) { out.push(n); });
    out.sort(function (a, b) { return a - b; });
    return out;
  }

  /** Compress [1,2,3,7,9,10] into "1–3, 7, 9–10" for compact reporting. */
  function compressRanges(numbers) {
    if (!numbers.length) return "—";
    var parts = [];
    var start = numbers[0];
    var prev = numbers[0];

    for (var i = 1; i <= numbers.length; i++) {
      var current = numbers[i];
      if (current === prev + 1) { prev = current; continue; }
      if (start === prev) parts.push(String(start));
      else if (prev === start + 1) parts.push(start + ", " + prev);
      else parts.push(start + "–" + prev);
      start = current;
      prev = current;
    }
    return parts.join(", ");
  }

  function truncate(value, limit) {
    var text = String(value == null ? "" : value);
    return text.length <= limit ? text : text.slice(0, limit - 1).trimEnd() + "…";
  }

  /* ===========================================================================
   * 3. PARSING CORE
   * ---------------------------------------------------------------------------
   * A "paragraph model" is a plain object produced by the Office.js reader:
   *   { index, text, listString, isListItem, style, effective }
   * `effective` prepends Word's auto-numbering label so that an auto-numbered
   * bibliography (where the digit is not part of paragraph.text) still parses.
   * =========================================================================*/

  /** True when a paragraph looks like the heading of the reference section. */
  function isReferenceHeading(paragraph) {
    var line = collapseWhitespace(paragraph.text);
    if (!line || line.length > 80) return false;
    return PATTERNS.referenceHeading.test(foldKey(line));
  }

  /** Match a numbered bibliography entry; returns { number, description } or null. */
  function matchReferenceEntry(paragraph) {
    var line = collapseWhitespace(paragraph.effective);
    if (!line) return null;

    var match = PATTERNS.referenceEntry.exec(line);
    if (!match) return null;

    var number = parseInt(match[1], 10);
    if (!isFinite(number) || number < 1 || number > CONFIG.maxSourceNumber) return null;

    return { number: number, description: collapseWhitespace(match[2]) };
  }

  /**
   * True when a paragraph opens a new structural block rather than continuing a
   * bibliographic record: a Word heading style, a reference-section heading, or
   * a short all-caps line (typical of a translated title block).
   */
  function isHeadingLike(paragraph) {
    if (/^(heading|заголовок)/i.test(String(paragraph.style || ""))) return true;

    var line = collapseWhitespace(paragraph.text);
    if (!line || line.length > 120) return false;
    if (PATTERNS.referenceHeading.test(foldKey(line))) return true;

    // Deliberately avoids Unicode property escapes: older Office webviews lack them.
    var letters = line.replace(/[^A-Za-zА-Яа-яЁёЇїІіҒғҚқҲҳҶҷӢӣӮӯ]/g, "");
    return letters.length >= 8 && letters === letters.toUpperCase();
  }

  /**
   * Язык короткой строки (заголовка). Порог по доле букв здесь не работает:
   * в «Рӯйхати адабиёт» специфическая буква всего одна.
   */
  function detectLanguageShort(text) {
    var counts = countScripts(text);
    if (counts.latin + counts.cyrillic === 0) return "";
    if (counts.tajik >= 1) return "tj";
    return counts.cyrillic > counts.latin ? "ru" : "en";
  }

  /**
   * Fallback detection when no heading exists: find the trailing block of
   * strictly consecutive numbered entries that terminates the document and
   * counts back to 1.
   */
  function findTrailingNumberedBlock(paragraphs) {
    var entries = [];
    for (var i = 0; i < paragraphs.length; i++) {
      var entry = matchReferenceEntry(paragraphs[i]);
      if (entry) entries.push({ index: i, number: entry.number });
    }
    if (entries.length < CONFIG.minTrailingBlock) return -1;

    var last = entries.length - 1;
    var runStart = last;
    for (var k = last - 1; k >= 0; k--) {
      if (entries[k].number === entries[runStart].number - 1) runStart = k;
      else break;
    }

    if (entries[runStart].number !== 1) return -1;
    if (last - runStart + 1 < CONFIG.minTrailingBlock) return -1;

    return entries[runStart].index;
  }

  /* --------------------------------------------------------------------------
   * Раскладка документа
   * ------------------------------------------------------------------------
   * Многоязычная статья устроена как чередование блоков:
   *     [рус. текст][Список литературы][англ. текст][References][тадж. текст][Адабиёт]
   * Поэтому документ разбирается не на «текст + список», а на произвольную
   * последовательность секций, каждой из которых присваивается язык.
   * ---------------------------------------------------------------------- */

  function joinSectionText(paragraphs, from, to) {
    var parts = [];
    for (var i = from; i <= to && i < paragraphs.length; i++) {
      if (i >= 0) parts.push(paragraphs[i].text);
    }
    return parts.join(" ");
  }

  function makeBodySection(paragraphs, from, to) {
    return {
      kind: "body",
      from: from,
      to: to,
      paragraphCount: to - from + 1,
      language: detectLanguage(joinSectionText(paragraphs, from, to))
    };
  }

  /**
   * Полная раскладка документа.
   * @returns {{ sections, method, headingIndex, paraMeta, bodyParagraphs }}
   */
  function analyseLayout(paragraphs) {
    var total = paragraphs.length;
    var sections = [];
    var method = "none";
    var firstHeading = -1;
    var cursor = 0;
    var i;

    var headings = [];
    for (i = 0; i < total; i++) {
      if (isReferenceHeading(paragraphs[i])) headings.push(i);
    }

    function pushReferenceSection(headingIdx, firstEntryIdx, parsed, lastIdx) {
      var headingText = headingIdx >= 0 ? paragraphs[headingIdx].text : "";
      var language = headingIdx >= 0 ? detectLanguageShort(headingText) : "";

      if (!language) {
        // Заголовка нет или он невыразителен — берём язык предыдущего текста.
        for (var k = sections.length - 1; k >= 0; k--) {
          if (sections[k].kind === "body" && sections[k].language) {
            language = sections[k].language;
            break;
          }
        }
      }
      if (!language) language = detectLanguage(joinSectionText(paragraphs, firstEntryIdx, lastIdx));

      sections.push({
        kind: "references",
        from: headingIdx >= 0 ? headingIdx : firstEntryIdx,
        to: lastIdx,
        headingIndex: headingIdx,
        headingText: collapseWhitespace(headingText),
        paragraphCount: parsed.consumed,
        language: language,
        parsed: parsed
      });
    }

    if (headings.length) {
      for (var h = 0; h < headings.length; h++) {
        var idx = headings[h];
        if (idx < cursor) continue;                 // заголовок внутри уже разобранного списка

        var parsed = parseReferenceList(paragraphs.slice(idx + 1));
        if (!parsed.entries.length) continue;       // ложное срабатывание — это обычный абзац

        if (idx > cursor) sections.push(makeBodySection(paragraphs, cursor, idx - 1));

        var lastIdx = idx + parsed.consumed;
        pushReferenceSection(idx, idx + 1, parsed, lastIdx);
        cursor = lastIdx + 1;
        if (firstHeading < 0) firstHeading = idx;
        method = "heading";
      }

      if (cursor < total) sections.push(makeBodySection(paragraphs, cursor, total - 1));
    }

    if (method === "none") {
      sections.length = 0;
      var fallback = findTrailingNumberedBlock(paragraphs);

      if (fallback >= 0) {
        var tailParsed = parseReferenceList(paragraphs.slice(fallback));
        if (tailParsed.entries.length) {
          method = "trailing-block";
          if (fallback > 0) sections.push(makeBodySection(paragraphs, 0, fallback - 1));
          var tailLast = fallback + tailParsed.consumed - 1;
          pushReferenceSection(-1, fallback, tailParsed, tailLast);
          if (tailLast + 1 < total) sections.push(makeBodySection(paragraphs, tailLast + 1, total - 1));
        }
      }

      if (method === "none") sections.push(makeBodySection(paragraphs, 0, total - 1));
    }

    // Карта «абзац → секция» и плоский список абзацев основного текста.
    var paraMeta = {};
    var bodyParagraphs = [];

    for (var s = 0; s < sections.length; s++) {
      var section = sections[s];
      if (section.kind !== "body") continue;

      for (var p = section.from; p <= section.to; p++) {
        if (p < 0 || p >= total) continue;
        paraMeta[paragraphs[p].index] = { language: section.language, sectionIndex: s };
        bodyParagraphs.push(paragraphs[p]);
      }
    }

    return {
      sections: sections,
      method: method,
      headingIndex: firstHeading,
      paraMeta: paraMeta,
      bodyParagraphs: bodyParagraphs
    };
  }

  /**
   * Split the document into narrative body and reference section.
   * Сохранён для обратной совместимости и точечных проверок.
   * @returns {{ bodyParagraphs, referenceParagraphs, headingIndex, method }}
   */
  function splitDocument(paragraphs) {
    var headingIndex = -1;

    // Search from the end: a document may quote the words "список литературы"
    // in its introduction, but the real section is the last such heading.
    for (var i = paragraphs.length - 1; i >= 0; i--) {
      if (isReferenceHeading(paragraphs[i])) { headingIndex = i; break; }
    }

    var method = "heading";
    var sectionStart;

    if (headingIndex >= 0) {
      sectionStart = headingIndex + 1;
    } else {
      var fallback = findTrailingNumberedBlock(paragraphs);
      if (fallback < 0) {
        return {
          bodyParagraphs: paragraphs.slice(),
          referenceParagraphs: [],
          headingIndex: -1,
          method: "none"
        };
      }
      sectionStart = fallback;
      method = "trailing-block";
    }

    return {
      bodyParagraphs: paragraphs.slice(0, headingIndex >= 0 ? headingIndex : sectionStart),
      referenceParagraphs: paragraphs.slice(sectionStart),
      headingIndex: headingIndex,
      method: method
    };
  }

  /**
   * Turn reference paragraphs into structured entries, folding wrapped
   * continuation lines into the preceding entry.
   *
   * The list is terminated as soon as the text stops looking like a bibliography
   * (a heading, or too many consecutive non-entry paragraphs). `consumed` reports
   * how many paragraphs actually belong to the list so the caller can return the
   * remainder — e.g. the second-language version of a bilingual article — to the
   * body scan instead of silently discarding it.
   */
  function parseReferenceList(referenceParagraphs) {
    var entries = [];
    var byNumber = new Map();
    var duplicates = [];
    var unnumbered = [];
    var current = null;
    var consumed = 0;
    var continuationRun = 0;

    for (var i = 0; i < referenceParagraphs.length; i++) {
      var paragraph = referenceParagraphs[i];
      var line = collapseWhitespace(paragraph.effective);
      if (!line) continue;                       // blank lines never terminate the list

      var match = matchReferenceEntry(paragraph);

      if (match) {
        current = {
          number: match.number,
          description: match.description,
          paragraphIndex: paragraph.index
        };
        entries.push(current);

        if (byNumber.has(match.number)) duplicates.push(match.number);
        else byNumber.set(match.number, current);

        consumed = i + 1;
        continuationRun = 0;
        continue;
      }

      if (current) {
        // A new structural block ends the bibliography outright.
        if (isHeadingLike(paragraph)) break;

        continuationRun++;
        if (continuationRun > CONFIG.maxContinuationParagraphs) break;

        // Wrapped line of the previous entry (common with long GOST descriptions).
        current.description = collapseWhitespace(current.description + " " + line);
        consumed = i + 1;
        continue;
      }

      // Text between the heading and the first entry, or an unnumbered entry.
      unnumbered.push({ paragraphIndex: paragraph.index, text: truncate(line, 160) });
      consumed = i + 1;
    }

    entries.forEach(function (entry) {
      entry.description = truncate(entry.description, CONFIG.maxDescriptionLength);
    });

    // Numbering gaps: declared 1..max but some ordinals never appear.
    var declared = sortedNumbers(new Set(entries.map(function (e) { return e.number; })));
    var gaps = [];
    if (declared.length) {
      for (var n = 1; n <= declared[declared.length - 1]; n++) {
        if (!byNumber.has(n)) gaps.push(n);
      }
    }

    return {
      entries: entries,
      byNumber: byNumber,
      duplicates: sortedNumbers(new Set(duplicates)),
      unnumbered: unnumbered,
      gaps: gaps,
      consumed: consumed
    };
  }

  /**
   * Parse the inner text of one bracket span.
   * @returns {{ isCitation, numbers, ranges, malformed, pageLocators, foreignTokens }}
   */
  function parseBracketContent(rawContent) {
    var content = normalizeDashes(collapseWhitespace(rawContent));
    var tokens = content.split(/[;,]/);

    var numbers = new Set();
    var ranges = [];
    var malformed = [];
    var pageLocators = [];
    var foreignTokens = [];
    var numericTokenCount = 0;

    for (var i = 0; i < tokens.length; i++) {
      var token = tokens[i].trim();
      if (!token) continue;

      // --- plain number ---------------------------------------------------
      if (PATTERNS.number.test(token)) {
        var value = parseInt(token, 10);
        if (value >= 1 && value <= CONFIG.maxSourceNumber) {
          numbers.add(value);
          numericTokenCount++;
        } else {
          malformed.push({ token: token, reason: "номер вне допустимого диапазона" });
        }
        continue;
      }

      // --- range ----------------------------------------------------------
      var rangeMatch = PATTERNS.range.exec(token);
      if (rangeMatch) {
        var from = parseInt(rangeMatch[1], 10);
        var to = parseInt(rangeMatch[2], 10);

        if (to < from) {
          malformed.push({ token: token, reason: "конец диапазона меньше начала" });
        } else if (to - from + 1 > CONFIG.maxRangeExpansion) {
          malformed.push({ token: token, reason: "диапазон превышает лимит развёртывания" });
        } else if (to > CONFIG.maxSourceNumber) {
          malformed.push({ token: token, reason: "номер вне допустимого диапазона" });
        } else {
          for (var v = from; v <= to; v++) numbers.add(v);
          ranges.push({ from: from, to: to, size: to - from + 1 });
          numericTokenCount++;
        }
        continue;
      }

      // --- page locator (ignored, but proves this is a citation) ----------
      if (PATTERNS.pageLocator.test(token)) {
        pageLocators.push(token);
        continue;
      }

      // --- anything else --------------------------------------------------
      foreignTokens.push(token);
    }

    return {
      isCitation: numericTokenCount > 0,
      numbers: numbers,
      ranges: ranges,
      malformed: malformed,
      pageLocators: pageLocators,
      foreignTokens: foreignTokens
    };
  }

  /** Build a readable context snippet around a bracket occurrence. */
  function buildSnippet(paragraphText, matchIndex, matchLength) {
    var radius = CONFIG.snippetRadius;
    var start = Math.max(0, matchIndex - radius);
    var end = Math.min(paragraphText.length, matchIndex + matchLength + radius);
    var snippet = collapseWhitespace(paragraphText.slice(start, end));
    if (start > 0) snippet = "…" + snippet;
    if (end < paragraphText.length) snippet = snippet + "…";
    return snippet;
  }

  /**
   * Scan paragraphs for bracketed citations.
   * @returns {{ citations, ignoredBrackets, occurrenceCount, wordCount }}
   */
  function collectCitations(paragraphs, bulkThreshold, paraMeta) {
    var citations = [];
    var ignoredBrackets = [];
    var wordCount = 0;
    var wordsByLanguage = {};

    for (var i = 0; i < paragraphs.length; i++) {
      var paragraph = paragraphs[i];
      var text = paragraph.text || "";
      var normalized = collapseWhitespace(text);
      var meta = (paraMeta && paraMeta[paragraph.index]) || null;
      var language = meta ? meta.language : "";

      if (normalized) {
        var words = normalized.split(" ").filter(Boolean).length;
        wordCount += words;
        wordsByLanguage[language] = (wordsByLanguage[language] || 0) + words;
      }

      PATTERNS.bracket.lastIndex = 0;
      var match;

      while ((match = PATTERNS.bracket.exec(text)) !== null) {
        var parsed = parseBracketContent(match[1]);
        var record = {
          raw: "[" + collapseWhitespace(match[1]) + "]",
          paragraphIndex: paragraph.index,
          charIndex: match.index,
          language: language,
          snippet: buildSnippet(text, match.index, match[0].length)
        };

        if (!parsed.isCitation) {
          record.reason = parsed.foreignTokens.length
            ? "нечисловое содержимое"
            : "пустая скобка";
          ignoredBrackets.push(record);
          continue;
        }

        record.numbers = sortedNumbers(parsed.numbers);
        record.ranges = parsed.ranges;
        record.malformed = parsed.malformed;
        record.pageLocators = parsed.pageLocators;
        record.foreignTokens = parsed.foreignTokens;
        record.size = record.numbers.length;
        record.isBulk = record.size > bulkThreshold;
        record.widestRange = parsed.ranges.reduce(function (acc, r) {
          return !acc || r.size > acc.size ? r : acc;
        }, null);

        citations.push(record);
      }
    }

    return {
      citations: citations,
      ignoredBrackets: ignoredBrackets,
      occurrenceCount: citations.length,
      wordCount: wordCount,
      wordsByLanguage: wordsByLanguage
    };
  }

  /* ===========================================================================
   * 4. AUDIT ENGINE
   * =========================================================================*/

  /** Складывает записи одной секции в объединённый список своего языка. */
  function appendEntries(list, unionByNumber, entries, language) {
    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      var record = {
        number: entry.number,
        description: entry.description,
        paragraphIndex: entry.paragraphIndex,
        language: language
      };

      list.entries.push(record);
      if (list.byNumber.has(entry.number)) list.duplicates.push(entry.number);
      else list.byNumber.set(entry.number, record);
      if (!unionByNumber.has(entry.number)) unionByNumber.set(entry.number, record);
    }
  }

  /** Пропуски нумерации и дедупликация служебных массивов списка. */
  function finaliseList(list) {
    var max = 0;
    list.byNumber.forEach(function (value, number) { if (number > max) max = number; });

    for (var n = 1; n <= max; n++) {
      if (!list.byNumber.has(n)) list.gaps.push(n);
    }

    list.duplicates = sortedNumbers(new Set(list.duplicates));
    list.maxNumber = max;
    list.deadCount = 0;
  }

  /* --------------------------------------------------------------------------
   * Библиометрический профиль списка литературы
   * ------------------------------------------------------------------------
   * Год издания, наличие DOI и язык записи извлекаются из самого описания.
   * Индексация в Scopus и Web of Science из текста НЕ определяется — это
   * свойство внешних баз. Считается только явное упоминание в записи.
   * ---------------------------------------------------------------------- */

  var BIB = {
    /** Год: 1800..текущий+1, проверяется отдельно. */
    year: /\b(1[89]\d{2}|20\d{2}|21\d{2})\b/g,
    /** DOI по стандарту: 10.<регистрант>/<суффикс>. */
    doi: /\b10\.\d{4,9}\/\S{1,200}/,
    doiWord: /\bdoi\b/i,
    /** Явные пометки индексации, проставленные автором. */
    scopus: /(scopus|скопус)/i,
    wos: /(web\s*of\s*science|\bwos\b|веб\s*оф\s*сайенс)/i,
    /** Ссылки и идентификаторы вычищаются перед определением языка. */
    noise: /(https?:\/\/\S+|www\.\S+|10\.\d{4,9}\/\S+|[A-Za-z-]*\d{4}-\d{3}[\dXx])/g
  };

  /** Год издания записи: наибольший правдоподобный год в описании. */
  function extractYear(description, currentYear) {
    var text = String(description || "");
    var best = 0;
    var match;

    BIB.year.lastIndex = 0;
    while ((match = BIB.year.exec(text)) !== null) {
      var year = parseInt(match[1], 10);
      if (year >= 1800 && year <= currentYear + 1 && year > best) best = year;
    }
    return best || null;
  }

  /** Язык библиографической записи по преобладающему алфавиту. */
  function entryLanguage(description) {
    var clean = String(description || "").replace(BIB.noise, " ");
    var counts = countScripts(clean);

    if (counts.latin + counts.cyrillic < 6) return "";
    if (counts.cyrillic > counts.latin) {
      return (counts.tajik >= 2 && counts.tajik / counts.cyrillic >= 0.012) ? "tj" : "ru";
    }
    return "en";
  }

  /**
   * Корзины по годам: последние пять лет поштучно, дальше пятилетиями.
   * @returns {Array<{label,from,to,count}>}
   */
  function buildYearBuckets(years, currentYear) {
    var buckets = [];
    var i;

    for (i = 0; i < 5; i++) {
      var single = currentYear - i;
      buckets.push({ label: String(single), from: single, to: single, count: 0 });
    }

    var oldest = currentYear;
    for (i = 0; i < years.length; i++) if (years[i] < oldest) oldest = years[i];

    var edge = currentYear - 5;
    var guard = 0;
    while (edge >= oldest && guard < 10) {
      var from = edge - 4;
      buckets.push({ label: from + "–" + edge, from: from, to: edge, count: 0 });
      edge = from - 1;
      guard++;
    }

    if (oldest <= edge) {
      buckets.push({ label: "до " + (edge + 1), from: -1, to: edge, count: 0 });
    }

    for (i = 0; i < years.length; i++) {
      for (var b = 0; b < buckets.length; b++) {
        var bucket = buckets[b];
        var lower = bucket.from === -1 ? -Infinity : bucket.from;
        if (years[i] >= lower && years[i] <= bucket.to) { bucket.count++; break; }
      }
    }

    return buckets;
  }

  /** Полный библиометрический разбор всех записей всех списков. */
  function analyseBibliometrics(entries, currentYear) {
    var years = [];
    var withDoi = 0;
    var byLanguage = { ru: 0, en: 0, tj: 0, "": 0 };
    var englishWithDoi = 0;
    var scopusMentions = 0;
    var wosMentions = 0;
    var noYear = [];
    var noDoi = [];

    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      var text = entry.description || "";

      var year = extractYear(text, currentYear);
      entry.year = year;
      if (year) years.push(year); else noYear.push(entry);

      var hasDoi = BIB.doi.test(text) || (BIB.doiWord.test(text) && /10\.\d{4}/.test(text));
      entry.hasDoi = hasDoi;
      if (hasDoi) withDoi++; else noDoi.push(entry);

      var lang = entryLanguage(text);
      entry.entryLanguage = lang;
      byLanguage[lang] = (byLanguage[lang] || 0) + 1;
      if (lang === "en" && hasDoi) englishWithDoi++;

      if (BIB.scopus.test(text)) scopusMentions++;
      if (BIB.wos.test(text)) wosMentions++;
    }

    years.sort(function (a, b) { return a - b; });

    var total = entries.length;
    var recent5 = 0;
    var recent10 = 0;
    for (var k = 0; k < years.length; k++) {
      if (years[k] >= currentYear - 4) recent5++;
      if (years[k] >= currentYear - 9) recent10++;
    }

    var median = null;
    if (years.length) {
      var mid = Math.floor(years.length / 2);
      median = years.length % 2 ? years[mid] : Math.round((years[mid - 1] + years[mid]) / 2);
    }

    return {
      currentYear: currentYear,
      total: total,
      buckets: buildYearBuckets(years, currentYear),
      datedCount: years.length,
      undatedCount: total - years.length,
      oldest: years.length ? years[0] : null,
      newest: years.length ? years[years.length - 1] : null,
      medianYear: median,
      recent5: recent5,
      recent10: recent10,
      recent5Share: total ? recent5 / total : 0,
      recent10Share: total ? recent10 / total : 0,
      doiCount: withDoi,
      doiShare: total ? withDoi / total : 0,
      englishCount: byLanguage.en || 0,
      englishShare: total ? (byLanguage.en || 0) / total : 0,
      englishWithDoi: englishWithDoi,
      russianCount: byLanguage.ru || 0,
      tajikCount: byLanguage.tj || 0,
      undeterminedLanguage: byLanguage[""] || 0,
      scopusMentions: scopusMentions,
      wosMentions: wosMentions,
      undatedSamples: noYear.slice(0, 5).map(function (e) {
        return { number: e.number, description: truncate(e.description, 120) };
      }),
      withoutDoiSamples: noDoi.slice(0, 5).map(function (e) {
        return { number: e.number, description: truncate(e.description, 120) };
      })
    };
  }

  /** Detect numbers whose first mention breaks ascending GOST ordering. */
  function analyseMentionOrder(citations) {
    var seen = new Set();
    var firstOrder = [];

    citations.forEach(function (citation) {
      citation.numbers.forEach(function (n) {
        if (!seen.has(n)) {
          seen.add(n);
          firstOrder.push({ number: n, paragraphIndex: citation.paragraphIndex, raw: citation.raw });
        }
      });
    });

    var violations = [];
    var highest = 0;
    firstOrder.forEach(function (item) {
      if (item.number < highest) {
        violations.push({
          number: item.number,
          afterNumber: highest,
          paragraphIndex: item.paragraphIndex,
          raw: item.raw
        });
      } else {
        highest = item.number;
      }
    });

    return { firstOrder: firstOrder, violations: violations };
  }

  /** Compute the composite integrity score and its verdict band. */
  function buildVerdict(metrics) {
    var penalties = [];
    var score = 100;

    var deadPenalty = Math.round(metrics.deadRatio * 50);
    if (deadPenalty > 0) {
      score -= deadPenalty;
      penalties.push({
        level: metrics.deadRatio > 0.25 ? "danger" : "warning",
        text: "Мёртвые источники: " + metrics.deadCount + " из " + metrics.totalReferences +
              " (" + Math.round(metrics.deadRatio * 100) + "%). Штраф −" + deadPenalty + "."
      });
    }

    var phantomPenalty = Math.min(30, metrics.phantomCount * 6);
    if (phantomPenalty > 0) {
      score -= phantomPenalty;
      penalties.push({
        level: "danger",
        text: "Фантомные ссылки: " + metrics.phantomCount +
              ". Нарушение прослеживаемости источника. Штраф −" + phantomPenalty + "."
      });
    }

    var bulkPenalty = Math.min(20, metrics.bulkCount * 4);
    if (bulkPenalty > 0) {
      score -= bulkPenalty;
      penalties.push({
        level: "warning",
        text: "Пакетные интервалы: " + metrics.bulkCount +
              ". Снижение адресности аргументации. Штраф −" + bulkPenalty + "."
      });
    }

    var densityPenalty = 0;
    if (metrics.wordCount >= 300) {
      if (metrics.citationsPer1000Words < 2) densityPenalty = 15;
      else if (metrics.citationsPer1000Words < 4) densityPenalty = 8;
      else if (metrics.citationsPer1000Words < 6) densityPenalty = 3;
    }
    if (densityPenalty > 0) {
      score -= densityPenalty;
      penalties.push({
        level: densityPenalty >= 15 ? "danger" : "warning",
        text: "Плотность цитирования " + metrics.citationsPer1000Words.toFixed(1) +
              " на 1000 слов — ниже нормы обзорного текста. Штраф −" + densityPenalty + "."
      });
    }

    var crossPenalty = Math.min(15, (metrics.crossLanguageIssues || 0) * 5);
    if (crossPenalty > 0) {
      score -= crossPenalty;
      penalties.push({
        level: "warning",
        text: "Расхождения между языковыми версиями: " + metrics.crossLanguageIssues +
              ". Штраф −" + crossPenalty + "."
      });
    }

    var duplicatePenalty = Math.min(10, metrics.duplicateCount * 3);
    if (duplicatePenalty > 0) {
      score -= duplicatePenalty;
      penalties.push({
        level: "warning",
        text: "Дублирующиеся номера в списке: " + metrics.duplicateCount +
              ". Штраф −" + duplicatePenalty + "."
      });
    }

    score = clamp(Math.round(score), 0, 100);

    var band;
    if (score >= 85) band = { key: "ok", label: "Низкий риск", caption: "Аппарат цитирования согласован" };
    else if (score >= 60) band = { key: "warn", label: "Умеренный риск", caption: "Требуется точечная правка" };
    else band = { key: "risk", label: "Высокий риск", caption: "Аппарат цитирования недостоверен" };

    return { score: score, band: band, penalties: penalties };
  }

  /**
   * Full audit. Pure function over paragraph models.
   * @param {Array} paragraphs paragraph models
   * @param {{ bulkThreshold:number, scanWholeDocument:boolean, checkOrder:boolean }} options
   */
  function auditDocument(paragraphs, options) {
    var bulkThreshold = options.bulkThreshold;
    var layout = analyseLayout(paragraphs);
    var i, k;

    // ---- списки литературы, сгруппированные по языку ----------------------
    var lists = {};
    var listOrder = [];
    var unionByNumber = new Map();
    var referenceSections = [];
    var referenceParagraphCount = 0;

    for (i = 0; i < layout.sections.length; i++) {
      var section = layout.sections[i];
      if (section.kind !== "references") continue;

      referenceSections.push(section);
      referenceParagraphCount += section.parsed.consumed;

      var lang = section.language || "";
      if (!lists[lang]) {
        lists[lang] = {
          language: lang,
          entries: [],
          byNumber: new Map(),
          duplicates: [],
          gaps: [],
          sectionCount: 0,
          headings: [],
          unnumbered: []
        };
        listOrder.push(lang);
      }

      var list = lists[lang];
      list.sectionCount++;
      list.unnumbered = list.unnumbered.concat(section.parsed.unnumbered);
      if (section.headingText) list.headings.push(section.headingText);

      appendEntries(list, unionByNumber, section.parsed.entries, lang);
    }

    for (i = 0; i < listOrder.length; i++) finaliseList(lists[listOrder[i]]);

    // ---- сканирование основного текста -----------------------------------
    var scanScope = options.scanWholeDocument ? paragraphs : layout.bodyParagraphs;
    var scan = collectCitations(scanScope, bulkThreshold, layout.paraMeta);

    // Если список в документе один, он обслуживает все языковые сегменты:
    // русский перечень литературы под английским переводом статьи — норма.
    var singleList = listOrder.length === 1 ? listOrder[0] : null;

    function resolveTarget(language) {
      if (language && lists[language]) return language;
      if (singleList !== null) return singleList;
      return "__union__";
    }

    function numbersOf(target) {
      return target === "__union__" ? unionByNumber : lists[target].byNumber;
    }

    // ---- множества цитируемого -------------------------------------------
    var citedSet = new Set();
    var frequency = new Map();
    var citedByTarget = {};
    var citedByLanguage = {};
    var citationsByLanguage = {};

    scan.citations.forEach(function (citation) {
      var target = resolveTarget(citation.language);
      var langKey = citation.language || "";
      citation.target = target;

      if (!citedByTarget[target]) citedByTarget[target] = new Set();
      if (!citedByLanguage[langKey]) citedByLanguage[langKey] = new Set();
      citationsByLanguage[langKey] = (citationsByLanguage[langKey] || 0) + 1;

      citation.numbers.forEach(function (n) {
        citedSet.add(n);
        frequency.set(n, (frequency.get(n) || 0) + 1);
        citedByTarget[target].add(n);
        citedByLanguage[langKey].add(n);
      });
    });
    var citedNumbers = sortedNumbers(citedSet);
    var unionCited = citedByTarget["__union__"] || new Set();

    // ---- раздел 2: мёртвые источники --------------------------------------
    var deadSources = [];

    for (i = 0; i < listOrder.length; i++) {
      var deadLang = listOrder[i];
      var deadList = lists[deadLang];
      var citedHere = citedByTarget[deadLang] || new Set();
      deadList.deadCount = 0;

      deadList.entries.forEach(function (entry) {
        if (citedHere.has(entry.number) || unionCited.has(entry.number)) return;
        deadList.deadCount++;
        deadSources.push({
          number: entry.number,
          description: entry.description,
          paragraphIndex: entry.paragraphIndex,
          language: deadLang,
          languageLabel: languageLabel(deadLang)
        });
      });
    }

    deadSources.sort(function (a, b) {
      if (a.language === b.language) return a.number - b.number;
      return a.language < b.language ? -1 : 1;
    });

    // ---- раздел 3: фантомные ссылки ---------------------------------------
    var phantomIndex = {};
    var phantomReferences = [];

    scan.citations.forEach(function (citation) {
      var byNumber = numbersOf(citation.target);
      var langKey = citation.language || "";

      citation.numbers.forEach(function (n) {
        if (byNumber.has(n)) return;

        var key = langKey + "|" + n;
        if (!phantomIndex[key]) {
          phantomIndex[key] = {
            number: n,
            language: langKey,
            languageLabel: languageLabel(langKey),
            occurrences: [],
            count: 0
          };
          phantomReferences.push(phantomIndex[key]);
        }
        phantomIndex[key].occurrences.push({
          raw: citation.raw,
          paragraphIndex: citation.paragraphIndex,
          snippet: citation.snippet
        });
        phantomIndex[key].count++;
      });
    });

    phantomReferences.sort(function (a, b) {
      if (a.language === b.language) return a.number - b.number;
      return a.language < b.language ? -1 : 1;
    });

    // ---- section 4: bulk intervals ---------------------------------------
    var bulkCitations = scan.citations
      .filter(function (c) { return c.isBulk; })
      .map(function (c) {
        return {
          raw: c.raw,
          size: c.size,
          numbers: c.numbers,
          widestRange: c.widestRange,
          paragraphIndex: c.paragraphIndex,
          snippet: c.snippet
        };
      })
      .sort(function (a, b) { return b.size - a.size; });

    // ---- diagnostics -----------------------------------------------------
    var malformedTokens = [];
    scan.citations.forEach(function (c) {
      c.malformed.forEach(function (m) {
        malformedTokens.push({
          raw: c.raw,
          token: m.token,
          reason: m.reason,
          paragraphIndex: c.paragraphIndex
        });
      });
    });

    var order = options.checkOrder
      ? analyseMentionOrder(scan.citations)
      : { firstOrder: [], violations: [] };

    // ---- библиометрический профиль -----------------------------------------
    var allEntries = [];
    for (i = 0; i < listOrder.length; i++) {
      allEntries = allEntries.concat(lists[listOrder[i]].entries);
    }
    var bibliometrics = analyseBibliometrics(allEntries, new Date().getFullYear());

    // ---- языковой разрез ---------------------------------------------------
    var languageKeys = {};
    for (i = 0; i < layout.sections.length; i++) languageKeys[layout.sections[i].language || ""] = true;
    for (i = 0; i < listOrder.length; i++) languageKeys[listOrder[i]] = true;

    var languages = [];
    for (var code in languageKeys) {
      if (!Object.prototype.hasOwnProperty.call(languageKeys, code)) continue;

      var bodyParas = 0;
      for (k = 0; k < layout.sections.length; k++) {
        var bodySection = layout.sections[k];
        if (bodySection.kind === "body" && (bodySection.language || "") === code) {
          bodyParas += bodySection.paragraphCount;
        }
      }

      var phantomHere = 0;
      for (k = 0; k < phantomReferences.length; k++) {
        if (phantomReferences[k].language === code) phantomHere++;
      }

      var langList = lists[code];
      languages.push({
        code: code,
        label: languageLabel(code),
        bodyParagraphs: bodyParas,
        words: scan.wordsByLanguage[code] || 0,
        citations: citationsByLanguage[code] || 0,
        uniqueCited: citedByLanguage[code] ? citedByLanguage[code].size : 0,
        referenceEntries: langList ? langList.entries.length : 0,
        hasList: !!langList,
        deadCount: langList ? langList.deadCount : 0,
        phantomCount: phantomHere
      });
    }
    languages.sort(function (a, b) { return b.words - a.words; });

    // ---- межъязыковая сверка -----------------------------------------------
    var crossLanguage = { lists: [], issues: [] };

    for (i = 0; i < listOrder.length; i++) {
      var summaryList = lists[listOrder[i]];
      crossLanguage.lists.push({
        language: summaryList.language,
        label: languageLabel(summaryList.language),
        count: summaryList.entries.length,
        maxNumber: summaryList.maxNumber,
        sections: summaryList.sectionCount,
        headings: summaryList.headings
      });
    }

    if (crossLanguage.lists.length > 1) {
      var base = crossLanguage.lists[0];
      for (i = 1; i < crossLanguage.lists.length; i++) {
        var other = crossLanguage.lists[i];
        if (other.count !== base.count) {
          crossLanguage.issues.push({
            level: "danger",
            text: "Списки различаются по объёму: " + base.label + " — " + base.count + " " +
                  plural(base.count, "запись", "записи", "записей") + ", " +
                  other.label + " — " + other.count + ". В многоязычной статье " +
                  "перечни источников должны совпадать."
          });
        } else if (other.maxNumber !== base.maxNumber) {
          crossLanguage.issues.push({
            level: "warning",
            text: "Совпадает число записей, но расходится нумерация: " + base.label +
                  " доходит до " + base.maxNumber + ", " + other.label + " — до " + other.maxNumber + "."
          });
        }
      }
    }

    for (i = 0; i < languages.length; i++) {
      var langInfo = languages[i];
      if (!langInfo.code || langInfo.bodyParagraphs === 0 || langInfo.words < 100) continue;

      if (!langInfo.hasList && listOrder.length > 1) {
        crossLanguage.issues.push({
          level: "warning",
          text: "У раздела на " + langInfo.label + " языке (" + langInfo.words + " " +
                plural(langInfo.words, "слово", "слова", "слов") +
                ") нет собственного списка литературы."
        });
      }

      if (langInfo.citations === 0) {
        crossLanguage.issues.push({
          level: "danger",
          text: "Раздел на " + langInfo.label + " языке (" + langInfo.words + " " +
                plural(langInfo.words, "слово", "слова", "слов") +
                ") не содержит ни одной внутритекстовой ссылки."
        });
      }
    }

    crossLanguage.consistent = crossLanguage.issues.length === 0;

    // ---- сводные показатели ------------------------------------------------
    var totalReferences = 0;
    var duplicateNumbers = [];
    var numberingGaps = [];
    var unnumberedLines = [];

    for (i = 0; i < listOrder.length; i++) {
      var aggregate = lists[listOrder[i]];
      totalReferences += aggregate.entries.length;
      duplicateNumbers = duplicateNumbers.concat(aggregate.duplicates);
      numberingGaps = numberingGaps.concat(aggregate.gaps);
      unnumberedLines = unnumberedLines.concat(aggregate.unnumbered);
    }
    duplicateNumbers = sortedNumbers(new Set(duplicateNumbers));
    numberingGaps = sortedNumbers(new Set(numberingGaps));

    var resolvedCited = totalReferences - deadSources.length;

    var usedNumbers = 0;
    var usedTotal = 0;
    citedNumbers.forEach(function (n) {
      if (!unionByNumber.has(n)) return;
      usedNumbers++;
      usedTotal += frequency.get(n) || 0;
    });

    var tailParagraphCount = 0;
    var passedReferences = false;
    for (i = 0; i < layout.sections.length; i++) {
      if (layout.sections[i].kind === "references") { passedReferences = true; continue; }
      if (passedReferences) tailParagraphCount += layout.sections[i].paragraphCount;
    }

    var metrics = {
      totalReferences: totalReferences,
      referenceListCount: listOrder.length,
      uniqueCited: citedNumbers.length,
      resolvedCited: resolvedCited,
      occurrenceCount: scan.occurrenceCount,
      wordCount: scan.wordCount,
      deadCount: deadSources.length,
      phantomCount: phantomReferences.length,
      bulkCount: bulkCitations.length,
      duplicateCount: duplicateNumbers.length,
      languageCount: crossLanguage.lists.length,
      crossLanguageIssues: crossLanguage.issues.length,
      recent5Share: bibliometrics.recent5Share,
      doiShare: bibliometrics.doiShare,
      englishCount: bibliometrics.englishCount,
      deadRatio: totalReferences ? deadSources.length / totalReferences : 0,
      coverage: totalReferences ? resolvedCited / totalReferences : 0,
      citationsPer1000Words: scan.wordCount ? (scan.occurrenceCount / scan.wordCount) * 1000 : 0,
      averageUsesPerSource: usedNumbers ? usedTotal / usedNumbers : 0
    };

    var verdict = buildVerdict(metrics);

    var mostCited = citedNumbers
      .map(function (n) { return { number: n, count: frequency.get(n) || 0 }; })
      .sort(function (a, b) { return b.count - a.count || a.number - b.number; })
      .slice(0, 5);

    return {
      generatedAt: new Date().toISOString(),
      options: {
        bulkThreshold: bulkThreshold,
        scanWholeDocument: !!options.scanWholeDocument,
        checkOrder: !!options.checkOrder
      },
      detection: {
        method: layout.method,
        headingParagraphIndex: layout.headingIndex,
        bodyParagraphCount: layout.bodyParagraphs.length,
        referenceParagraphCount: referenceParagraphCount,
        tailParagraphCount: tailParagraphCount,
        sectionCount: layout.sections.length,
        referenceListCount: listOrder.length
      },
      layout: layout.sections.map(function (s) {
        return {
          kind: s.kind,
          from: s.from,
          to: s.to,
          language: s.language || "",
          languageLabel: languageLabel(s.language),
          paragraphCount: s.paragraphCount,
          heading: s.headingText || "",
          entryCount: s.parsed ? s.parsed.entries.length : 0
        };
      }),
      languages: languages,
      crossLanguage: crossLanguage,
      bibliometrics: bibliometrics,
      metrics: metrics,
      citedNumbers: citedNumbers,
      citedCompressed: compressRanges(citedNumbers),
      mostCited: mostCited,
      deadSources: deadSources,
      phantomReferences: phantomReferences,
      bulkCitations: bulkCitations,
      diagnostics: {
        ignoredBrackets: scan.ignoredBrackets,
        malformedTokens: malformedTokens,
        duplicateNumbers: duplicateNumbers,
        numberingGaps: numberingGaps,
        unnumberedLines: unnumberedLines,
        orderViolations: order.violations
      },
      verdict: verdict
    };
  }

  /* ===========================================================================
   * 5. REPORT SERIALISERS
   * =========================================================================*/

  /** Human-readable prose for section 5. */
  function buildConclusionSentences(report) {
    var m = report.metrics;
    var lines = [];

    lines.push(
      "Список литературы содержит " + m.totalReferences + " " +
      plural(m.totalReferences, "запись", "записи", "записей") +
      "; в тексте задействовано " + m.resolvedCited +
      " (" + Math.round(m.coverage * 100) + "% охвата). " +
      "Всего зафиксировано " + m.occurrenceCount + " " +
      plural(m.occurrenceCount, "внутритекстовая ссылка", "внутритекстовые ссылки", "внутритекстовых ссылок") +
      " при объёме основного текста " + m.wordCount + " " +
      plural(m.wordCount, "слово", "слова", "слов") +
      " (" + m.citationsPer1000Words.toFixed(1) + " на 1000 слов)."
    );

    if (m.deadCount === 0 && m.phantomCount === 0) {
      lines.push("Двусторонняя сверка не выявила расхождений: каждая позиция списка отработана в тексте, каждая внутритекстовая ссылка разрешается в списке.");
    } else {
      var parts = [];
      if (m.deadCount) {
        parts.push(m.deadCount + " " + plural(m.deadCount, "позиция списка не цитируется", "позиции списка не цитируются", "позиций списка не цитируются") +
          " (" + Math.round(m.deadRatio * 100) + "% объёма списка)");
      }
      if (m.phantomCount) {
        parts.push(m.phantomCount + " " + plural(m.phantomCount, "ссылка не имеет", "ссылки не имеют", "ссылок не имеют") +
          " соответствия в списке");
      }
      lines.push("Расхождения двусторонней сверки: " + parts.join("; ") + ".");
    }

    if (m.bulkCount) {
      lines.push(
        "Обнаружено " + m.bulkCount + " " +
        plural(m.bulkCount, "пакетный интервал", "пакетных интервала", "пакетных интервалов") +
        " (более " + report.options.bulkThreshold + " источников в одной скобке). " +
        "Такие конструкции не позволяют соотнести конкретное утверждение с конкретным источником и функционально эквивалентны отсутствию ссылки."
      );
    }

    if (report.bibliometrics && report.bibliometrics.total) {
      var bib = report.bibliometrics;
      var depth = "";

      if (bib.datedCount) {
        depth = "Хронологическая глубина: " + bib.oldest + "–" + bib.newest +
          ", медиана " + bib.medianYear + "; за последние пять лет — " + bib.recent5 +
          " из " + bib.total + " (" + Math.round(bib.recent5Share * 100) + "%), " +
          "за десять — " + bib.recent10 + " (" + Math.round(bib.recent10Share * 100) + "%).";
      } else {
        depth = "Год издания не удалось определить ни у одной записи.";
      }
      lines.push(depth);

      lines.push(
        "Состав: " + bib.englishCount + " " +
        plural(bib.englishCount, "англоязычный источник", "англоязычных источника", "англоязычных источников") +
        " (" + Math.round(bib.englishShare * 100) + "%), " +
        bib.doiCount + " " + plural(bib.doiCount, "запись", "записи", "записей") +
        " с DOI (" + Math.round(bib.doiShare * 100) + "%), из них англоязычных с DOI — " +
        bib.englishWithDoi + ". " + indexingCaveat()
      );
    }

    lines.push(
      "Интегральная оценка: " + report.verdict.score + "/100 — " +
      report.verdict.band.label.toLowerCase() + ". " + report.verdict.band.caption + "."
    );

    return lines;
  }

  /** Markdown export, five mandated sections. */
  function reportToMarkdown(report) {
    var m = report.metrics;
    var out = [];

    out.push("# Аудит цитирования — BiblioAudit");
    out.push("");
    out.push("Сформировано: " + new Date(report.generatedAt).toLocaleString("ru-RU"));
    out.push("Порог пакетного цитирования: > " + report.options.bulkThreshold + " источников в скобке.");
    out.push("Метод определения списка: " + (
      report.detection.method === "heading" ? "по заголовку раздела" :
      report.detection.method === "trailing-block" ? "по завершающему нумерованному блоку" :
      "раздел не обнаружен"
    ) + ".");
    out.push("");

    // --- 1 ----------------------------------------------------------------
    out.push("## 1. Общее количество источников и уникальных упоминаний");
    out.push("");
    out.push("| Показатель | Значение |");
    out.push("| --- | --- |");
    out.push("| Записей в списке литературы | " + m.totalReferences + " |");
    out.push("| Уникальных источников, упомянутых в тексте | " + m.uniqueCited + " |");
    out.push("| Из них разрешаются в списке | " + m.resolvedCited + " |");
    out.push("| Всего внутритекстовых ссылок (вхождений) | " + m.occurrenceCount + " |");
    out.push("| Охват списка | " + Math.round(m.coverage * 100) + "% |");
    out.push("| Слов в основном тексте | " + m.wordCount + " |");
    out.push("| Плотность цитирования | " + m.citationsPer1000Words.toFixed(1) + " на 1000 слов |");
    out.push("| Среднее число обращений к источнику | " + m.averageUsesPerSource.toFixed(2) + " |");
    out.push("");
    out.push("Цитируемые номера: " + report.citedCompressed);
    out.push("");

    if (report.languages && report.languages.length > 1) {
      out.push("Разрез по языковым версиям (списков литературы: " +
               m.referenceListCount + "):");
      out.push("");
      var langTable = buildLanguageTable(report);
      out.push("| " + langTable.header.join(" | ") + " |");
      out.push("| " + langTable.header.map(function () { return "---"; }).join(" | ") + " |");
      langTable.rows.forEach(function (row) { out.push("| " + row.join(" | ") + " |"); });
      out.push("");
    }

    [
      { title: "Распределение источников по годам издания:", table: buildYearTable(report) },
      { title: "Библиометрический профиль списка:", table: buildBibliometricTable(report) }
    ].forEach(function (part) {
      out.push(part.title);
      out.push("");
      out.push("| " + part.table.header.join(" | ") + " |");
      out.push("| " + part.table.header.map(function () { return "---"; }).join(" | ") + " |");
      part.table.rows.forEach(function (row) { out.push("| " + row.join(" | ") + " |"); });
      out.push("");
    });

    out.push(indexingCaveat());
    out.push("");

    // --- 2 ----------------------------------------------------------------
    out.push("## 2. Мёртвые источники (в списке есть, в тексте нет)");
    out.push("");
    if (!report.deadSources.length) {
      out.push("Не выявлено. Каждая позиция списка имеет минимум одно упоминание в тексте.");
    } else {
      out.push("Всего: **" + report.deadSources.length + "** из " + m.totalReferences +
        " (" + Math.round(m.deadRatio * 100) + "%).");
      out.push("");
      report.deadSources.forEach(function (item) {
        out.push("- **[" + item.number + "]** " + item.description);
      });
    }
    out.push("");

    // --- 3 ----------------------------------------------------------------
    out.push("## 3. Фантомные ссылки (в тексте есть, в списке нет)");
    out.push("");
    if (!report.phantomReferences.length) {
      out.push("Не выявлено. Все внутритекстовые номера разрешаются в списке литературы.");
    } else {
      out.push("Всего: **" + report.phantomReferences.length + "**.");
      out.push("");
      report.phantomReferences.forEach(function (item) {
        out.push("- **[" + item.number + "]** — " + item.count + " " +
          plural(item.count, "вхождение", "вхождения", "вхождений") +
          "; первое: абзац " + (item.occurrences[0].paragraphIndex + 1) +
          ", `" + item.occurrences[0].raw + "`");
        out.push("  > " + item.occurrences[0].snippet);
      });
    }
    out.push("");

    // --- 4 ----------------------------------------------------------------
    out.push("## 4. Подозрительные пакетные интервалы (> " + report.options.bulkThreshold + " источников)");
    out.push("");
    if (!report.bulkCitations.length) {
      out.push("Не выявлено. Ни одна скобка не превышает установленный порог.");
    } else {
      out.push("Всего: **" + report.bulkCitations.length + "**.");
      out.push("");
      report.bulkCitations.forEach(function (item) {
        out.push("- `" + item.raw + "` — " + item.size + " " +
          plural(item.size, "источник", "источника", "источников") +
          " (абзац " + (item.paragraphIndex + 1) + ")" +
          (item.widestRange ? "; максимальный диапазон " + item.widestRange.from + "–" + item.widestRange.to : ""));
        out.push("  > " + item.snippet);
      });
    }
    out.push("");

    // --- 5 ----------------------------------------------------------------
    out.push("## 5. Экспертное заключение");
    out.push("");
    out.push("**Интегральная оценка: " + report.verdict.score + "/100 — " + report.verdict.band.label + ".**");
    out.push("");
    buildConclusionSentences(report).forEach(function (line) {
      out.push(line);
      out.push("");
    });

    if (report.crossLanguage && report.crossLanguage.lists.length > 1) {
      out.push("Межъязыковая сверка:");
      if (report.crossLanguage.issues.length) {
        report.crossLanguage.issues.forEach(function (issue) { out.push("- " + issue.text); });
      } else {
        out.push("- Перечни источников во всех языковых версиях совпадают по объёму и нумерации.");
      }
      out.push("");
    }

    if (report.verdict.penalties.length) {
      out.push("Структура штрафов:");
      report.verdict.penalties.forEach(function (p) { out.push("- " + p.text); });
      out.push("");
    }

    var d = report.diagnostics;
    var diagnosticLines = [];
    if (d.duplicateNumbers.length) diagnosticLines.push("Дублирующиеся номера в списке: " + d.duplicateNumbers.join(", ") + ".");
    if (d.numberingGaps.length) diagnosticLines.push("Пропуски в нумерации списка: " + compressRanges(d.numberingGaps) + ".");
    if (d.malformedTokens.length) diagnosticLines.push("Некорректные токены в скобках: " + d.malformedTokens.length + ".");
    if (d.ignoredBrackets.length) diagnosticLines.push("Нечисловых скобок пропущено: " + d.ignoredBrackets.length + ".");
    if (d.unnumberedLines.length) diagnosticLines.push("Ненумерованных строк в разделе литературы: " + d.unnumberedLines.length + ".");
    if (d.orderViolations.length) diagnosticLines.push("Нарушений порядка первого упоминания: " + d.orderViolations.length + ".");

    if (diagnosticLines.length) {
      out.push("Диагностика парсинга:");
      diagnosticLines.forEach(function (line) { out.push("- " + line); });
      out.push("");
    }

    return out.join("\n");
  }

  /**
   * Plain-text lines used by the "insert into document" action.
   * Markdown tables are flattened to "Label: value" and separator rows dropped.
   */
  function reportToDocumentLines(report) {
    var lines = reportToMarkdown(report).split("\n");
    var out = [];

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];

      if (/^\|.*\|$/.test(line)) {
        var cells = line
          .replace(/^\|\s*/, "")
          .replace(/\s*\|$/, "")
          .split(/\s*\|\s*/);

        // Separator row: every cell is dashes/colons only.
        var isSeparator = cells.every(function (cell) { return /^:?-{2,}:?$/.test(cell.trim()); });
        if (isSeparator) continue;

        line = cells.join(": ");
      }

      out.push(line.replace(/\*\*/g, "").replace(/`/g, ""));
    }

    return out;
  }

  function reportToJson(report) {
    return JSON.stringify(report, null, 2);
  }

  /* ---------------------------------------------------------------------------
   * 5b. Word export model
   * ---------------------------------------------------------------------------
   * The exported document is described as a flat list of typed blocks so that the
   * Office.js writer stays trivial and the layout is testable without Word.
   *   { type: "title" | "meta" | "h" | "p" | "li" | "table" }
   * -------------------------------------------------------------------------*/

  /** "АудитЛит Статья АМА СШ СО" — the name Word will suggest on save. */
  function buildExportFileName(report) {
    var base = (report.source && report.source.fileName) ? String(report.source.fileName).trim() : "";
    if (!base) base = "Документ";
    // Strip characters Windows forbids in file names.
    base = base.replace(/[\\/:*?"<>|]/g, " ").replace(/\s+/g, " ").trim();
    return CONFIG.exportFilePrefix + " " + base;
  }

  /** Section 1 as a two-column table: [header, ...rows]. */
  function buildSummaryTable(report) {
    var m = report.metrics;
    return {
      type: "table",
      header: ["Показатель", "Значение"],
      rows: [
        ["Записей в списке литературы", String(m.totalReferences)],
        ["Уникальных источников, упомянутых в тексте", String(m.uniqueCited)],
        ["Из них разрешаются в списке", String(m.resolvedCited)],
        ["Всего внутритекстовых ссылок (вхождений)", String(m.occurrenceCount)],
        ["Охват списка", Math.round(m.coverage * 100) + "%"],
        ["Мёртвых источников", String(m.deadCount)],
        ["Фантомных ссылок", String(m.phantomCount)],
        ["Пакетных интервалов (> " + report.options.bulkThreshold + ")", String(m.bulkCount)],
        ["Слов в основном тексте", String(m.wordCount)],
        ["Плотность цитирования (на 1000 слов)", m.citationsPer1000Words.toFixed(1)],
        ["Среднее число обращений к источнику", m.averageUsesPerSource.toFixed(2)],
        ["Цитируемые номера", report.citedCompressed]
      ]
    };
  }

  /** Разрез по языковым версиям — вторая таблица раздела 1. */
  function buildLanguageTable(report) {
    var rows = [];

    for (var i = 0; i < report.languages.length; i++) {
      var item = report.languages[i];
      if (!item.words && !item.referenceEntries) continue;
      rows.push([
        item.label,
        String(item.words),
        String(item.citations),
        String(item.referenceEntries),
        String(item.deadCount),
        String(item.phantomCount)
      ]);
    }

    return {
      type: "table",
      header: ["Язык", "Слов", "Ссылок", "Записей", "Мёртвых", "Фантомов"],
      rows: rows
    };
  }

  /** Распределение источников по годам: последние 5 лет поштучно, дальше пятилетиями. */
  function buildYearTable(report) {
    var bib = report.bibliometrics;
    var rows = [];
    var total = bib.total || 1;

    for (var i = 0; i < bib.buckets.length; i++) {
      var bucket = bib.buckets[i];
      // Последние пять лет показываем всегда, пятилетия — только непустые.
      if (i >= 5 && bucket.count === 0) continue;
      rows.push([
        bucket.label,
        String(bucket.count),
        Math.round((bucket.count / total) * 100) + "%"
      ]);
    }

    if (bib.undatedCount) {
      rows.push([
        "год не определён",
        String(bib.undatedCount),
        Math.round((bib.undatedCount / total) * 100) + "%"
      ]);
    }

    return { type: "table", header: ["Период", "Источников", "Доля"], rows: rows };
  }

  /** Библиометрический профиль: глубина, DOI, языковой состав, пометки индексации. */
  function buildBibliometricTable(report) {
    var bib = report.bibliometrics;
    var pct = function (share) { return Math.round(share * 100) + "%"; };
    var rows = [];

    rows.push(["Всего записей в списках", String(bib.total)]);
    rows.push(["Год издания определён", String(bib.datedCount)]);
    if (bib.undatedCount) rows.push(["Год издания не определён", String(bib.undatedCount)]);

    if (bib.datedCount) {
      rows.push(["Диапазон лет", bib.oldest + "–" + bib.newest]);
      rows.push(["Медианный год", String(bib.medianYear)]);
      rows.push([
        "За последние 5 лет (с " + (bib.currentYear - 4) + ")",
        bib.recent5 + " (" + pct(bib.recent5Share) + ")"
      ]);
      rows.push([
        "За последние 10 лет (с " + (bib.currentYear - 9) + ")",
        bib.recent10 + " (" + pct(bib.recent10Share) + ")"
      ]);
    }

    rows.push(["Источников с DOI", bib.doiCount + " (" + pct(bib.doiShare) + ")"]);
    rows.push(["Англоязычных источников", bib.englishCount + " (" + pct(bib.englishShare) + ")"]);
    rows.push(["Из них с DOI", String(bib.englishWithDoi)]);
    rows.push(["Русскоязычных источников", String(bib.russianCount)]);
    if (bib.tajikCount) rows.push(["Таджикоязычных источников", String(bib.tajikCount)]);
    rows.push(["Явная пометка Scopus в записи", String(bib.scopusMentions)]);
    rows.push(["Явная пометка Web of Science в записи", String(bib.wosMentions)]);

    return { type: "table", header: ["Показатель", "Значение"], rows: rows };
  }

  /** Оговорка о том, что индексацию по тексту записи установить нельзя. */
  function indexingCaveat() {
    return "Принадлежность к Scopus и Web of Science по тексту записи не определяется: " +
           "это свойство внешних баз, проверяемое только запросом к ним. Подсчитаны " +
           "англоязычные источники, наличие DOI и явные пометки, проставленные автором.";
  }

  /** Full block list of the exported report. */
  function buildExportBlocks(report, fileName) {
    var m = report.metrics;
    var blocks = [];

    blocks.push({ type: "title", text: fileName });
    blocks.push({
      type: "meta",
      text: "Отчёт аудита цитирования. Сформировано: " +
        new Date(report.generatedAt).toLocaleString("ru-RU") +
        ". Порог пакетного цитирования: более " + report.options.bulkThreshold +
        " источников в одной скобке. Метод определения списка: " + (
          report.detection.method === "heading" ? "по заголовку раздела" :
          report.detection.method === "trailing-block" ? "по завершающему нумерованному блоку" :
          "раздел не обнаружен"
        ) + "."
    });

    // --- 1 ----------------------------------------------------------------
    blocks.push({ type: "h", text: "1. Общее количество источников и уникальных упоминаний" });
    blocks.push(buildSummaryTable(report));

    if (report.languages && report.languages.length > 1) {
      blocks.push({
        type: "p",
        text: "Разрез по языковым версиям (списков литературы в документе: " +
              report.metrics.referenceListCount + "):"
      });
      blocks.push(buildLanguageTable(report));
    }

    blocks.push({ type: "p", text: "Распределение источников по годам издания:" });
    blocks.push(buildYearTable(report));

    blocks.push({ type: "p", text: "Библиометрический профиль списка:" });
    blocks.push(buildBibliometricTable(report));
    blocks.push({ type: "p", text: indexingCaveat() });

    // --- 2 ----------------------------------------------------------------
    blocks.push({ type: "h", text: "2. Мёртвые источники (в списке есть, в тексте нет)" });
    if (!report.deadSources.length) {
      blocks.push({ type: "p", text: "Не выявлено. Каждая позиция списка имеет минимум одно упоминание в тексте." });
    } else {
      blocks.push({
        type: "p",
        text: "Всего: " + report.deadSources.length + " из " + m.totalReferences +
              " (" + Math.round(m.deadRatio * 100) + "% объёма списка)."
      });
      report.deadSources.forEach(function (item) {
        var tag = report.metrics.referenceListCount > 1 ? " (" + item.languageLabel + " список)" : "";
        blocks.push({ type: "li", text: "[" + item.number + "]" + tag + " " + item.description });
      });
    }

    // --- 3 ----------------------------------------------------------------
    blocks.push({ type: "h", text: "3. Фантомные ссылки (в тексте есть, в списке нет)" });
    if (!report.phantomReferences.length) {
      blocks.push({ type: "p", text: "Не выявлено. Все внутритекстовые номера разрешаются в списке литературы." });
    } else {
      blocks.push({ type: "p", text: "Всего: " + report.phantomReferences.length + "." });
      report.phantomReferences.forEach(function (item) {
        var tag = report.languages.length > 1 ? " (" + item.languageLabel + " раздел)" : "";
        blocks.push({
          type: "li",
          text: "[" + item.number + "]" + tag + " — " + item.count + " " +
                plural(item.count, "вхождение", "вхождения", "вхождений") +
                "; первое: абзац " + (item.occurrences[0].paragraphIndex + 1) +
                ". Контекст: " + item.occurrences[0].snippet
        });
      });
    }

    // --- 4 ----------------------------------------------------------------
    blocks.push({
      type: "h",
      text: "4. Подозрительные пакетные интервалы (более " + report.options.bulkThreshold + " источников)"
    });
    if (!report.bulkCitations.length) {
      blocks.push({ type: "p", text: "Не выявлено. Ни одна скобка не превышает установленный порог." });
    } else {
      blocks.push({ type: "p", text: "Всего: " + report.bulkCitations.length + "." });
      report.bulkCitations.forEach(function (item) {
        blocks.push({
          type: "li",
          text: item.raw + " — " + item.size + " " +
                plural(item.size, "источник", "источника", "источников") +
                " в одной скобке (абзац " + (item.paragraphIndex + 1) + ")" +
                (item.widestRange ? "; максимальный диапазон " + item.widestRange.from + "–" + item.widestRange.to : "") +
                ". Контекст: " + item.snippet
        });
      });
    }

    // --- 5 ----------------------------------------------------------------
    blocks.push({ type: "h", text: "5. Экспертное заключение" });
    blocks.push({
      type: "p",
      text: "Интегральная оценка: " + report.verdict.score + "/100 — " + report.verdict.band.label + "."
    });
    buildConclusionSentences(report).forEach(function (line) {
      blocks.push({ type: "p", text: line });
    });

    if (report.crossLanguage && report.crossLanguage.lists.length > 1) {
      if (report.crossLanguage.issues.length) {
        blocks.push({ type: "p", text: "Межъязыковая сверка — выявлены расхождения:" });
        report.crossLanguage.issues.forEach(function (issue) {
          blocks.push({ type: "li", text: issue.text });
        });
      } else {
        blocks.push({
          type: "p",
          text: "Межъязыковая сверка: перечни источников во всех языковых версиях " +
                "совпадают по объёму и нумерации."
        });
      }
    }

    if (report.verdict.penalties.length) {
      blocks.push({ type: "p", text: "Структура штрафов:" });
      report.verdict.penalties.forEach(function (p) {
        blocks.push({ type: "li", text: p.text });
      });
    }

    var d = report.diagnostics;
    var diagnostics = [];
    if (d.duplicateNumbers.length) diagnostics.push("дублирующиеся номера в списке: " + d.duplicateNumbers.join(", "));
    if (d.numberingGaps.length) diagnostics.push("пропуски в нумерации списка: " + compressRanges(d.numberingGaps));
    if (d.malformedTokens.length) diagnostics.push("некорректных токенов в скобках: " + d.malformedTokens.length);
    if (d.ignoredBrackets.length) diagnostics.push("нечисловых скобок пропущено: " + d.ignoredBrackets.length);
    if (d.unnumberedLines.length) diagnostics.push("ненумерованных строк в разделе литературы: " + d.unnumberedLines.length);
    if (d.orderViolations.length) diagnostics.push("нарушений порядка первого упоминания: " + d.orderViolations.length);

    if (diagnostics.length) {
      blocks.push({ type: "p", text: "Диагностика парсинга: " + diagnostics.join("; ") + "." });
    }

    return blocks;
  }

  /* ===========================================================================
   * 6. DOM HELPERS + RENDERING
   * =========================================================================*/

  var dom = {};

  function cacheDom() {
    dom.runButton = document.getElementById("run-audit");
    dom.bulkThreshold = document.getElementById("bulk-threshold");
    dom.optScanWhole = document.getElementById("opt-scan-whole-doc");
    dom.optStrictOrder = document.getElementById("opt-strict-order");
    dom.status = document.getElementById("status");
    dom.skeleton = document.getElementById("skeleton");
    dom.errorBanner = document.getElementById("error-banner");
    dom.errorText = document.getElementById("error-text");
    dom.results = document.getElementById("results");
    dom.exportPanel = document.getElementById("export-panel");
    dom.exportStatus = document.getElementById("export-status");
    dom.emptyState = document.getElementById("empty-state");
    dom.copyMarkdown = document.getElementById("copy-markdown");
    dom.copyJson = document.getElementById("copy-json");
    dom.insertReport = document.getElementById("insert-report");
    dom.exportDocx = document.getElementById("export-docx");
    dom.exportFileName = document.getElementById("export-filename");
    dom.compatBanner = document.getElementById("compat-banner");
    dom.compatTitle = document.getElementById("compat-title");
    dom.compatText = document.getElementById("compat-text");
    dom.hostInfo = document.getElementById("host-info");
  }

  /** Безопасная запись текста: отсутствие узла не должно ронять весь аудит. */
  function setText(node, value) {
    if (node) node.textContent = value || "";
  }

  function setStatus(node, message, state) {
    if (!node) return;
    node.textContent = message || "";
    node.classList.remove("is-error", "is-success");
    if (state) node.classList.add("is-" + state);
  }

  function setLoading(isLoading) {
    dom.runButton.disabled = isLoading;
    dom.runButton.classList.toggle("is-loading", isLoading);
    dom.runButton.querySelector(".ba-btn__label").textContent =
      isLoading ? "Анализ документа…" : "Провести аудит";
    dom.skeleton.hidden = !isLoading;
    if (isLoading) {
      dom.results.hidden = true;
      dom.exportPanel.hidden = true;
      dom.emptyState.hidden = true;
      dom.errorBanner.hidden = true;
    }
  }

  function showError(message) {
    dom.errorText.textContent = message;
    dom.errorBanner.hidden = false;
    dom.skeleton.hidden = true;
  }

  /** Build one collapsible section. */
  function sectionMarkup(index, title, count, tone, bodyHtml, open) {
    var countClass = tone ? " ba-section__count--" + tone : "";
    return "" +
      '<article class="ba-section' + (open ? " is-open" : "") + '">' +
        '<button class="ba-section__head" type="button" aria-expanded="' + (open ? "true" : "false") + '">' +
          '<span class="ba-section__index">' + index + "</span>" +
          '<span class="ba-section__title">' + escapeHtml(title) + "</span>" +
          '<span class="ba-section__count' + countClass + '">' + escapeHtml(count) + "</span>" +
          '<span class="ba-section__chevron" aria-hidden="true"></span>' +
        "</button>" +
        '<div class="ba-section__body">' + bodyHtml + "</div>" +
      "</article>";
  }

  function cleanMarkup(message) {
    return '<p class="ba-clean">' + escapeHtml(message) + "</p>";
  }

  function statTile(value, label, tone) {
    return '<div class="ba-stat' + (tone ? " ba-stat--" + tone : "") + '">' +
      '<div class="ba-stat__value">' + escapeHtml(value) + "</div>" +
      '<div class="ba-stat__label">' + escapeHtml(label) + "</div>" +
      "</div>";
  }

  /** Табличный блок панели из описания {header, rows}. */
  function tableCard(title, meta, table) {
    var head = "<tr>" + table.header.map(function (cell) {
      return "<th>" + escapeHtml(cell) + "</th>";
    }).join("") + "</tr>";

    var body = table.rows.map(function (row) {
      return "<tr>" + row.map(function (cell, columnIndex) {
        return "<td" + (columnIndex ? ' class="ba-num"' : "") + ">" + escapeHtml(cell) + "</td>";
      }).join("") + "</tr>";
    }).join("");

    return '<div class="ba-item ba-item--neutral"><div class="ba-item__body">' +
      '<div class="ba-item__text"><strong>' + escapeHtml(title) + "</strong></div>" +
      (meta ? '<div class="ba-item__meta">' + escapeHtml(meta) + "</div>" : "") +
      '<div class="ba-tablewrap"><table class="ba-table">' +
        "<thead>" + head + "</thead><tbody>" + body + "</tbody>" +
      "</table></div>" +
    "</div></div>";
  }

  function renderReport(report) {
    var m = report.metrics;
    var html = [];

    // --- verdict banner ---------------------------------------------------
    html.push(
      '<div class="ba-verdict ba-verdict--' + report.verdict.band.key + '">' +
        '<div class="ba-verdict__score">' + report.verdict.score + "</div>" +
        "<div>" +
          '<div class="ba-verdict__label">' + escapeHtml(report.verdict.band.label) + "</div>" +
          '<div class="ba-verdict__caption">' + escapeHtml(report.verdict.band.caption) + "</div>" +
        "</div>" +
      "</div>"
    );

    // --- stat tiles -------------------------------------------------------
    html.push(
      '<div class="ba-stats">' +
        statTile(m.totalReferences, "записей в списке", null) +
        statTile(m.uniqueCited, "уникальных в тексте", null) +
        statTile(m.deadCount, "мёртвых источников", m.deadCount ? "danger" : "success") +
        statTile(m.phantomCount, "фантомных ссылок", m.phantomCount ? "danger" : "success") +
        statTile(m.bulkCount, "пакетных интервалов", m.bulkCount ? "warning" : "success") +
        statTile(Math.round(m.coverage * 100) + "%", "охват списка", m.coverage >= 0.9 ? "success" : "warning") +
      "</div>"
    );

    // === Section 1 ========================================================
    var detectionLabel =
      report.detection.method === "heading" ? "по заголовку раздела" :
      report.detection.method === "trailing-block" ? "по завершающему нумерованному блоку" :
      "раздел не обнаружен";

    var s1 = [];
    s1.push('<div class="ba-list">');
    s1.push(
      '<div class="ba-item ba-item--neutral"><div class="ba-item__body">' +
        '<div class="ba-item__text"><strong>' + m.totalReferences + "</strong> " +
          plural(m.totalReferences, "запись", "записи", "записей") + " в списке литературы; " +
          "<strong>" + m.uniqueCited + "</strong> " +
          plural(m.uniqueCited, "уникальный источник", "уникальных источника", "уникальных источников") +
          " упомянуто в тексте, из них <strong>" + m.resolvedCited + "</strong> разрешаются в списке." +
        "</div>" +
        '<div class="ba-item__meta">' +
          "вхождений ссылок: " + m.occurrenceCount + " · " +
          "слов: " + m.wordCount + " · " +
          "плотность: " + m.citationsPer1000Words.toFixed(1) + "/1000 слов · " +
          "среднее обращений к источнику: " + m.averageUsesPerSource.toFixed(2) +
        "</div>" +
        '<div class="ba-item__meta">раздел литературы определён ' + escapeHtml(detectionLabel) +
          " · абзацев текста: " + report.detection.bodyParagraphCount +
          " · абзацев списка: " + report.detection.referenceParagraphCount +
          (report.detection.tailParagraphCount
            ? " · абзацев после списка (возвращены в анализ): " + report.detection.tailParagraphCount
            : "") + "</div>" +
      "</div></div>"
    );

    s1.push(
      '<div class="ba-item ba-item--neutral"><div class="ba-item__body">' +
        '<div class="ba-item__text"><strong>Цитируемые номера</strong></div>' +
        '<div class="ba-item__meta">' + escapeHtml(report.citedCompressed) + "</div>" +
      "</div></div>"
    );

    if (report.mostCited.length) {
      s1.push(
        '<div class="ba-item ba-item--neutral"><div class="ba-item__body">' +
          '<div class="ba-item__text"><strong>Наиболее востребованные источники</strong></div>' +
          '<div class="ba-chips">' +
            report.mostCited.map(function (item) {
              return '<span class="ba-chip">[' + item.number + "] ×" + item.count + "</span>";
            }).join("") +
          "</div>" +
        "</div></div>"
      );
    }

    if (report.languages && report.languages.length > 1) {
      var langTableUi = buildLanguageTable(report);
      var head = "<tr>" + langTableUi.header.map(function (cell) {
        return "<th>" + escapeHtml(cell) + "</th>";
      }).join("") + "</tr>";

      var body = langTableUi.rows.map(function (row) {
        return "<tr>" + row.map(function (cell, columnIndex) {
          return "<td" + (columnIndex ? ' class="ba-num"' : "") + ">" + escapeHtml(cell) + "</td>";
        }).join("") + "</tr>";
      }).join("");

      s1.push(
        '<div class="ba-item ba-item--neutral"><div class="ba-item__body">' +
          '<div class="ba-item__text"><strong>Разрез по языковым версиям</strong></div>' +
          '<div class="ba-item__meta">списков литературы в документе: ' +
            m.referenceListCount + " · секций: " + report.detection.sectionCount + "</div>" +
          '<div class="ba-tablewrap"><table class="ba-table">' +
            "<thead>" + head + "</thead><tbody>" + body + "</tbody>" +
          "</table></div>" +
        "</div></div>"
      );
    }

    var bib = report.bibliometrics;
    s1.push(tableCard(
      "Источники по годам издания",
      "год определён у " + bib.datedCount + " из " + bib.total +
        (bib.medianYear ? " · медиана " + bib.medianYear : ""),
      buildYearTable(report)
    ));
    s1.push(tableCard(
      "Библиометрический профиль",
      "индексация в Scopus/WoS по тексту записи не определяется",
      buildBibliometricTable(report)
    ));

    // parsing diagnostics folded into section 1
    var d = report.diagnostics;
    var diagnosticItems = [];
    if (d.duplicateNumbers.length) diagnosticItems.push("дублирующиеся номера в списке: " + d.duplicateNumbers.join(", "));
    if (d.numberingGaps.length) diagnosticItems.push("пропуски в нумерации: " + compressRanges(d.numberingGaps));
    if (d.malformedTokens.length) diagnosticItems.push("некорректных токенов в скобках: " + d.malformedTokens.length);
    if (d.ignoredBrackets.length) diagnosticItems.push("нечисловых скобок пропущено: " + d.ignoredBrackets.length);
    if (d.unnumberedLines.length) diagnosticItems.push("ненумерованных строк в разделе литературы: " + d.unnumberedLines.length);
    if (d.orderViolations.length) diagnosticItems.push("нарушений порядка первого упоминания: " + d.orderViolations.length);

    if (diagnosticItems.length) {
      s1.push(
        '<div class="ba-item ba-item--warning">' +
          '<span class="ba-item__num">!</span>' +
          '<div class="ba-item__body">' +
            '<div class="ba-item__text"><strong>Диагностика парсинга</strong></div>' +
            '<div class="ba-item__meta">' + escapeHtml(diagnosticItems.join(" · ")) + "</div>" +
          "</div>" +
        "</div>"
      );
    }
    s1.push("</div>");

    html.push(sectionMarkup(
      1,
      "Объём списка и уникальные упоминания",
      m.totalReferences + " / " + m.uniqueCited,
      null,
      s1.join(""),
      true
    ));

    // === Section 2: dead sources =========================================
    var s2;
    if (!report.deadSources.length) {
      s2 = cleanMarkup("Мёртвых источников не обнаружено — каждая позиция списка отработана в тексте.");
    } else {
      s2 = '<div class="ba-list">' + report.deadSources.map(function (item) {
        return '<div class="ba-item ba-item--danger">' +
          '<span class="ba-item__num">' + item.number + "</span>" +
          '<div class="ba-item__body">' +
            '<div class="ba-item__text">' + escapeHtml(item.description) + "</div>" +
            '<div class="ba-item__meta">абзац ' + (item.paragraphIndex + 1) +
              " · упоминаний в тексте: 0" +
              (m.referenceListCount > 1 ? " · " + escapeHtml(item.languageLabel) + " список" : "") +
              "</div>" +
          "</div>" +
        "</div>";
      }).join("") + "</div>";
    }
    html.push(sectionMarkup(
      2,
      "Мёртвые источники",
      String(report.deadSources.length),
      report.deadSources.length ? "danger" : "success",
      s2,
      report.deadSources.length > 0
    ));

    // === Section 3: phantom references ===================================
    var s3;
    if (!report.phantomReferences.length) {
      s3 = cleanMarkup("Фантомных ссылок не обнаружено — все номера разрешаются в списке литературы.");
    } else {
      s3 = '<div class="ba-list">' + report.phantomReferences.map(function (item) {
        return '<div class="ba-item ba-item--danger">' +
          '<span class="ba-item__num">' + item.number + "</span>" +
          '<div class="ba-item__body">' +
            '<div class="ba-item__text">Номер отсутствует в списке литературы. Вхождений: <strong>' +
              item.count + "</strong>.</div>" +
            '<div class="ba-item__meta">первое вхождение: абзац ' +
              (item.occurrences[0].paragraphIndex + 1) + " · " +
              escapeHtml(item.occurrences[0].raw) +
              (report.languages.length > 1 ? " · " + escapeHtml(item.languageLabel) + " раздел" : "") +
              "</div>" +
            '<div class="ba-item__quote">' + escapeHtml(item.occurrences[0].snippet) + "</div>" +
          "</div>" +
        "</div>";
      }).join("") + "</div>";
    }
    html.push(sectionMarkup(
      3,
      "Фантомные ссылки",
      String(report.phantomReferences.length),
      report.phantomReferences.length ? "danger" : "success",
      s3,
      report.phantomReferences.length > 0
    ));

    // === Section 4: bulk intervals =======================================
    var s4;
    if (!report.bulkCitations.length) {
      s4 = cleanMarkup("Пакетных интервалов не обнаружено — ни одна скобка не превышает порог в " +
        report.options.bulkThreshold + " " +
        plural(report.options.bulkThreshold, "источник", "источника", "источников") + ".");
    } else {
      s4 = '<div class="ba-list">' + report.bulkCitations.map(function (item) {
        var chips = item.numbers.slice(0, 24).map(function (n) {
          return '<span class="ba-chip ba-chip--danger">' + n + "</span>";
        }).join("");
        if (item.numbers.length > 24) {
          chips += '<span class="ba-chip">+' + (item.numbers.length - 24) + "</span>";
        }
        return '<div class="ba-item ba-item--warning">' +
          '<span class="ba-item__num">' + item.size + "</span>" +
          '<div class="ba-item__body">' +
            '<div class="ba-item__text"><span class="ba-code">' + escapeHtml(item.raw) + "</span> — " +
              item.size + " " + plural(item.size, "источник", "источника", "источников") +
              " в одной скобке" +
              (item.widestRange
                ? " (диапазон " + item.widestRange.from + "–" + item.widestRange.to + ")"
                : "") +
              ".</div>" +
            '<div class="ba-item__meta">абзац ' + (item.paragraphIndex + 1) + "</div>" +
            '<div class="ba-item__quote">' + escapeHtml(item.snippet) + "</div>" +
            '<div class="ba-chips">' + chips + "</div>" +
          "</div>" +
        "</div>";
      }).join("") + "</div>";
    }
    html.push(sectionMarkup(
      4,
      "Пакетные интервалы (> " + report.options.bulkThreshold + ")",
      String(report.bulkCitations.length),
      report.bulkCitations.length ? "warning" : "success",
      s4,
      report.bulkCitations.length > 0
    ));

    // === Section 5: expert conclusion ====================================
    var s5 = '<div class="ba-prose">' +
      buildConclusionSentences(report).map(function (line) {
        return "<p>" + escapeHtml(line) + "</p>";
      }).join("");

    if (report.crossLanguage && report.crossLanguage.lists.length > 1) {
      s5 += "<p><strong>Межъязыковая сверка</strong></p>";
      if (report.crossLanguage.issues.length) {
        s5 += "<ul>" + report.crossLanguage.issues.map(function (issue) {
          return '<li class="is-' + issue.level + '">' + escapeHtml(issue.text) + "</li>";
        }).join("") + "</ul>";
      } else {
        s5 += '<ul><li class="is-success">Перечни источников во всех языковых версиях ' +
              "совпадают по объёму и нумерации.</li></ul>";
      }
    }

    if (report.verdict.penalties.length) {
      s5 += "<ul>" + report.verdict.penalties.map(function (p) {
        return '<li class="is-' + p.level + '">' + escapeHtml(p.text) + "</li>";
      }).join("") + "</ul>";
    } else {
      s5 += "<ul><li class=\"is-success\">Штрафных факторов не зафиксировано.</li></ul>";
    }
    s5 += "</div>";

    html.push(sectionMarkup(
      5,
      "Экспертное заключение",
      report.verdict.score + "/100",
      report.verdict.band.key === "ok" ? "success" : report.verdict.band.key === "warn" ? "warning" : "danger",
      s5,
      true
    ));

    dom.results.innerHTML = html.join("");
    dom.results.hidden = false;
    dom.exportPanel.hidden = false;
    dom.emptyState.hidden = true;

    bindSectionToggles();
  }

  /** Attach collapse/expand behaviour to freshly rendered sections. */
  function bindSectionToggles() {
    var heads = dom.results.querySelectorAll(".ba-section__head");
    Array.prototype.forEach.call(heads, function (head) {
      head.addEventListener("click", function () {
        var section = head.parentNode;
        var isOpen = section.classList.toggle("is-open");
        head.setAttribute("aria-expanded", isOpen ? "true" : "false");
      });
    });
  }

  /* ===========================================================================
   * 7. OFFICE.JS INTEGRATION
   * =========================================================================*/

  /**
   * Read every paragraph of the document into plain models.
   * Auto-numbered list labels are fetched separately via listItemOrNullObject
   * because Word does not include them in paragraph.text.
   */
  function readDocumentParagraphs() {
    return Word.run(async function (context) {
      var paragraphCollection = context.document.body.paragraphs;
      paragraphCollection.load("items/text,items/style,items/isListItem");
      await context.sync();

      var items = paragraphCollection.items;

      // Second round-trip: resolve auto-numbering labels.
      var listRefs = items.map(function (paragraph) {
        var listItem = paragraph.listItemOrNullObject;
        listItem.load("listString,isNullObject");
        return listItem;
      });
      await context.sync();

      return items.map(function (paragraph, index) {
        // Word emits CR, vertical tab (soft line break) and NBSP inside paragraph text.
        var text = String(paragraph.text == null ? "" : paragraph.text)
          .replace(/[\r\u000B\u00A0]/g, " ");

        var listString = "";
        var ref = listRefs[index];
        if (ref && ref.isNullObject === false && ref.listString) {
          listString = String(ref.listString);
        }

        return {
          index: index,
          text: text,
          listString: listString,
          isListItem: !!paragraph.isListItem,
          style: paragraph.style || "",
          effective: listString ? listString + " " + text.trim() : text
        };
      });
    });
  }

  /** Append the rendered report to the end of the document. */
  function insertReportIntoDocument(report) {
    var STYLES = (typeof Word !== "undefined" && Word.Style)
      ? Word.Style
      : { heading1: "Heading1", heading2: "Heading2", normal: "Normal" };

    var lines = reportToDocumentLines(report);

    return Word.run(async function (context) {
      var body = context.document.body;

      body.insertParagraph("", Word.InsertLocation.end);

      var heading = body.insertParagraph("Отчёт аудита цитирования (BiblioAudit)", Word.InsertLocation.end);
      try { heading.styleBuiltIn = STYLES.heading1; } catch (e) { /* style unavailable in template */ }

      for (var i = 0; i < lines.length; i++) {
        var line = lines[i];

        if (/^#\s+/.test(line)) continue;                    // top-level title already inserted

        var isSubheading = /^##\s+/.test(line);
        var clean = line.replace(/^#{1,6}\s+/, "").replace(/\*\*/g, "").replace(/`/g, "");

        var paragraph = body.insertParagraph(clean, Word.InsertLocation.end);
        try {
          paragraph.styleBuiltIn = isSubheading ? STYLES.heading2 : STYLES.normal;
        } catch (e) { /* style unavailable in template */ }
      }

      await context.sync();
    });
  }

  /**
   * Classify the container format of the open document.
   * Word opens .doc / .rtf in Compatibility Mode, where part of the object model
   * behaves differently — worth telling the user before they read the numbers.
   */
  function getDocumentFormat() {
    var url = "";
    try {
      url = (Office.context && Office.context.document && Office.context.document.url) || "";
    } catch (error) {
      url = "";
    }

    if (!url) return { extension: "", legacy: false, saved: false };

    var match = /\.([A-Za-z0-9]+)$/.exec(url.split(/[\\/]/).pop() || "");
    var extension = match ? match[1].toLowerCase() : "";

    return {
      extension: extension,
      legacy: ["doc", "rtf", "txt", "wps", "dot"].indexOf(extension) !== -1,
      saved: true
    };
  }

  /** Base name of the document the task pane is attached to, without extension. */
  function getDocumentBaseName() {
    var url = "";
    try {
      url = (Office.context && Office.context.document && Office.context.document.url) || "";
    } catch (error) {
      url = "";
    }
    if (!url) return "";

    var name = url.split(/[\\/]/).pop() || "";
    try { name = decodeURIComponent(name); } catch (error) { /* keep raw */ }
    return name.replace(/\.[^.]+$/, "");
  }

  /** Times New Roman 12, single spacing, no extra space between paragraphs. */
  function applyBaseFormatting(paragraph) {
    try { paragraph.styleBuiltIn = Word.Style.normal; } catch (error) { /* template lacks style */ }
    paragraph.font.name = CONFIG.exportFontName;
    paragraph.font.size = CONFIG.exportFontSize;
    paragraph.font.color = "#000000";
    paragraph.font.bold = false;
    paragraph.font.italic = false;
    paragraph.lineSpacing = CONFIG.exportLineSpacing;   // 12pt == single
    paragraph.spaceBefore = 0;
    paragraph.spaceAfter = 0;
    paragraph.leftIndent = 0;
    paragraph.firstLineIndent = 0;
  }

  function applyBlockFormatting(paragraph, type) {
    applyBaseFormatting(paragraph);

    if (type === "title") {
      paragraph.font.bold = true;
      paragraph.alignment = Word.Alignment.centered;
      paragraph.spaceAfter = 6;
    } else if (type === "meta") {
      paragraph.font.italic = true;
      paragraph.alignment = Word.Alignment.left;
      paragraph.spaceAfter = 6;
    } else if (type === "h") {
      paragraph.font.bold = true;
      paragraph.alignment = Word.Alignment.left;
      paragraph.spaceBefore = 6;
    } else if (type === "li") {
      paragraph.alignment = Word.Alignment.justified;
      paragraph.leftIndent = 18;
    } else {
      paragraph.alignment = Word.Alignment.justified;
    }
  }

  /** Single-line black grid, header row in bold. */
  function applyTableFormatting(table) {
    table.headerRowCount = 1;
    table.alignment = Word.Alignment.left;
    table.font.name = CONFIG.exportFontName;
    table.font.size = CONFIG.exportFontSize;
    table.font.color = "#000000";

    var locations = [
      "Top", "Left", "Bottom", "Right", "InsideHorizontal", "InsideVertical"
    ];
    locations.forEach(function (location) {
      try {
        var border = table.getBorder(location);
        border.type = Word.BorderType.single;
        border.width = 1;
        border.color = "#000000";
      } catch (error) { /* border location unsupported */ }
    });
  }

  /** Write the block list into a body, in order. */
  function writeBlocks(body, blocks) {
    for (var i = 0; i < blocks.length; i++) {
      var block = blocks[i];

      if (block.type === "table") {
        var values = [block.header].concat(block.rows);
        var table = body.insertTable(
          values.length,
          block.header.length,
          Word.InsertLocation.end,
          values
        );
        applyTableFormatting(table);

        // Spacer paragraph so the next heading does not stick to the grid.
        applyBaseFormatting(body.insertParagraph("", Word.InsertLocation.end));
        continue;
      }

      var paragraph = body.insertParagraph(block.text, Word.InsertLocation.end);
      applyBlockFormatting(paragraph, block.type);
    }
  }

  /**
   * Build the report as a NEW Word document and open it in a separate window.
   *
   * The document is created unsaved: Office.js cannot write to an arbitrary path.
   * The file name is carried two ways so that Save As proposes it automatically —
   * as the document Title property and as the first paragraph of the report.
   */
  function exportReportToWordFile(report, fileName) {
    return Word.run(async function (context) {
      var application = context.application;

      if (!application || typeof application.createDocument !== "function") {
        throw new Error(
          "Эта версия Word не поддерживает создание документа из надстройки. " +
          "Используйте «Вставить в документ»."
        );
      }

      var newDocument = application.createDocument();
      await context.sync();

      var body = newDocument.body;
      body.clear();

      writeBlocks(body, buildExportBlocks(report, fileName));

      try { newDocument.properties.title = fileName; } catch (error) { /* properties unavailable */ }

      await context.sync();

      newDocument.open();
      await context.sync();
    });
  }

  /* ===========================================================================
   * 8. CLIPBOARD
   * =========================================================================*/

  async function copyToClipboard(text) {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (error) {
      // Fall through to the legacy path: some Word webviews block the async API.
    }

    var textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.top = "-1000px";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();

    var succeeded = false;
    try {
      succeeded = document.execCommand("copy");
    } catch (error) {
      succeeded = false;
    }
    document.body.removeChild(textarea);
    return succeeded;
  }

  /* ===========================================================================
   * 9. CONTROLLER
   * =========================================================================*/

  var state = { report: null };

  function readOptions() {
    var threshold = parseInt(dom.bulkThreshold.value, 10);
    if (!isFinite(threshold)) threshold = CONFIG.defaultBulkThreshold;
    threshold = clamp(threshold, 2, 50);
    dom.bulkThreshold.value = String(threshold);

    return {
      bulkThreshold: threshold,
      scanWholeDocument: dom.optScanWhole.checked,
      checkOrder: dom.optStrictOrder.checked
    };
  }

  async function handleRunAudit() {
    setLoading(true);
    setStatus(dom.status, "Чтение абзацев документа…");

    try {
      var options = readOptions();
      var paragraphs = await readDocumentParagraphs();

      if (!paragraphs.length) {
        setLoading(false);
        setStatus(dom.status, "Документ пуст — анализировать нечего.", "error");
        dom.emptyState.hidden = false;
        return;
      }

      setStatus(dom.status, "Разбор цитат и сверка со списком…");

      var report = auditDocument(paragraphs, options);
      report.source = { fileName: getDocumentBaseName() };
      state.report = report;

      if (report.detection.method === "none") {
        setStatus(
          dom.status,
          "Раздел литературы не найден. Проверьте заголовок («Список литературы», «References») " +
          "или нумерацию завершающего блока. Сверка выполнена без списка.",
          "error"
        );
      } else {
        setStatus(
          dom.status,
          "Готово: " + paragraphs.length + " " + plural(paragraphs.length, "абзац", "абзаца", "абзацев") +
          ", " + report.metrics.occurrenceCount + " " +
          plural(report.metrics.occurrenceCount, "ссылка", "ссылки", "ссылок") + " проанализировано.",
          "success"
        );
      }

      setLoading(false);
      renderReport(report);
      setStatus(dom.exportStatus, "");
      setText(dom.exportFileName, "Файл отчёта: " + buildExportFileName(report) + ".docx");
    } catch (error) {
      setLoading(false);
      setStatus(dom.status, "Аудит прерван.", "error");
      showError(describeError(error));
    }
  }

  /** Turn an Office.js OfficeExtension.Error (or any throwable) into a readable line. */
  function describeError(error) {
    if (!error) return "Неизвестная ошибка.";
    var parts = [];
    if (error.code) parts.push(error.code);
    if (error.message) parts.push(error.message);
    if (error.debugInfo && error.debugInfo.errorLocation) {
      parts.push("location: " + error.debugInfo.errorLocation);
    }
    return parts.length ? parts.join(" | ") : String(error);
  }

  async function handleCopy(serializer, label) {
    if (!state.report) return;
    var ok = await copyToClipboard(serializer(state.report));
    setStatus(
      dom.exportStatus,
      ok ? label + " скопирован в буфер обмена." : "Буфер обмена недоступен в этом окружении.",
      ok ? "success" : "error"
    );
  }

  async function handleInsertReport() {
    if (!state.report) return;
    dom.insertReport.disabled = true;
    setStatus(dom.exportStatus, "Вставка отчёта в конец документа…");
    try {
      await insertReportIntoDocument(state.report);
      setStatus(dom.exportStatus, "Отчёт добавлен в конец документа.", "success");
    } catch (error) {
      setStatus(dom.exportStatus, "Не удалось вставить отчёт.", "error");
      showError(describeError(error));
    } finally {
      dom.insertReport.disabled = false;
    }
  }

  async function handleExportDocx() {
    if (!state.report) return;

    var fileName = buildExportFileName(state.report);

    dom.exportDocx.disabled = true;
    dom.exportDocx.classList.add("is-loading");
    setStatus(dom.exportStatus, "Формирование документа «" + fileName + "»…");

    try {
      await exportReportToWordFile(state.report, fileName);
      setStatus(
        dom.exportStatus,
        "Отчёт открыт в новом окне Word. Сохраните его (Ctrl+S) — имя «" + fileName + "» уже подставлено.",
        "success"
      );
    } catch (error) {
      setStatus(dom.exportStatus, "Не удалось создать файл отчёта.", "error");
      showError(describeError(error));
    } finally {
      dom.exportDocx.disabled = false;
      dom.exportDocx.classList.remove("is-loading");
    }
  }

  function bindEvents() {
    dom.runButton.addEventListener("click", handleRunAudit);
    dom.exportDocx.addEventListener("click", handleExportDocx);
    dom.copyMarkdown.addEventListener("click", function () {
      handleCopy(reportToMarkdown, "Markdown-отчёт");
    });
    dom.copyJson.addEventListener("click", function () {
      handleCopy(reportToJson, "JSON-отчёт");
    });
    dom.insertReport.addEventListener("click", handleInsertReport);

    dom.bulkThreshold.addEventListener("change", function () {
      readOptions();
      if (state.report) setStatus(dom.status, "Порог изменён — запустите аудит повторно.");
    });
  }

  /* ===========================================================================
   * 10. BOOTSTRAP
   * =========================================================================*/

  Office.onReady(function (info) {
    cacheDom();

    if (info.host !== Office.HostType.Word) {
      dom.hostInfo.textContent = "Неподдерживаемое приложение: " + (info.host || "неизвестно");
      showError("BiblioAudit работает только в Microsoft Word.");
      dom.runButton.disabled = true;
      return;
    }

    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      dom.hostInfo.textContent = "WordApi 1.3 недоступен";
      showError("Требуется WordApi 1.3 или выше. Обновите Microsoft Word.");
      dom.runButton.disabled = true;
      return;
    }

    bindEvents();

    var format = getDocumentFormat();
    if (format.legacy && dom.compatBanner) {
      setText(dom.compatTitle, "Документ в формате ." + format.extension + " (режим совместимости)");
      setText(dom.compatText,
        "Аудит текста выполнится, но Word ограничивает часть операций для форматов до 2007 года. " +
        "Если экспорт отчёта или вставка в документ дадут сбой — преобразуйте файл командой " +
        "«Файл → Сведения → Преобразовать» и запустите аудит заново.");
      dom.compatBanner.hidden = false;
    }

    dom.hostInfo.textContent =
      "Word · " + (info.platform || "desktop") + " · WordApi 1.3 · BiblioAudit 1.0.0" +
      (format.extension ? " · ." + format.extension : "");
    setStatus(dom.status, "Готов к аудиту.");
  });

  /* Expose the pure core for unit tests / console experiments. */
  window.BiblioAudit = {
    auditDocument: auditDocument,
    analyseLayout: analyseLayout,
    detectLanguage: detectLanguage,
    detectLanguageShort: detectLanguageShort,
    foldKey: foldKey,
    splitDocument: splitDocument,
    parseReferenceList: parseReferenceList,
    parseBracketContent: parseBracketContent,
    collectCitations: collectCitations,
    reportToMarkdown: reportToMarkdown,
    reportToJson: reportToJson,
    buildExportBlocks: buildExportBlocks,
    analyseBibliometrics: analyseBibliometrics,
    buildYearTable: buildYearTable,
    buildBibliometricTable: buildBibliometricTable,
    extractYear: extractYear,
    entryLanguage: entryLanguage,
    buildExportFileName: buildExportFileName,
    compressRanges: compressRanges,
    CONFIG: CONFIG
  };
})();
