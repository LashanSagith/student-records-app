import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

const FULL_CREDIT = { level04: 120, level05: 120 };

/**
 * Extracts raw text from every page of a PDF File/Blob, in reading order.
 */
export async function extractPdfText(file) {
  const buf = await file.arrayBuffer();
  const doc = await pdfjsLib.getDocument({ data: buf }).promise;
  let fullText = "";
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    // Sort items top-to-bottom, left-to-right so rows read naturally.
    const items = content.items.slice().sort((a, b) => {
      const dy = b.transform[5] - a.transform[5];
      if (Math.abs(dy) > 2) return dy;
      return a.transform[4] - b.transform[4];
    });
    let lastY = null;
    items.forEach((item) => {
      const y = item.transform[5];
      if (lastY !== null && Math.abs(y - lastY) > 2) fullText += "\n";
      fullText += item.str + " ";
      lastY = y;
    });
    fullText += "\n\n";
  }
  return fullText;
}

/**
 * Parses "Board of Examiners (BoE) Report" text into one entry per student
 * block:  "(1104915) Mohamed Farshan Samry Ahamed" followed by a
 * Level / Credit attained / Level mark table containing FHEQ Level 4 / 5 rows.
 *
 * Returns: [{ ljmuId, name, level04Credit, level04Mark, level05Credit, level05Mark }]
 */
export function parseBoEReport(text) {
  const results = [];
  // Split on each "(digits) Name" occurrence — one per student block.
  const studentHeaderRe = /\((\d{5,10})\)\s+([A-Za-z][A-Za-z .'-]+?)(?=\n|\s{2,}|Programme status)/g;
  const matches = [...text.matchAll(studentHeaderRe)];

  matches.forEach((m, idx) => {
    const ljmuId = m[1].trim();
    const name = m[2].trim().replace(/\s+/g, " ");
    const blockStart = m.index;
    const blockEnd = idx + 1 < matches.length ? matches[idx + 1].index : text.length;
    const block = text.slice(blockStart, blockEnd);

    const level04 = extractLevelRow(block, "4");
    const level05 = extractLevelRow(block, "5");

    results.push({
      ljmuId,
      name,
      level04Credit: level04.credit,
      level04Mark: level04.mark,
      level05Credit: level05.credit,
      level05Mark: level05.mark,
    });
  });

  return results;
}

function extractLevelRow(block, levelDigit) {
  // Matches e.g. "FHEQ Level 4 120 48.62"
  const re = new RegExp(`FHEQ\\s+Level\\s+${levelDigit}\\s+([\\d.]+)\\s+([\\d.]+)`, "i");
  const m = block.match(re);
  if (!m) return { credit: null, mark: null };
  return { credit: parseFloat(m[1]), mark: parseFloat(m[2]) };
}

/**
 * Given a parsed report row, decides an auto Level 04/05 status.
 * Only ever auto-decides "Pass" (full credit attained). Anything short of
 * full credit is left null — those need a human to choose between
 * Progressing / Repeat / Fail & Exit, so we never guess on those.
 */
export function deriveLevelStatus(creditAttained, levelKey) {
  if (creditAttained === null || creditAttained === undefined) return null;
  const full = FULL_CREDIT[levelKey];
  return creditAttained >= full ? "Pass" : null;
}

/**
 * Builds the full diff plan for a parsed report against the current
 * students list, matching by LJMU ID.
 *
 * Returns { matched: [...], unmatched: [...] }
 * Each matched entry: { student, reportName, level04: {from,to,creditAttained,mark,needsReview}, level05: {...} }
 */
export function buildUpdatePlan(parsedRows, students) {
  const byLjmuId = new Map();
  students.forEach((s) => {
    const key = (s.ljmuId || "").trim().toLowerCase();
    if (key) byLjmuId.set(key, s);
  });

  const matched = [];
  const unmatched = [];

  parsedRows.forEach((row) => {
    const key = row.ljmuId.trim().toLowerCase();
    const student = byLjmuId.get(key);
    if (!student) {
      unmatched.push(row);
      return;
    }
    const l4To = deriveLevelStatus(row.level04Credit, "level04");
    const l5To = deriveLevelStatus(row.level05Credit, "level05");
    matched.push({
      student,
      reportName: row.name,
      level04: {
        creditAttained: row.level04Credit,
        mark: row.level04Mark,
        from: student.level04Status || "",
        to: l4To,
        needsReview: row.level04Credit !== null && l4To === null,
        willChange: l4To !== null && l4To !== (student.level04Status || ""),
      },
      level05: {
        creditAttained: row.level05Credit,
        mark: row.level05Mark,
        from: student.level05Status || "",
        to: l5To,
        needsReview: row.level05Credit !== null && l5To === null,
        willChange: l5To !== null && l5To !== (student.level05Status || ""),
      },
    });
  });

  return { matched, unmatched };
}
