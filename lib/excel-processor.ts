import * as XLSX from "xlsx";

export interface ProcessingStats {
  originalRows: number;
  cleanedRows: number;
  filteredRows: number;
  uomMatched: number;
  prefixesStripped: number;
}

interface FilteredRow {
  row: (string | number | null)[];
  reason: string;
}

const NOISE_PATTERNS = [
  /sage\s*200\s*evolution/i,
  /inventory\s*count\s*listing/i,
  /agri\s*technovation/i,
  /registered\s*to/i,
];

const PAGE_PATTERN = /page\s+\d+\s+of\s+\d+/i;

const HEADER_ROW_SIGNATURE = ["item code", "item description", "whse"];

const UOM_REGEX =
  /(\d+(?:\.\d+)?\s*[xX]\s*\d+(?:\.\d+)?\s*(?:L|ml|ML|kg|KG|g|G|oz)\s*(?:Bag|bag|Drum|drum|Can|can|Sachets?|sachets?|Pack|pack|Box|box|Bottle|bottle|IBC)?|\d+(?:\.\d+)?\s*(?:L|ml|ML|kg|KG|g|G|oz|litre|liter|ton|tonne)\s*(?:IBC|Bag|bag|Drum|drum|Can|can|Sachets?|sachets?|Pack|pack|Box|box|Bottle|bottle)?|IBC)/g;

const COUNTRY_CODE_REGEX = /^[A-Za-z]+_/;

function cellToString(val: unknown): string {
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

function isNoiseRow(row: (string | number | null)[]): string | null {
  const cells = row.map((c) => cellToString(c));
  const joined = cells.join(" ");

  for (const pattern of NOISE_PATTERNS) {
    if (pattern.test(joined)) return "Sage/company metadata";
  }
  if (PAGE_PATTERN.test(joined)) return "Page number row";
  if (
    cells[0] === "Totals" ||
    (cells[0].toLowerCase() === "totals" && cells.slice(1).every((c) => c === "" || !isNaN(Number(c))))
  )
    return "Totals row";

  return null;
}

function isHeaderRow(row: (string | number | null)[]): boolean {
  const cells = row.map((c) => cellToString(c).toLowerCase());
  return HEADER_ROW_SIGNATURE.every((sig) => cells.some((c) => c.includes(sig)));
}

function extractUOM(description: string): string {
  const matches = description.match(UOM_REGEX);
  if (!matches || matches.length === 0) return "";
  return matches[matches.length - 1].trim();
}

function stripCountryPrefix(code: string): string {
  return code.replace(COUNTRY_CODE_REGEX, "");
}

function sheetToRows(sheet: XLSX.WorkSheet): (string | number | null)[][] {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const rows: (string | number | null)[][] = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: (string | number | null)[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[cellAddress];
      row.push(cell ? cell.v ?? null : null);
    }
    rows.push(row);
  }
  return rows;
}

export function processWorkbook(buffer: Buffer): {
  workbook: XLSX.WorkBook;
  stats: ProcessingStats;
} {
  const inputWb = XLSX.read(buffer, { type: "buffer", raw: true });
  const sheetName = inputWb.SheetNames[0];
  const inputSheet = inputWb.Sheets[sheetName];

  const allRows = sheetToRows(inputSheet);
  const stats: ProcessingStats = {
    originalRows: allRows.length,
    cleanedRows: 0,
    filteredRows: 0,
    uomMatched: 0,
    prefixesStripped: 0,
  };

  // --- Sheet 2: Cleaned data ---
  const cleanedRows: (string | number | null)[][] = [];
  const filteredRows: FilteredRow[] = [];
  let headerFound = false;

  // Column indices (0-based): Item Code=0, Item Desc=1, Whse=2, Group=3, Category=4, Unit=5, SysQty=6, ActQty=7, Variance=8
  // We keep: 0,1,2,3,4,5 and add UOM as new col 6
  // Drop: 6 (System Qty), 7 (Actual Qty), 8 (Variance)

  for (const row of allRows) {
    // Check noise
    const noiseReason = isNoiseRow(row);
    if (noiseReason) {
      filteredRows.push({ row, reason: noiseReason });
      continue;
    }

    // Check duplicate header
    if (isHeaderRow(row)) {
      if (headerFound) {
        filteredRows.push({ row, reason: "Duplicate header row" });
        continue;
      }
      headerFound = true;
      // Output header row with UOM column added, drop last 3
      const headerOut = row.slice(0, 6);
      headerOut.push("UOM");
      cleanedRows.push(headerOut);
      continue;
    }

    // Data row: strip prefix from item code, extract UOM, drop last 3 cols
    const itemCode = cellToString(row[0]);
    const itemDesc = cellToString(row[1]);

    const stripped = stripCountryPrefix(itemCode);
    if (stripped !== itemCode) stats.prefixesStripped++;

    const uom = extractUOM(itemDesc);
    if (uom) stats.uomMatched++;

    const cleanRow: (string | number | null)[] = [
      stripped || null,
      row[1] ?? null,
      row[2] ?? null,
      row[3] ?? null,
      row[4] ?? null,
      row[5] ?? null,
      uom || null,
    ];
    cleanedRows.push(cleanRow);
  }

  stats.cleanedRows = cleanedRows.length;
  stats.filteredRows = filteredRows.length;

  // --- Build output workbook ---
  const outWb = XLSX.utils.book_new();

  // Sheet 1: Original (deep copy of input sheet)
  const originalSheet: XLSX.WorkSheet = Object.assign({}, inputSheet);
  XLSX.utils.book_append_sheet(outWb, originalSheet, "Original");

  // Sheet 2: Cleaned
  const cleanedSheet = XLSX.utils.aoa_to_sheet(cleanedRows);
  styleHeaderRow(cleanedSheet, cleanedRows[0]?.length ?? 7);
  XLSX.utils.book_append_sheet(outWb, cleanedSheet, "Cleaned");

  // Sheet 3: Filtered Out
  const filteredHeader = [
    "Item Code",
    "Item Description",
    "Whse",
    "Group",
    "Category",
    "Unit",
    "System Qty",
    "Actual Qty",
    "Variance",
    "Reason",
  ];
  const filteredData = [
    filteredHeader,
    ...filteredRows.map((fr) => [...fr.row, fr.reason]),
  ];
  const filteredSheet = XLSX.utils.aoa_to_sheet(filteredData);
  styleHeaderRow(filteredSheet, filteredHeader.length);
  XLSX.utils.book_append_sheet(outWb, filteredSheet, "Filtered Out");

  return { workbook: outWb, stats };
}

function styleHeaderRow(sheet: XLSX.WorkSheet, colCount: number) {
  // Set column widths
  sheet["!cols"] = Array(colCount).fill({ wch: 20 });
}
