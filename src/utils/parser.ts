export interface ParsedItem {
  position: string;
  qty: number;
  description: string; // profile
  length: number; // in mm
  name?: string; // assembly name
  isSheetMetal?: boolean;
  width?: number; // only for sheet metal
}

export interface ParsedAssembly {
  name: string;
  position: string;
  qty: number;
  items: ParsedItem[];
}

/**
 * Parses a copied table from Solidworks BOM / Excel.
 * Handles tab-separated text (standard clipboard format for spreadsheet columns).
 */
export function parseTableText(text: string): ParsedAssembly[] {
  // Parse the TSV/CSV text respecting quoted fields (which can contain newlines)
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          // Double quote inside quotes is a literal quote
          currentField += '"';
          i++;
        } else {
          // Closing quote
          inQuotes = false;
        }
      } else {
        currentField += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === '\t') {
        currentRow.push(currentField);
        currentField = '';
      } else if (char === '\n' || char === '\r') {
        if (char === '\r' && nextChar === '\n') {
          i++;
        }
        currentRow.push(currentField);
        rows.push(currentRow);
        currentRow = [];
        currentField = '';
      } else {
        currentField += char;
      }
    }
  }
  if (currentField || currentRow.length > 0) {
    currentRow.push(currentField);
    rows.push(currentRow);
  }

  // Clean cell values: collapse newlines inside cells to spaces, and trim whitespace
  const cleanRows = rows
    .map((r) => r.map((cell) => cell.replace(/\r?\n/g, ' ').trim()))
    .filter((r) => r.some((cell) => cell.length > 0));

  if (cleanRows.length === 0) return [];

  // Determine headers
  let headerIndex = -1;
  let colMap = {
    position: -1,
    qty: -1,
    description: -1,
    length: -1,
    name: -1, // e.g. დასახელება (Name)
    flatLength: -1,
    flatWidth: -1,
    thickness: -1,
  };

  // Common synonyms for headers (English + Georgian)
  const synonyms = {
    position: ['პოზიცია', 'position', 'pos', 'item no', 'no.', 'item', 'პოზ'],
    qty: ['რაოდ', 'რაოდენობა', 'qty', 'quantity', 'count', 'amount'],
    description: ['აღწერა', 'აღწერილობა', 'description', 'profile', 'size', 'ზომა', 'მასალა'],
    length: ['სიგრძე', 'length', 'len', 'l', 'სიგრძე (მმ)'],
    name: ['დასახელება', 'სახელი', 'name', 'title', 'assembly', 'construction'],
    flatLength: ['განშ. სიგრძე', 'განმ. სიგრძე', 'განშლა სიგრძე', 'flat length', 'flattened length', 'unfolded length'],
    flatWidth: ['განშ. სიგანე', 'განმ. სიგანე', 'განშლა სიგანე', 'flat width', 'flattened width', 'unfolded width'],
    thickness: ['ფ. სისქე', 'ფურცლის სისქე', 'სისქე', 'thickness', 'sheet thickness', 'thk'],
  };

  // Try to find a header row
  for (let r = 0; r < Math.min(cleanRows.length, 5); r++) {
    const row = cleanRows[r].map((cell) => cell.toLowerCase().trim());
    let matches = 0;
    const tempColMap = { ...colMap };
    
    // Check if this row looks like a header
    for (const key of Object.keys(colMap) as (keyof typeof colMap)[]) {
      // 1. Try exact match first to prevent substring collisions (e.g. matching "განმ. სიგრძე" as "სიგრძე")
      let idx = row.findIndex((cell) => 
        synonyms[key].some((syn) => cell === syn)
      );
      
      // 2. Fall back to includes search if no exact match is found
      if (idx === -1) {
        idx = row.findIndex((cell) => 
          synonyms[key].some((syn) => cell.includes(syn))
        );
      }
      
      if (idx !== -1) {
        tempColMap[key] = idx;
        matches++;
      }
    }

    if (matches >= 2) {
      headerIndex = r;
      colMap = tempColMap;
      break;
    }
  }

  // If no header found, make a default mapping based on typical columns
  if (headerIndex === -1) {
    colMap = guessColumns(cleanRows);
  }

  // Slice data rows (after header, if found)
  const dataRows = headerIndex !== -1 ? cleanRows.slice(headerIndex + 1) : cleanRows;

  const assemblies: ParsedAssembly[] = [];
  let currentAssembly: ParsedAssembly | null = null;

  for (const row of dataRows) {
    // Skip empty or short rows
    if (row.length < 2) continue;

    // Get cell values safely
    const getVal = (colIdx: number) => (colIdx >= 0 && colIdx < row.length ? row[colIdx].trim() : '');

    const posRaw = getVal(colMap.position);
    const qtyRaw = getVal(colMap.qty);
    const descRaw = getVal(colMap.description);
    const lenRaw = getVal(colMap.length);
    const nameRaw = getVal(colMap.name);
    const flatLenRaw = getVal(colMap.flatLength);
    const flatWidthRaw = getVal(colMap.flatWidth);
    const thicknessRaw = getVal(colMap.thickness);

    // Skip if position is empty and there's no descriptive data
    if (!posRaw && !descRaw && !lenRaw && !flatLenRaw && !flatWidthRaw) continue;

    const qty = parseInt(qtyRaw.replace(/\s/g, ''), 10) || 1;
    const length = parseFloat(lenRaw.replace(/\s/g, '').replace(',', '.')) || 0;
    const flatLength = parseFloat(flatLenRaw.replace(/\s/g, '').replace(',', '.')) || 0;
    const flatWidth = parseFloat(flatWidthRaw.replace(/\s/g, '').replace(',', '.')) || 0;

    // Determine if it's a parent assembly or a child item
    const isParent = isParentRow(posRaw, descRaw, length, flatLength, flatWidth);

    if (isParent) {
      currentAssembly = {
        name: nameRaw || descRaw || `კონსტრუქცია ${posRaw || assemblies.length + 1}`,
        position: posRaw,
        qty: qty,
        items: [],
      };
      assemblies.push(currentAssembly);
    } else {
      // A sheet metal part is identified by having a non-zero flatWidth
      const isSheetMetal = flatWidth > 0;
      
      // Determine description: if sheet and descRaw is empty, use thickness as profile size (e.g. "3mm Plate")
      let description = descRaw;
      if (!description) {
        if (isSheetMetal) {
          description = thicknessRaw ? `${thicknessRaw}მმ ფურცელი` : 'ფურცლოვანი ლითონი';
        } else {
          description = 'უცნობი პროფილი';
        }
      }

      // Determine cut length: for sheet metal, it is the flatLength (or regular length if flatLength is empty)
      const cutLength = isSheetMetal ? (flatLength || length) : length;

      const item: ParsedItem = {
        position: posRaw,
        qty: qty,
        description,
        length: cutLength,
        isSheetMetal,
        width: isSheetMetal ? flatWidth : undefined,
      };

      if (!currentAssembly) {
        // Create a default assembly if items appear before any parent assembly
        currentAssembly = {
          name: 'საერთო კრებული',
          position: '1',
          qty: 1,
          items: [],
        };
        assemblies.push(currentAssembly);
      }
      currentAssembly.items.push(item);
    }
  }

  // Cleanup: if any assembly has no items and there are other assemblies with items,
  // we keep it, but let's filter out empty assemblies if they are redundant.
  return assemblies.filter((ass) => ass.items.length > 0 || assemblies.length === 1);
}

function isParentRow(pos: string, desc: string, length: number, flatLength: number, flatWidth: number): boolean {
  // If position is a flat integer like "1", "2" and all lengths are 0, it's a parent
  if (pos && /^\d+$/.test(pos.trim()) && length === 0 && flatLength === 0 && flatWidth === 0) {
    return true;
  }
  // If it doesn't have a profile description (does not contain 'x' or 'x' and is long text), and no length
  if (desc && !desc.includes('x') && length === 0 && flatLength === 0 && flatWidth === 0) {
    return true;
  }
  return false;
}

function guessColumns(rows: string[][]): {
  position: number;
  qty: number;
  description: number;
  length: number;
  name: number;
  flatLength: number;
  flatWidth: number;
  thickness: number;
} {
  const map = { position: 0, qty: 1, name: 2, description: 3, length: 8, flatLength: -1, flatWidth: -1, thickness: -1 }; // Typical fallback
  
  // Let's analyze the first data row to guess
  const sampleRow = rows.find(r => r.length >= 3);
  if (!sampleRow) return map;

  let posIdx = 0;
  let qtyIdx = 1;
  let nameIdx = -1;
  let descIdx = -1;
  let lenIdx = -1;

  for (let i = 0; i < sampleRow.length; i++) {
    const val = sampleRow[i].trim();
    // Position: typically short, first column, dotted number (e.g. 1.1) or integer
    if (i === 0 && (/^\d+(\.\d+)?$/.test(val))) {
      posIdx = i;
    }
    // Quantity: integer
    else if (i === 1 && /^\d+$/.test(val)) {
      qtyIdx = i;
    }
    // Profile description: contains dimensions like "50 x 50" or "50x30"
    else if (/\d+\s*x\s*\d+/.test(val)) {
      descIdx = i;
    }
    // Length: number, usually > 100
    else if (i > 2 && /^\d+$/.test(val) && parseInt(val, 10) > 100) {
      if (lenIdx === -1) {
        lenIdx = i;
      }
    }
    // Name: text
    else if (i === 2 && val.length > 3 && !/\d+/.test(val)) {
      nameIdx = i;
    }
  }

  return {
    position: posIdx,
    qty: qtyIdx !== -1 ? qtyIdx : 1,
    name: nameIdx !== -1 ? nameIdx : (descIdx !== -1 && descIdx > 2 ? 2 : -1),
    description: descIdx !== -1 ? descIdx : 3,
    length: lenIdx !== -1 ? lenIdx : 8,
    flatLength: -1,
    flatWidth: -1,
    thickness: -1,
  };
}
