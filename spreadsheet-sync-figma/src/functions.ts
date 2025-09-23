import { SpreadsheetSheetsData } from './models';

export function formatDataFillEmpty(data: SpreadsheetSheetsData): SpreadsheetSheetsData {
  const newData = data.map(sheet => {
    let maxCells = sheet[0].length;
    return sheet.map(elem => elem.length < maxCells ? elem = [...elem, ...Array(maxCells - elem.length).fill('').map((_, i) => '')] : elem);
  });

  return newData;
}

// code chunk from (https://github.com/yuanqing/create-figma-plugin)
export function showLoadingNotification(message: string): () => void {
  const notificationHandler = figma.notify(message, {
    timeout: 60000
  });

  return (): void => {
    notificationHandler.cancel();
  }
}

export function formatMsToTime(ms: number): string {
  const pad = (num: number, size: number = 2) => `00${num}`.slice(-size);

  const hours: string = pad(ms / 3.6e6 | 0);
  const minutes: string = pad((ms % 3.6e6) / 6e4 | 0);
  const seconds: string = pad((ms % 6e4) / 1000 | 0);
  const milliseconds: string = pad(ms % 1000, 3);
  return `${minutes}:${seconds}s`;
}

export function formatTimeToMs(hrs: number = 0, min: number = 0, sec: number = 0) {
  return ((hrs * 60 * 60 + min * 60 + sec) * 1000);
}

// csv --------------------

export function isCsvUrl(url: string): boolean {
  if (!url) return false;
  const trimmed = url.trim().toLowerCase();
  if (trimmed.startsWith('data:text/csv')) return true;
  try {
    const u = new URL(url);
    return u.pathname.toLowerCase().endsWith('.csv');
  } catch {
    return false;
  }
}

export function parseCsv(csvText: string): string[][] {
  const rows: string[][] = [];
  if (!csvText) return rows;

  // Normalize newlines
  const text = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  let i = 0;
  const len = text.length;
  let cur: string[] = [];
  let field = '';
  let inQuotes = false;

  while (i < len) {
    const char = text[i];

    if (inQuotes) {
      if (char === '"') {
        // Escaped quote
        if (i + 1 < len && text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        } else {
          inQuotes = false;
          i++;
          continue;
        }
      } else {
        field += char;
        i++;
        continue;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
        i++;
        continue;
      }
      if (char === ',') {
        cur.push(field);
        field = '';
        i++;
        continue;
      }
      if (char === '\n') {
        cur.push(field);
        rows.push(cur);
        cur = [];
        field = '';
        i++;
        continue;
      }
      field += char;
      i++;
    }
  }
  // push last field
  cur.push(field);
  rows.push(cur);

  return rows;
}
