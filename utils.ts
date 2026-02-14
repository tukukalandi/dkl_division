import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { RawBookingRow, MasterDataRow } from './types';
import { PRODUCTS } from './constants';

// Default fallback if fetch fails
import { MASTER_DATA as DEFAULT_MASTER_DATA } from './constants';

export const getProductCategory = (productName: string): 'Domestic' | 'Parcel' | 'International' | 'Other' => {
  const name = productName.toLowerCase().trim();
  
  const isDomestic = PRODUCTS.DOMESTIC.some(p => p.toLowerCase().trim() === name);
  if (isDomestic) return 'Domestic';
  
  const isParcel = PRODUCTS.PARCEL.some(p => p.toLowerCase().trim() === name);
  if (isParcel) return 'Parcel';
  
  const isInternational = PRODUCTS.INTERNATIONAL.some(p => p.toLowerCase().trim() === name);
  if (isInternational) return 'International';
  
  return 'Other';
};

// Helper to safely parse numbers that might contain commas, currency symbols, or be strings
const parseNumber = (val: any): number => {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') {
     return isNaN(val) ? 0 : val;
  }
  
  // Convert to string and trim
  let strVal = String(val).trim();
  
  if (!strVal) return 0;

  // Remove commas (thousands separator)
  strVal = strVal.replace(/,/g, '');

  // Remove any remaining characters that are NOT digits, decimal point, or minus sign
  // This handles cases like "Rs. 1200", "₹ 1200", "1200/-"
  strVal = strVal.replace(/[^0-9.-]/g, '');

  const num = parseFloat(strVal);
  return isNaN(num) ? 0 : num;
};

export const parseExcelFile = (file: File): Promise<RawBookingRow[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      if (!data) return reject("No data read");
      
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

      const parsedData: RawBookingRow[] = [];

      jsonData.forEach((row, index) => {
        // Skip header row if it contains text in the numeric columns
        if (index === 0) {
           const rawVal = String(row[3] || '').replace(/,/g, '');
           const potentialNumber = parseFloat(rawVal);
           if (isNaN(potentialNumber)) return;
        }

        const officeId = row[0] ? String(row[0]).trim() : '';
        if (!officeId) return; // Skip empty rows

        const entry: RawBookingRow = {
          officeId: officeId,
          officeName: row[1] ? String(row[1]).trim() : '',
          productName: row[2] ? String(row[2]).trim() : '',
          articles: parseNumber(row[3]), // Col D
          postage: parseNumber(row[4]), // Col E
          vas: parseNumber(row[5]),
          tax: parseNumber(row[6]),
          prepaidFm: parseNumber(row[7]),
          prepaidPs: parseNumber(row[8]),
          prepaidSs: parseNumber(row[9]),
          totalAmount: parseNumber(row[10]), // Col K (Revenue)
          avgWeight: parseNumber(row[11]),
        };
        parsedData.push(entry);
      });

      resolve(parsedData);
    };
    reader.onerror = (err) => reject(err);
    reader.readAsBinaryString(file);
  });
};

const parseCSVLine = (text: string): string[] => {
  const result: string[] = [];
  let current = '';
  let inQuote = false;
  
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"') {
      inQuote = !inQuote;
    } else if (char === ',' && !inQuote) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result.map(field => {
    if (field.startsWith('"') && field.endsWith('"')) {
      field = field.slice(1, -1);
    }
    return field.replace(/""/g, '"');
  });
};

export const fetchMasterDataFromUrl = async (url: string): Promise<MasterDataRow[]> => {
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error('Network response was not ok');
    const text = await response.text();
    
    const lines = text.split(/\r?\n/);
    const masterData: MasterDataRow[] = [];
    
    lines.forEach((line, index) => {
      if (index === 0) return; // Skip header
      if (!line.trim()) return;

      const row = parseCSVLine(line);
      if (!row[0]) return;

      masterData.push({
        officeId: row[0].trim(),
        officeName: row[1] ? row[1].trim() : '',
        officeType: row[2] ? row[2].trim() : '',
        officeJurisdiction: row[3] ? row[3].trim() : '',
        headOfficeName: row[4] ? row[4].trim() : '',
        subDivisionName: row[5] ? row[5].trim() : '',
        areaType: row[6] ? row[6].trim() : 'Rural', // Col G (index 6)
      });
    });
    return masterData;
  } catch (error) {
    console.error("Failed to fetch master data", error);
    return DEFAULT_MASTER_DATA;
  }
};

export const fetchBookingDataFromUrl = async (url: string): Promise<RawBookingRow[]> => {
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error('Failed to fetch report data');
    const text = await response.text();
    
    const lines = text.split(/\r?\n/);
    const parsedData: RawBookingRow[] = [];
    
    lines.forEach((line, index) => {
      if (!line.trim()) return;
      const row = parseCSVLine(line);
      
      if (index === 0) {
         const rawVal = String(row[3] || '').replace(/,/g, '');
         const potentialNumber = parseFloat(rawVal);
         if (isNaN(potentialNumber)) return;
      }

      const officeId = row[0] ? String(row[0]).trim() : '';
      if (!officeId || officeId.toLowerCase() === 'office id') return;

      const entry: RawBookingRow = {
        officeId: officeId,
        officeName: row[1] ? String(row[1]).trim() : '',
        productName: row[2] ? String(row[2]).trim() : '',
        articles: parseNumber(row[3]),
        postage: parseNumber(row[4]),
        vas: parseNumber(row[5]),
        tax: parseNumber(row[6]),
        prepaidFm: parseNumber(row[7]),
        prepaidPs: parseNumber(row[8]),
        prepaidSs: parseNumber(row[9]),
        totalAmount: parseNumber(row[10]),
        avgWeight: parseNumber(row[11]),
      };
      parsedData.push(entry);
    });
    
    return parsedData;
  } catch (error) {
    console.error("Failed to fetch booking data", error);
    throw error;
  }
};

export const formatCurrency = (val: number) => {
  return new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR' }).format(val);
};

export const exportToPDF = (headers: string[], data: any[][], title: string, landscape = false) => {
  const doc = new jsPDF(landscape ? 'l' : 'p', 'mm', 'a4');
  doc.text(title, 14, 15);
  autoTable(doc, {
    head: [headers],
    body: data,
    startY: 20,
    theme: 'grid',
    styles: { fontSize: 8 },
    headStyles: { fillColor: [206, 32, 41] } // India Post Red
  });
  doc.save(`${title.replace(/\s+/g, '_')}.pdf`);
};

export const exportToExcel = (data: any[], fileName: string) => {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
};