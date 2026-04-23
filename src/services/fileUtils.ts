import * as XLSX from 'xlsx';
import { PDFDocument, PDFTextField, PDFCheckBox } from 'pdf-lib';
import JSZip from 'jszip';

export interface ExcelData {
  headers: string[];
  rows: any[];
  sheetNames: string[];
  selectedSheet: string;
}

export interface PDFFieldInfo {
  name: string;
  type: 'text' | 'checkbox' | 'other';
}

export const parseExcel = async (file: File, sheetName?: string): Promise<ExcelData> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: false, // Keep dates as serial numbers/strings to use Excel's formatting
          cellNF: true,    // Preserve number formats
          cellText: true   // Generate formatted text (the 'w' property)
        });
        const sheetNames = workbook.SheetNames;
        const targetSheetName = sheetName || sheetNames[0];
        const worksheet = workbook.Sheets[targetSheetName];
        
        if (!worksheet) {
          throw new Error(`Sheet "${targetSheetName}" not found`);
        }

        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
        
        if (jsonData.length === 0) {
          throw new Error('Selected sheet is empty');
        }

        const headers = Object.keys(jsonData[0] as object);
        resolve({ headers, rows: jsonData, sheetNames, selectedSheet: targetSheetName });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const parseMappingExcel = async (file: File): Promise<{ mapping: Record<string, string[]>; firstDocField: string }> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets['Sharp'];
        
        if (!worksheet) {
          throw new Error('Worksheet "Sharp" not found in the mapping file.');
        }

        const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];
        const mapping: Record<string, string[]> = {};
        let firstDocField = '';

        if (jsonData.length > 0 && jsonData[0]['Docfield']) {
          firstDocField = String(jsonData[0]['Docfield']).trim();
        }

        jsonData.forEach(row => {
          const docField = row['Docfield'];
          const assigned = row['Assigned'];
          if (docField && assigned) {
            const docFieldStr = String(docField).trim();
            const assignedStr = String(assigned).trim();
            
            if (!mapping[docFieldStr]) {
              mapping[docFieldStr] = [];
            }
            if (!mapping[docFieldStr].includes(assignedStr)) {
              mapping[docFieldStr].push(assignedStr);
            }
          }
        });

        resolve({ mapping, firstDocField });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const parsePDFFields = async (file: File): Promise<PDFFieldInfo[]> => {
  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await PDFDocument.load(arrayBuffer);
  const form = pdfDoc.getForm();
  const fields = form.getFields();

  return fields.map((field) => {
    const name = field.getName();
    let type: PDFFieldInfo['type'] = 'other';
    if (field instanceof PDFTextField) type = 'text';
    if (field instanceof PDFCheckBox) type = 'checkbox';
    return { name, type };
  });
};

export const sanitizeFilename = (name: string): string => {
  return name.replace(/[/\\?%*:|"<>]/g, '-').trim();
};

export interface GeneratedBatch {
  zip: Blob;
  files: { name: string; blob: Blob }[];
}

export const generateBatch = async (
  pdfFile: File,
  excelData: ExcelData,
  mapping: Record<string, string[]>,
  filenamePattern: string,
  onProgress: (current: number, total: number) => void
): Promise<GeneratedBatch> => {
  const zip = new JSZip();
  const pdfBytes = await pdfFile.arrayBuffer();
  const total = excelData.rows.length;
  const files: { name: string; blob: Blob }[] = [];

  for (let i = 0; i < total; i++) {
    const row = excelData.rows[i];
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const form = pdfDoc.getForm();

    // Fill fields
    for (const [pdfFieldName, excelHeaders] of Object.entries(mapping)) {
      const field = form.getField(pdfFieldName);
      const values = excelHeaders.map(header => row[header]);
      
      let combinedValue = values.join(', ');

      // Special logic: if Anst.nr. contains non-numeric value, leave Text1.0.7 blank
      if (pdfFieldName === 'Text1.0.7') {
        const anstNr = row['Anst.nr.'];
        if (anstNr !== undefined && anstNr !== null && anstNr !== '') {
          const isNumeric = /^\d+$/.test(String(anstNr).trim());
          if (!isNumeric) {
            combinedValue = '';
          }
        }
      }

      if (field instanceof PDFTextField) {
        field.setText(String(combinedValue));
      } else if (field instanceof PDFCheckBox) {
        const lowerVal = String(combinedValue).toLowerCase();
        if (lowerVal === 'true' || lowerVal === 'yes' || lowerVal === '1' || lowerVal === 'checked') {
          field.check();
        } else {
          field.uncheck();
        }
      }
    }

    // Generate filename
    let filename = filenamePattern;
    excelData.headers.forEach(header => {
      const val = row[header];
      filename = filename.replace(new RegExp(`\\[${header}\\]`, 'g'), String(val));
    });

    if (!filename.toLowerCase().endsWith('.pdf')) {
      filename += '.pdf';
    }
    filename = sanitizeFilename(filename);

    const savedPdfBytes = await pdfDoc.save();
    const pdfBlob = new Blob([savedPdfBytes], { type: 'application/pdf' });
    
    zip.file(filename, savedPdfBytes);
    files.push({ name: filename, blob: pdfBlob });

    onProgress(i + 1, total);
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  return { zip: zipBlob, files };
};
