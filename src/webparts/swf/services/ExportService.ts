import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable'; // Modern functional import
// eslint-disable-next-line @typescript-eslint/no-var-requires
const JSZip = require('jszip');

export class ExportService {

    // ─── Excel ────────────────────────────────────────────────────────────────
    public static exportToExcel(data: Record<string, unknown>[], fileName: string): void {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, `${fileName}.xlsx`);
    }

    // ─── CSV ──────────────────────────────────────────────────────────────────
    public static exportToCSV(data: Record<string, unknown>[], fileName: string): void {
        const ws = XLSX.utils.json_to_sheet(data);
        const csv = XLSX.utils.sheet_to_csv(ws);
        ExportService._download(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), `${fileName}.csv`);
    }

    // ─── PDF ──────────────────────────────────────────────────────────────────
    public static exportToPDF(
        data: Record<string, unknown>[],
        fileName: string,
        columnHeaders: string[],
        columnKeys: string[]
    ): void {
        try {
            // A4 landscape is 297mm wide
            const doc = new jsPDF({ orientation: 'l', unit: 'mm', format: 'a4' });

            doc.setFontSize(14);
            doc.setTextColor(14, 77, 146);
            doc.text(fileName, 14, 14);

            const body: string[][] = data.map(row =>
                columnKeys.map(k => String(row[k] == null ? '' : row[k]))
            );

            // Functional call to autoTable(doc, options)
            autoTable(doc, {
                head: [columnHeaders],
                body,
                startY: 20,
                styles: { 
                    fontSize: 5, // Reduced font to fit many columns
                    cellPadding: 1.5, 
                    overflow: 'linebreak',
                    halign: 'left',
                    valign: 'middle'
                },
                headStyles: { 
                    fillColor: [14, 77, 146], 
                    textColor: 255, 
                    fontStyle: 'bold',
                    lineWidth: 0.1,
                    lineColor: [255, 255, 255]
                },
                alternateRowStyles: { fillColor: [245, 248, 255] },
                margin: { left: 8, right: 8, bottom: 10 },
                theme: 'striped',
                columnStyles: {
                    0: { cellWidth: 10 }, // ID
                    1: { cellWidth: 25 }  // Subject
                }
            });

            doc.save(`${fileName.replace(/\s+/g, '_')}.pdf`);
        } catch (e) {
            console.error('[ExportService] PDF export error:', e);
            alert('PDF export failed. See console for details.');
        }
    }

    // ─── ZIP (Excel + CSV + PDF bundled) ─────────────────────────────────────
    public static async exportToZip(
        data: Record<string, unknown>[],
        fileName: string,
        columnHeaders: string[],
        columnKeys: string[]
    ): Promise<void> {
        try {
            const zip = new JSZip();
            const timestamp = new Date().getTime();

            // 1. Excel (data.xlsx)
            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Data');
            const excelBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            zip.file('data.xlsx', excelBytes);

            // 2. CSV (data.csv)
            zip.file('data.csv', XLSX.utils.sheet_to_csv(ws));

            // 3. PDF (data.pdf)
            const doc = new jsPDF({ orientation: 'l', unit: 'mm', format: 'a4' });
            doc.setFontSize(14);
            doc.setTextColor(14, 77, 146);
            doc.text(fileName, 14, 14);

            const body: string[][] = data.map(row =>
                columnKeys.map(k => String(row[k] == null ? '' : row[k]))
            );

            autoTable(doc, {
                head: [columnHeaders],
                body,
                startY: 20,
                styles: { fontSize: 5, cellPadding: 1.5, overflow: 'linebreak' },
                headStyles: { fillColor: [14, 77, 146], textColor: 255, fontStyle: 'bold' },
                margin: { left: 8, right: 8 },
            });

            const pdfBytes = doc.output('arraybuffer');
            zip.file('data.pdf', pdfBytes);

            // 4. Generate + download ZIP (export_<timestamp>.zip)
            const blob: Blob = await zip.generateAsync({ type: 'blob' });
            ExportService._download(blob, `export_${timestamp}.zip`);
        } catch (e) {
            console.error('[ExportService] ZIP export error:', e);
            alert('ZIP export failed. See console for details.');
        }
    }

    // ─── Download helper ──────────────────────────────────────────────────────
    private static _download(blob: Blob, fileName: string): void {
        try {
            const url = URL.createObjectURL(blob);
            const a   = document.createElement('a');
            a.href     = url;
            a.download = fileName;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                if (a.parentNode) document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 500);
        } catch (e) {
            console.error('[ExportService] Download trigger failed:', e);
        }
    }
}
