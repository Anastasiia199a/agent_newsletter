// Создаёт пример data/aktionen.xlsx. Цены вымышленные — замените своими данными.
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const today = new Date();
const d = (days) => new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate() + days));

const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Aktionen');
ws.columns = [
  { header: 'Produkt', key: 'produkt', width: 34 },
  { header: 'Marke', key: 'marke', width: 14 },
  { header: 'Alter Preis', key: 'alt', width: 12, style: { numFmt: '#,##0.00 €' } },
  { header: 'Neuer Preis', key: 'neu', width: 12, style: { numFmt: '#,##0.00 €' } },
  { header: 'Gültig von', key: 'von', width: 12, style: { numFmt: 'dd.mm.yyyy' } },
  { header: 'Gültig bis', key: 'bis', width: 12, style: { numFmt: 'dd.mm.yyyy' } },
  { header: 'Shop', key: 'shop', width: 14 },
  { header: 'Bedingungen', key: 'bed', width: 28 },
  { header: 'Link', key: 'link', width: 50 },
  { header: 'Bild', key: 'bild', width: 30 },
];
ws.getRow(1).font = { bold: true };
ws.addRows([
  { produkt: 'Red Bull Halfzware Shag 40g', marke: 'Red Bull', alt: 9.1, neu: 8.2, von: d(0), bis: d(14), shop: 'Online-Shop', bed: 'Ab 2 Packungen', link: 'https://onlinetabakshop.de/p/red-bull-halfzware-shag-40g-tabak-137' },
  { produkt: 'Beispiel Drehtabak 30g', marke: 'Marke A', alt: 7.5, neu: 6.9, von: d(0), bis: d(10), shop: 'Online-Shop', bed: '', link: '' },
  { produkt: 'Beispiel Filterhülsen 200 Stk.', marke: 'Marke B', alt: 2.99, neu: 2.29, von: d(-3), bis: d(20), shop: 'Online-Shop', bed: 'Solange Vorrat reicht', link: '' },
  { produkt: 'Beispiel Zigarillos 10er', marke: 'Marke C', alt: 5.4, neu: 4.9, von: d(0), bis: d(8), shop: 'Filiale Berlin', bed: '', link: '' },
  { produkt: 'Abgelaufene Aktion (wird ignoriert)', marke: 'Marke D', alt: 6, neu: 5, von: d(-30), bis: d(-2), shop: 'Online-Shop', bed: '', link: '' },
  { produkt: 'Fehlerhafte Zeile (Preis zu hoch)', marke: 'Marke E', alt: 5, neu: 5.5, von: d(0), bis: d(10), shop: 'Online-Shop', bed: '', link: '' },
]);
await wb.xlsx.writeFile(path.join(ROOT, 'data', 'aktionen.xlsx'));
console.log('Создан data/aktionen.xlsx');
