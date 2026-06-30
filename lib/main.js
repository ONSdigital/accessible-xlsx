// Workaround to override default font + size in ExcelJS
// See issue: https://github.com/exceljs/exceljs/issues/572#issuecomment-631788521
import StylesXform from "@protobi/exceljs/lib/xlsx/xform/style/styles-xform.js";

const defaultFont = {
	size: 12,
	color: { theme: 1 },
	name: "Arial",
	family: 2,
	scheme: "minor"
};
const origStylesXformInit = StylesXform.prototype.init;
StylesXform.prototype.init = function () {
	origStylesXformInit.apply(this, arguments);
	this._addFont(defaultFont);
};

import ExcelJS from "@protobi/exceljs";

function addTextRow(sheet, text, options = {}) {
	const row = sheet.getRow(sheet.rowCount + 1);
	const cell = row.getCell(1);
	if (text.startsWith("# ")) {
		cell.value = text.slice(2);
		cell.font = { ...defaultFont, size: 18, bold: true };
	} else if (text.startsWith("## ")) {
		row.height = 40;
		cell.value = text.slice(3);
		cell.font = { ...defaultFont, size: 14, bold: true };
	} else if (text.startsWith("[")) {
		cell.value = {
			text: text.match(/(?<=\[).*(?=\])/)[0],
			hyperlink: text.match(/(?<=\().*(?=\))/)[0]
		};
		cell.font = { ...defaultFont, underline: true, color: { argb: "0000FF" } };
	} else {
		cell.value = text;
	}
	if (options.height) row.height = options.height;
	if (options.alignment) row.alignment = options.alignment;
}

function getColFormat(column) {
	if (["number_with_commas", "number_0dp"].includes(column.style)) return "#,##0";
	if (/^number_\d+dp$/.test(column.style)) {
		const dp = +column.style.match(/\d+/)[0];
		return "#,##0.".padEnd(dp + 6, "0");
	}
	return null;
}

function getColWidth(values = null) {
	const maxColWidth = 28;
	const minColWidth = 12;

	if (!values) return minColWidth;

	let maxLength = 0;
	for (let i = 0; i < values.length; i++) {
		const length = String(values[i]).length;
		if (length > maxLength) maxLength = length;
	}
	return maxLength < minColWidth
		? minColWidth
		: maxLength > maxColWidth
			? maxColWidth
			: maxLength;
}

function getDataRows(sheet) {
	if (sheet.rows && Array.isArray(sheet.rows?.[0])) return sheet.rows;
	if (sheet.rows) {
		const cols = sheet.columns.map((c) => c.key || c.heading);
		return sheet.rows.map((d) => cols.map((c) => d[c]));
	}
	const rowCount = Math.max(...sheet.columns.map((c) => c.values.length));
	return [...Array(rowCount).keys()].map((i) => sheet.columns.map((c) => c.values[i] || null));
}

function slugify(string) {
	return string
		.trim()
		.replace(/[\s_]+/g, "_")
		.match(/[A-Za-z0-9_]/g)
		.join("");
}

export default async function accessibleXLSX(data) {
	const workbook = new ExcelJS.Workbook();
	const oneTableMessage = "This worksheet contains one table.";
	const isSingleSheet = !data.sheets;
	const creator = data.creator || "Anonymous";
	const created = data.created || new Date();

	workbook.title = isSingleSheet ? data.sheetName : data.coverSheetTitle;
	workbook.creator = creator;
	workbook.lastModifiedBy = creator;
	workbook.created = created;
	workbook.modified = created;

	if (!isSingleSheet) {
		const coverSheet = workbook.addWorksheet("Cover_sheet");
		coverSheet.columns = [{ width: 80, style: { alignment: { wrapText: true } } }];
		addTextRow(coverSheet, `# ${data.coverSheetTitle}`);

		for (let i = 0; i < data.coverSheetContents.length; i++) {
			addTextRow(coverSheet, data.coverSheetContents[i]);
		}

		const contentsSheet = workbook.addWorksheet("Table_of_contents");
		contentsSheet.columns = [{ width: 10 }, { width: 70 }];
		addTextRow(contentsSheet, `# Table of contents`);
		addTextRow(contentsSheet, oneTableMessage, {
			height: 40,
			alignment: { vertical: "top" }
		});
		contentsSheet.addTable({
			name: "table_of_contents",
			ref: "A3",
			headerRow: true,
			style: {
				theme: null,
				showRowStripes: false
			},
			columns: [
				{
					name: "Table",
					style: { font: { underline: true, color: { argb: "0000FF" } } }
				},
				{ name: "Name", style: { alignment: { wrapText: true } } }
			],
			rows: data.sheets.map((d, i) => [
				{
					text: `Table ${i + 1}`,
					hyperlink: `#'${i + 1}'!A1`
				},
				d.sheetName
			])
		});
		contentsSheet.getRow(3).font = { ...defaultFont, bold: true };

		if (data.notes.length) {
			const notesSheet = workbook.addWorksheet("Notes");
			notesSheet.columns = [{ width: 10 }, { width: 70 }];
			addTextRow(notesSheet, `# Notes`);
			addTextRow(notesSheet, oneTableMessage, {
				height: 40,
				alignment: { vertical: "top" }
			});
			notesSheet.addTable({
				name: "notes",
				ref: "A3",
				headerRow: true,
				style: {
					theme: null,
					showRowStripes: false
				},
				columns: [
					{ name: "Number" },
					{ name: "Note", style: { alignment: { wrapText: true } } }
				],
				rows: data.notes.map((n) => [n.name, n.text])
			});
			notesSheet.getRow(3).font = { ...defaultFont, bold: true };
		}
	}

	const sheets = isSingleSheet ? [data] : data.sheets;

	for (let i = 0; i < sheets.length; i++) {
		const s = sheets[i];
		const rows = getDataRows(s);
		const sheet = workbook.addWorksheet(String(i + 1));

		addTextRow(sheet, `# ${s.sheetName}`);
		for (let j = 0; j < s.sheetIntroText?.length || 0; j++) {
			addTextRow(
				sheet,
				`${s.sheetIntroText[j]}`,
				j === s.sheetIntroText.length - 1
					? { height: 40, alignment: { vertical: "top" } }
					: {}
			);
		}

		const tableRowNumber = sheet.rowCount + 1;
		sheet.addTable({
			name: s.tableName || slugify(s.sheetName),
			ref: `A${tableRowNumber}`,
			headerRow: true,
			style: {
				theme: null,
				showRowStripes: false
			},
			columns: s.columns.map((c) => ({ name: c.heading })),
			rows
		});
		sheet.getRow(tableRowNumber).font = { ...defaultFont, bold: true };
		sheet.getRow(tableRowNumber).alignment = { wrapText: true };

		for (let i = 0; i < s.columns.length; i++) {
			const meta = s.columns[i];
			const col = sheet.getColumn(i + 1);
			const colFormat = getColFormat(s.columns[i]);
			if (colFormat) {
				col.numFmt = colFormat;
				col.width = getColWidth();
			} else {
				col.width = getColWidth(rows.map((d) => d[i]));
			}
		}
	}

	workbook.views = [{ activeTab: 0, activeCell: "A1" }];

	return workbook.xlsx.writeBuffer();
}
