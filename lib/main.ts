import { Workbook, Worksheet, Table, Row, Column, Cell } from "documonster/excel";

type cellValue = string | number | null;
type dataRow = cellValue[];
type dataObjectRow = { [key: string]: cellValue };
type sheetColumn = { style: string; heading: string; key?: string; values?: cellValue[] };
type sheetData = {
	sheetName: string;
	tableName?: string;
	sheetIntroText: string[];
	columns: sheetColumn[];
	rows?: dataRow[] & dataObjectRow[];
};
type tableData = {
	coverSheetTitle?: string;
	coverSheetContents?: string[];
	creator?: string;
	created?: Date;
	notes?: { name: string; text: string }[];
	sheets: sheetData[];
};

const defaultFont = {
	size: 12,
	color: { theme: 1 },
	name: "Arial",
	family: 2
};
const defaultView = {
	x: 0,
	y: 0,
	width: 10000,
	height: 20000,
	firstSheet: 0,
	activeTab: 0,
	visibility: "visible"
};

// @ts-ignore -- need WorksheetData type from documonster
function addTextRow(sheet, text: string, options: { height?: number; alignment?: {} } = {}) {
	const rowNum = Worksheet.rowCount(sheet) + 1;
	const cellAddr = `A${rowNum}`;
	if (text.startsWith("# ")) {
		Cell.setValue(sheet, cellAddr, text.slice(2));
		Cell.setFont(sheet, cellAddr, { ...defaultFont, size: 18, bold: true });
	} else if (text.startsWith("## ")) {
		Cell.setValue(sheet, cellAddr, text.slice(3));
		Cell.setFont(sheet, cellAddr, { ...defaultFont, size: 14, bold: true });
		Row.setHeight(sheet, rowNum, 40);
	} else if (text.startsWith("[")) {
		Cell.setValue(sheet, cellAddr, {
			text: text.match(/(?<=\[).*(?=\])/)?.[0] || "",
			hyperlink: text.match(/(?<=\().*(?=\))/)?.[0] || ""
		});
		Cell.setFont(sheet, cellAddr, {
			...defaultFont,
			underline: true,
			color: { argb: "0000FF" }
		});
	} else {
		Cell.setValue(sheet, cellAddr, text);
	}
	if (options.height) Row.setHeight(sheet, rowNum, options.height);
	if (options.alignment) Row.setAlignment(sheet, rowNum, options.alignment);
}

function getColFormat(column: sheetColumn) {
	if (["number_with_commas", "number_0dp"].includes(column.style)) return "#,##0";
	if (/^number_\d+dp$/.test(column.style)) {
		const dp = +(column.style.match(/\d+/)?.[0] || 0);
		return "#,##0.".padEnd(dp + 6, "0");
	}
	return null;
}

function getColWidth(values: cellValue[] | null = null) {
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

function getDataRows(sheet: sheetData): dataRow[] {
	if (sheet.rows && Array.isArray(sheet.rows?.[0])) return sheet.rows;
	if (sheet.rows) {
		const keys = sheet.columns.map((c) => c.key || c.heading || "");
		// @ts-ignore
		return sheet.rows.map((d) => keys.map((key) => d[key]));
	}
	// @ts-ignore
	const rowCount = Math.max(...sheet.columns.map((c) => c.values.length));
	// @ts-ignore
	return [...Array(rowCount).keys()].map((i) => sheet.columns.map((c) => c.values[i] || null));
}

function slugify(string: string) {
	return (
		string
			.trim()
			.replace(/[\s_]+/g, "_")
			.match(/[A-Za-z0-9_]/g) || []
	).join("");
}

export default async function accessibleXLSX(data: tableData & sheetData) {
	const workbook = Workbook.create();
	const oneTableMessage = "This worksheet contains one table.";
	const isSingleSheet = !data.sheets;

	if (!isSingleSheet) {
		const coverSheet = Workbook.addWorksheet(workbook, "Cover_sheet");
		Worksheet.setColumns(coverSheet, [{ width: 80, style: { alignment: { wrapText: true } } }]);
		addTextRow(coverSheet, `# ${data.coverSheetTitle || "Cover sheet"}`);

		for (const content of data.coverSheetContents || []) {
			addTextRow(coverSheet, content);
		}

		const contentsSheet = Workbook.addWorksheet(workbook, "Table_of_contents");
		Worksheet.setColumns(contentsSheet, [{ width: 10 }, { width: 70 }]);
		addTextRow(contentsSheet, `# Table of contents`);
		addTextRow(contentsSheet, oneTableMessage, {
			height: 40,
			alignment: { vertical: "top" }
		});
		Table.add(contentsSheet, {
			name: "table_of_contents",
			ref: "A3",
			headerRow: true,
			columns: [
				{
					name: "Table",
					style: { font: { ...defaultFont, underline: true, color: { argb: "0000FF" } } },
					filterButton: false
				},
				{ name: "Name", style: { alignment: { wrapText: true } }, filterButton: false }
			],
			rows: data.sheets.map((d, i) => [
				{
					text: `Table ${i + 1}`,
					hyperlink: `#'${i + 1}'!A1`
				},
				d.sheetName
			])
		});
		Row.setFont(contentsSheet, 3, { ...defaultFont, bold: true });

		if (data.notes?.length) {
			const notesSheet = Workbook.addWorksheet(workbook, "Notes");
			Worksheet.setColumns(notesSheet, [{ width: 10 }, { width: 70 }]);
			addTextRow(notesSheet, `# Notes`);
			addTextRow(notesSheet, oneTableMessage, {
				height: 40,
				alignment: { vertical: "top" }
			});
			Table.add(notesSheet, {
				name: "notes",
				ref: "A3",
				headerRow: true,
				columns: [
					{ name: "Number", filterButton: false },
					{ name: "Note", style: { alignment: { wrapText: true } }, filterButton: false }
				],
				rows: data.notes.map((n) => [n.name, n.text])
			});
			Row.setFont(notesSheet, 3, { ...defaultFont, bold: true });
		}
	}

	const sheets = isSingleSheet ? [data] : data.sheets;

	for (let i = 0; i < sheets.length; i++) {
		const s = sheets[i];
		const name = s.sheetName || `Sheet ${i + 1}`;
		const rows = getDataRows(s);
		const sheet = Workbook.addWorksheet(workbook, String(i + 1));

		addTextRow(sheet, `# ${name}`);
		for (let j = 0; j < s.sheetIntroText?.length || 0; j++) {
			addTextRow(
				sheet,
				`${s.sheetIntroText[j]}`,
				j === s.sheetIntroText.length - 1
					? { height: 40, alignment: { vertical: "top" } }
					: {}
			);
		}

		const tableRowNumber = Worksheet.rowCount(sheet) + 1;
		Table.add(sheet, {
			name: s.tableName || slugify(name),
			ref: `A${tableRowNumber}`,
			headerRow: true,
			columns: s.columns.map((c) => ({ name: c.heading, filterButton: false })),
			rows
		});
		Row.setFont(sheet, tableRowNumber, { ...defaultFont, bold: true });
		Row.setAlignment(sheet, tableRowNumber, { wrapText: true });

		for (let i = 0; i < s.columns.length; i++) {
			const colNum = i + 1;
			const colFormat = getColFormat(s.columns[i]);
			if (colFormat) {
				Column.setStyle(sheet, colNum, { numFmt: colFormat });
				Column.setWidth(sheet, colNum, getColWidth());
			} else {
				Column.setWidth(sheet, colNum, getColWidth(rows.map((d) => d[i])));
			}
		}
	}

	// Set workbook metadata and active sheet
	const model = Workbook.getModel(workbook);

	const creator = data.creator || "Anonymous";
	const created = data.created || new Date();

	model.title = isSingleSheet ? data.sheetName : data.coverSheetTitle || "";
	model.creator = creator;
	model.lastModifiedBy = creator;
	model.created = created;
	model.modified = created;
	model.defaultFont = defaultFont;
	model.views = [defaultView];

	Workbook.setModel(workbook, model);

	return Workbook.toBuffer(workbook);
}
