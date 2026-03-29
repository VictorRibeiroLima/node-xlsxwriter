// @ts-check
const { test } = require("node:test");
const assert = require("node:assert");
const { Workbook, Link, Format, Formula } = require("../src/index");
const fs = require("fs");
const findRootDir = require("./util");
const rootPath = findRootDir(__dirname);
const path = rootPath + "/temp";

test("should write merged cells to file", async (t) => {
	const workbook = new Workbook();
	const sheet = workbook.addSheet();

	//String value
	sheet.writeMergedString({
		firstCol: 1,
		firstRow: 1,
		lastCol: 3,
		lastRow: 3,
		value: "Merged Cells",
		format: new Format({
			align: "center",
			bold: true,
		}),
	});

	//Number value
	sheet.writeMergedNumber({
		firstCol: 4,
		firstRow: 1,
		lastCol: 6,
		lastRow: 3,
		value: 12345,
		format: new Format({
			align: "center",
			bold: true,
		}),
	});

	//Link value
	const link = new Link("http://example.com", "Example", "tooltip");
	sheet.writeMergedLink({
		firstCol: 7,
		firstRow: 1,
		lastCol: 9,
		lastRow: 3,
		value: link,
		format: new Format({
			align: "center",
			bold: true,
		}),
	});

	//Date value
	sheet.writeMergedDate({
		firstCol: 10,
		firstRow: 1,
		lastCol: 12,
		lastRow: 3,
		value: new Date("2024-01-01"),
		format: new Format({
			align: "center",
			bold: true,
			numFmt: "mm/dd/yyyy",
		}),
	});

	await workbook.saveToFile(`${path}/merged_cells.xlsx`);
	assert.ok(fs.existsSync(`${path}/merged_cells.xlsx`));
});

test("Should throw error when writing merged cells with invalid range", async (t) => {
	const workbook = new Workbook();
	const sheet = workbook.addSheet();

	const format = new Format({
		align: "center",
		bold: true,
	});

	assert.throws(() => {
		sheet.writeMergedString({
			firstCol: 1,
			firstRow: 1,
			lastCol: -1, // Invalid range
			lastRow: 3,
			value: "Invalid Range",
			format,
		});
	});
});

test("Should throw error if cell is only 1 cell", async (t) => {
	const workbook = new Workbook();
	const sheet = workbook.addSheet();

	const format = new Format({
		align: "center",
		bold: true,
	});

	assert.throws(() => {
		sheet.writeMergedString({
			firstCol: 1,
			firstRow: 1,
			lastCol: 1,
			lastRow: 1,
			value: "Invalid Range",
			format,
		});
	});
});

test("Should throw error if try to write merged cells in overlapping range", async (t) => {
	const workbook = new Workbook();
	const sheet = workbook.addSheet();

	const format = new Format({
		align: "center",
		bold: true,
	});

	sheet.writeMergedString({
		firstCol: 1,
		firstRow: 1,
		lastCol: 3,
		lastRow: 3,
		value: "Merged Cells",
		format,
	});

	sheet.writeMergedString({
		firstCol: 2, // Overlapping range
		firstRow: 2,
		lastCol: 4,
		lastRow: 4,
		value: "Overlapping Range",
		format,
	});

	assert.throws(() => {
		workbook.saveToFileSync(`${path}/overlapping_merged_cells.xlsx`);
	});
});
