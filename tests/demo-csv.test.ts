import assert from "node:assert/strict";
import test from "node:test";
import {
  CsvParseError,
  csvValueToCellInput,
  normalizeHeader,
  parseCsv
} from "../website/app/(home)/tools/update-excel-from-csv/csv.ts";

test("demo CSV parser handles quoted commas, escaped quotes, and embedded newlines", () => {
  const parsed = parseCsv(
    'Name,Note,Amount\r\n"North, Inc.","Said ""hello""",42\r\nSouth,"Line one\nLine two",17\r\n'
  );

  assert.equal(parsed.delimiter, ",");
  assert.deepEqual(parsed.headers, ["Name", "Note", "Amount"]);
  assert.deepEqual(parsed.rows, [
    ["North, Inc.", 'Said "hello"', "42"],
    ["South", "Line one\nLine two", "17"]
  ]);
});

test("demo CSV parser detects tabs and reports duplicate normalized headers", () => {
  const parsed = parseCsv("Customer ID\tcustomer_id\tAmount\n1\t2\t3");

  assert.equal(parsed.delimiter, "\t");
  assert.deepEqual(parsed.warnings, ["Duplicate CSV headers found: customer_id"]);
  assert.equal(normalizeHeader(" Customer-ID "), "customerid");
});

test("demo CSV values preserve leading zeros and formula-looking strings", () => {
  assert.equal(csvValueToCellInput("42.5"), 42.5);
  assert.equal(csvValueToCellInput("TRUE"), true);
  assert.equal(csvValueToCellInput("00123"), "00123");
  assert.equal(csvValueToCellInput("=2+2"), "=2+2");
  assert.equal(csvValueToCellInput(""), null);
});

test("demo CSV parser rejects empty headers and unclosed quoted fields", () => {
  assert.throws(() => parseCsv("Name,,Amount\nNorth,One,2"), CsvParseError);
  assert.throws(() => parseCsv('Name,Note\nNorth,"unfinished'), /unclosed quoted field/);
});
