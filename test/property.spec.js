import { describe, expect, it} from "vitest";
import fc from "fast-check";
import {
buildItemLine,
buildListArticleBlocks,
normalizeArticleUrl,
splitTitle,
} from "../src/article.js";
import { sheetValuesToObjects} from "../src/common.js";
import { normalizeItemRow_} from "../src/admin.js";
import { subscriptionCells_} from "../src/subscriptions.js";

// Property-based tests (fast-check) over the same outside inputs as
// the abnormal-case suite: sheet cell values, JSON payloads, and the
// ROOM/article shaping built on them. The JSON domain includes an
// object whose toString is not callable ({"toString": []}), which
// JSON.parse can produce and String() chokes on.
const evil = JSON.parse('{"toString": []}');

const jsonish = fc.oneof(
fc.string(),
fc.double(),
fc.boolean(),
fc.constant(null),
fc.array(
fc.oneof(fc.string(), fc.double(), fc.boolean(), fc.constant(null)),
{ maxLength: 3},
),
fc.dictionary(fc.string(), fc.string()),
fc.constant(evil),
);

// The documented text coercion, restated as the reference for
// counting: null/absent and unstringifiable values read as empty.
const textOf = (v) => {
if (v === null || v === undefined) return "";
try {
return String(v);
} catch {
return "";
}
};

describe("splitTitle properties", () => {
// Tokens covering the name/property boundary: bare numbers stay in
// the name, numbers with units and colours are properties, and
// '&' / '＆' / '送料無料' are documented drops.
const token = fc.constantFrom(
"Anker",
"ケーブル",
"USB",
"543",
"1.8m",
"ブラック",
"ホワイト",
"特大",
"対応",
"3色ボールペン",
"&",
"＆",
"送料無料",
);
const plainTitle = fc
.array(token, { minLength: 1, maxLength: 8})
.map((tokens) => tokens.join(" "));

it("conserves the non-dropped tokens between name and props", () => {
fc.assert(
fc.property(plainTitle, (title) => {
const { name, props} = splitTitle(title);
const kept = title
.split(" ")
.filter((t) => t!== "&" && t!== "＆" && t!== "送料無料")
.sort();
const out = `${name} ${props}`
.split(" ")
.filter((t) => t!== "")
.sort();
expect(out).toEqual(kept);
}),
{ numRuns: 100},
);
});

it("is idempotent on the name it returns", () => {
fc.assert(
fc.property(plainTitle, (title) => {
const { name} = splitTitle(title);
expect(splitTitle(name).name).toBe(name);
}),
{ numRuns: 100},
);
});

it("never throws and keeps name tokens traceable with brackets", () => {
const bracketed = fc
.array(
fc.oneof(
token,
fc
.constantFrom("送料無料", "Office 2019", "新品", "50%オフ")
.map((c) => ``),
),
{ minLength: 1, maxLength: 6},
)
.map((tokens) => tokens.join(" "));
fc.assert(
fc.property(bracketed, (title) => {
const { name, props} = splitTitle(title);
expect(typeof name).toBe("string");
expect(typeof props).toBe("string");
const allowed = new Set(
title
.replace(/[]/g, " ")
.split(/[\s　]+/)
.filter(Boolean),
);
`${name} ${props}`
.split(" ")
.filter(Boolean)
.forEach((t) => expect(allowed.has(t)).toBe(true));
}),
{ numRuns: 100},
);
});

it("never throws on arbitrary text", () => {
fc.assert(
fc.property(fc.string({ unit: "grapheme"}), (title) => {
const { name, props} = splitTitle(title);
expect(typeof name).toBe("string");
expect(typeof props).toBe("string");
}),
{ numRuns: 100},
);
});
});

describe("buildItemLine price properties", () => {
const line = (price) => buildItemLine({ title: "x", price, unit: 1});

it("never throws, and any price shown has the ¥ shape", () => {
fc.assert(
fc.property(jsonish, (price) => {
const out = line(price);
expect(out === "x" || /^x（¥-?[\d,]+）$/.test(out)).toBe(true);
}),
{ numRuns: 100},
);
});

it("shows no price for a price that is neither a number nor a string", () => {
const notNumberOrString = fc.oneof(
fc.boolean(),
fc.constant(null),
fc.array(fc.double(), { maxLength: 3}),
fc.dictionary(fc.string(), fc.string()),
fc.constant(evil),
);
fc.assert(
fc.property(notNumberOrString, (price) => {
expect(line(price)).toBe("x");
}),
{ numRuns: 100},
);
});

it("formats a non-negative integer price as its grouped ¥ value", () => {
fc.assert(
fc.property(fc.integer({ min: 0, max: 999999999}), (price) => {
expect(line(price)).toBe(`x（¥${price.toLocaleString("en-US")}）`);
}),
{ numRuns: 100},
);
});
});

describe("buildItemLine totality over the JSON domain", () => {
it("always returns a clean string for deformed fields", () => {
const safeText = fc.stringMatching(/^[a-zA-Zぁ-ん]{0,12}$/);
fc.assert(
fc.property(
fc.record({
title: safeText,
comment: safeText,
price: jsonish,
unit: jsonish,
}),
(item) => {
const out = buildItemLine(item);
expect(typeof out).toBe("string");
expect(out).not.toContain("NaN");
expect(out).not.toContain("undefined");
expect(out).not.toContain("[object Object]");
expect(out).not.toContain("Infinity");
},
),
{ numRuns: 100},
);
});
});

describe("normalizeArticleUrl properties", () => {
// Domain: well-formed text (grapheme units). Sheet text cannot
// carry a lone surrogate, and encodeURIComponent cannot encode one
// by design, so lone surrogates are outside the domain.
it("never throws, returns pure ASCII without spaces, and is idempotent", () => {
fc.assert(
fc.property(fc.string({ unit: "grapheme"}), (url) => {
const out = normalizeArticleUrl(url);
expect(/^[\x00-\x7F]*$/.test(out)).toBe(true);
expect(out).not.toContain(" ");
expect(normalizeArticleUrl(out)).toBe(out);
}),
{ numRuns: 100},
);
});
});

describe("sheetValuesToObjects properties", () => {
// Header names are the sheets' own fixed column names; a column
// literally named "__proto__" is outside the domain, because a
// plain-object row cannot hold it (assignment would set the
// prototype instead of an own property).
const headers = fc.uniqueArray(
fc.stringMatching(/^[a-z][a-z0-9_]{0,7}$/),
{ minLength: 1, maxLength: 6},
);
const cell = fc.oneof(
fc.string(),
fc.double(),
fc.boolean(),
fc.constant(null),
);

it("conserves rows and maps every header onto its cell", () => {
fc.assert(
fc.property(
headers,
fc.array(fc.array(cell, { maxLength: 8}), { maxLength: 6}),
(headerRow, rows) => {
const objects = sheetValuesToObjects(
[headerRow,...rows].map((r) => r.slice()),
);
expect(objects).toHaveLength(rows.length);
objects.forEach((obj, i) => {
headerRow.forEach((header, j) => {
expect(
Object.prototype.hasOwnProperty.call(obj, header),
).toBe(true);
expect(obj[header]).toBe(
j < rows[i].length? rows[i][j]: undefined,
);
});
});
},
),
{ numRuns: 100},
);
});
});

describe("buildListArticleBlocks properties", () => {
const field = fc.oneof(fc.string({ unit: "grapheme"}), jsonish);
const item = fc.record({
title: field,
genre: field,
url: field,
price: field,
unit: field,
comment: field,
flag: field,
});
const itemsByTab = fc.record({
room: fc.array(item, { maxLength: 4}),
workman: fc.array(item, { maxLength: 4}),
variety: fc.array(item, { maxLength: 4}),
others: fc.array(item, { maxLength: 4}),
});
const subscriptions = fc.array(
fc.record({
name: fc.stringMatching(/^[a-zA-Z]{1,10}$/),
purpose: fc.string(),
url: fc.string(),
flag: fc.string(),
}),
{ maxLength: 3},
);

it("never throws, and every block carries string text", () => {
fc.assert(
fc.property(itemsByTab, subscriptions, (byTab, subs) => {
const blocks = buildListArticleBlocks(byTab, subs);
// 'toc' blocks carry no text by design: the Doc writer substitutes
// the [[目次]] marker for them. Every other block's text must be a
// string.
blocks.forEach((b) => {
if (b.type !== "toc") expect(typeof b.text).toBe("string");
});
}),
{ numRuns: 100},
);
});

it("numbers the titled items consecutively and keeps links ASCII", () => {
fc.assert(
fc.property(itemsByTab, subscriptions, (byTab, subs) => {
const blocks = buildListArticleBlocks(byTab, subs);
const numbered = blocks
.filter((b) => b.type === "itemTitle" && /^\d+\. /.test(b.text))
.map((b) => parseInt(b.text, 10));
const expectedCount = [
...byTab.room,
...byTab.workman,
...byTab.variety,
...byTab.others,
].filter((it) => textOf(it.title).trim()!== "").length;
expect(numbered).toEqual(
Array.from({ length: expectedCount}, (_, i) => i + 1),
);
blocks
.filter((b) => b.type === "link")
.forEach((b) => {
expect(/^[\x00-\x7F]*$/.test(b.text)).toBe(true);
expect(b.text).not.toContain(" ");
});
}),
{ numRuns: 100},
);
});
});

describe("normalizeItemRow_ properties", () => {
it("always returns eight scalar cells, with numeric cells finite or blank", () => {
const itemArb = fc.record({
title: jsonish,
image: jsonish,
url: jsonish,
price: jsonish,
unit: jsonish,
genre: jsonish,
comment: jsonish,
flag: jsonish,
});
fc.assert(
fc.property(itemArb, (item) => {
const row = normalizeItemRow_(item);
expect(row).toHaveLength(8);
row.forEach((cellValue) => {
expect(
typeof cellValue === "string" || typeof cellValue === "number",
).toBe(true);
});
[3, 4].forEach((i) => {
expect(
row[i] === "" ||
(typeof row[i] === "number" && Number.isFinite(row[i])),
).toBe(true);
});
}),
{ numRuns: 100},
);
});
});

describe("subscriptionCells_ properties", () => {
it("always returns seven scalar cells, with cost cells blank, numeric or the deriving formula", () => {
const subArb = fc.record({
service: jsonish,
purpose: jsonish,
url: jsonish,
flag: jsonish,
amount: jsonish,
basis: fc.oneof(
fc.constant("yearly"),
fc.constant("monthly"),
fc.string(),
fc.constant(null),
),
});
fc.assert(
fc.property(subArb, fc.integer({ min: 2, max: 999}), (sub, row) => {
const cells = subscriptionCells_(sub, row, "");
expect(cells).toHaveLength(7);
cells.forEach((cellValue) => {
expect(
typeof cellValue === "string" || typeof cellValue === "number",
).toBe(true);
});
[2, 3].forEach((i) => {
const v = cells[i];
expect(
v === "" ||
(typeof v === "number" && Number.isFinite(v)) ||
/^=C\d+\/12$/.test(v) ||
/^=D\d+\*12$/.test(v),
).toBe(true);
});
}),
{ numRuns: 100},
);
});
});
