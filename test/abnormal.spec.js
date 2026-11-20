import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  buildItemLine,
  buildListArticleBlocks,
  extractSubscriptions,
  normalizeArticleUrl,
  splitTitle,
} from "../src/article.js";
import { sheetValuesToObjects } from "../src/common.js";
import { normalizeItemRow_ } from "../src/admin.js";
import { subscriptionCells_ } from "../src/subscriptions.js";
import { crawlRoom } from "../src/room.js";

// Abnormal-case tests: the inputs from outside this code are sheet
// cell values (a cell can hold a string, a number, a boolean, a Date
// or nothing, whatever column it sits in), the admin API's JSON body,
// and the Rakuten ROOM API / home-server JSON. None of those are
// guaranteed to carry the type the shaping code expects, and a JSON
// payload can even carry an object whose toString is not callable
// ({"toString": []}), which makes String() throw. Every case below
// asserts graceful behaviour: no throw, and the deformed value reads
// as absent instead of leaking into a price, a count, a sheet row or
// the generated article.
const evil = JSON.parse('{"toString": []}');

describe("buildItemLine price cells", () => {
  const line = (price) => buildItemLine({ title: "x", price, unit: 1 });

  it("shows no price for a non-numeric price string", () => {
    // formatItemPrice documents "non-numeric cells produce no price
    // text"; a text cell used to come out as ¥0.
    expect(line("abc")).toBe("x");
    expect(line("未定")).toBe("x");
  });

  it("shows no price for boolean price cells", () => {
    expect(line(true)).toBe("x");
    expect(line(false)).toBe("x");
  });

  it("shows no price for an object price cell", () => {
    expect(line({})).toBe("x");
  });

  it("shows no price for a Date price cell", () => {
    // A date-typed cell used to have its digits concatenated into an
    // astronomical price.
    expect(line(new Date(2026, 10, 20))).toBe("x");
  });

  it("shows no price for an array price cell", () => {
    expect(line([500])).toBe("x");
  });

  it("shows no price and does not throw for an unstringifiable price", () => {
    expect(() => line(evil)).not.toThrow();
    expect(line(evil)).toBe("x");
  });

  it("shows no price for NaN or Infinity prices", () => {
    expect(line(NaN)).toBe("x");
    expect(line(Infinity)).toBe("x");
  });
});

describe("buildItemLine unit, title and comment cells", () => {
  it("does not turn an array unit into a count", () => {
    expect(buildItemLine({ title: "x", price: "", unit: [3] })).toBe("x");
  });

  it("does not throw on an unstringifiable unit", () => {
    expect(buildItemLine({ title: "x", price: "", unit: evil })).toBe("x");
  });

  it("keeps counting a unit stored as a string", () => {
    expect(buildItemLine({ title: "x", price: "", unit: "3" })).toBe(
      "x（×3）",
    );
  });

  it("does not throw on an unstringifiable title", () => {
    expect(buildItemLine({ title: evil, price: 1, unit: 1 })).toBe("（¥1）");
  });

  it("does not throw on an unstringifiable comment", () => {
    expect(
      buildItemLine({ title: "x", price: 1, unit: 1, comment: evil }),
    ).toBe("x（¥1）");
  });
});

describe("splitTitle with deformed titles", () => {
  it("reads an unstringifiable title as empty", () => {
    expect(splitTitle(evil)).toEqual({ name: "", props: "" });
  });

  it("reads a null title as empty", () => {
    expect(splitTitle(null)).toEqual({ name: "", props: "" });
  });
});

describe("normalizeArticleUrl with astral characters", () => {
  it("encodes an emoji in a search URL instead of throwing", () => {
    // The per-code-unit replace used to hand encodeURIComponent a
    // lone surrogate half, which throws URIError — one emoji in a
    // sheet URL killed the whole article regeneration.
    const url = "https://example.com/search?q=🎁 プレゼント";
    expect(normalizeArticleUrl(url)).toBe(
      "https://example.com/search?q=" +
        encodeURIComponent("🎁") +
        "%20" +
        encodeURIComponent("プレゼント"),
    );
  });

  it("encodes an astral character in the path", () => {
    expect(normalizeArticleUrl("https://example.com/𐮷野家")).toBe(
      "https://example.com/" +
        encodeURIComponent("𐮷") +
        encodeURIComponent("野") +
        encodeURIComponent("家"),
    );
  });
});

describe("buildListArticleBlocks with deformed items", () => {
  const empty = { room: [], workman: [], variety: [], others: [] };

  it("groups an item with an unstringifiable genre under その他", () => {
    const blocks = buildListArticleBlocks(
      {
        ...empty,
        room: [{ title: "x", genre: evil, url: "", price: 1, unit: 1 }],
      },
      [],
    );
    const headings = blocks.filter((b) => b.type === "heading2").map((b) => b.text);
    expect(headings).toEqual(["その他", "サブスクリプション"]);
  });

  it("drops an item whose title is unstringifiable instead of throwing", () => {
    const blocks = buildListArticleBlocks(
      {
        ...empty,
        room: [{ title: evil, genre: "g", url: "", price: 1, unit: 1 }],
      },
      [],
    );
    expect(blocks.filter((b) => b.type === "itemTitle")).toEqual([]);
  });

  it("keeps block text a string for an unstringifiable subscription name", () => {
    const blocks = buildListArticleBlocks(empty, [
      { name: evil, purpose: "p", url: "", flag: "" },
    ]);
    const titles = blocks.filter((b) => b.type === "itemTitle");
    expect(titles).toHaveLength(1);
    expect(titles[0].text).toBe("");
  });
});

describe("extractSubscriptions with deformed rows", () => {
  const header = ["Service", "Purpose", "URL", "flag"];

  it("drops a row whose service name is unstringifiable", () => {
    expect(extractSubscriptions([header, [evil, "p", "u", ""]])).toEqual([]);
  });

  it("stringifies a numeric service name", () => {
    expect(extractSubscriptions([header, [123, "p", "u", ""]])).toEqual([
      { name: "123", purpose: "p", url: "u", flag: "" },
    ]);
  });
});

describe("sheetValuesToObjects with deformed sheets", () => {
  it("reads an empty sheet as no objects", () => {
    expect(sheetValuesToObjects([])).toEqual([]);
  });

  it("reads a header-only sheet as no objects", () => {
    expect(sheetValuesToObjects([["title", "price"]])).toEqual([]);
  });

  it("keeps short rows, with the missing cells undefined", () => {
    const objects = sheetValuesToObjects([["title", "price"], ["x"]]);
    expect(objects).toHaveLength(1);
    expect(objects[0].title).toBe("x");
    expect(Object.prototype.hasOwnProperty.call(objects[0], "price")).toBe(
      true,
    );
    expect(objects[0].price).toBeUndefined();
  });
});

describe("normalizeItemRow_ with deformed payloads", () => {
  it("does not store a boolean price as 1", () => {
    expect(normalizeItemRow_({ title: "t", price: true, unit: 1 })[3]).toBe("");
  });

  it("does not store an array price as its element", () => {
    expect(normalizeItemRow_({ title: "t", price: [500], unit: 1 })[3]).toBe(
      "",
    );
  });

  it("reads an unstringifiable price as blank instead of throwing", () => {
    expect(normalizeItemRow_({ title: "t", price: evil, unit: 1 })[3]).toBe(
      "",
    );
  });

  it("reads an unstringifiable title as blank instead of throwing", () => {
    expect(normalizeItemRow_({ title: evil, price: 1, unit: 1 })[0]).toBe("");
  });

  it("does not store a boolean unit as 1", () => {
    expect(normalizeItemRow_({ title: "t", price: 1, unit: true })[4]).toBe(
      "",
    );
  });

  it("reads an unstringifiable flag as blank instead of throwing", () => {
    expect(
      normalizeItemRow_({ title: "t", price: 1, unit: 1, flag: evil })[7],
    ).toBe("");
  });
});

describe("subscriptionCells_ with deformed payloads", () => {
  it("does not store a boolean amount as 1", () => {
    const cells = subscriptionCells_(
      { service: "s", amount: true, basis: "yearly" },
      2,
      "",
    );
    expect(cells[2]).toBe("");
    expect(cells[3]).toBe("");
  });

  it("does not store an array amount as its element", () => {
    const cells = subscriptionCells_(
      { service: "s", amount: [500], basis: "monthly" },
      3,
      "",
    );
    expect(cells[2]).toBe("");
    expect(cells[3]).toBe("");
  });

  it("reads an unstringifiable amount as blank instead of throwing", () => {
    const cells = subscriptionCells_(
      { service: "s", amount: evil, basis: "yearly" },
      2,
      "",
    );
    expect(cells[2]).toBe("");
    expect(cells[3]).toBe("");
  });

  it("reads an unstringifiable service as blank instead of throwing", () => {
    const cells = subscriptionCells_(
      { service: evil, amount: 100, basis: "yearly" },
      2,
      "",
    );
    expect(cells[0]).toBe("");
    expect(cells[2]).toBe(100);
  });

  it("keeps parsing a string amount on the monthly basis", () => {
    const cells = subscriptionCells_(
      { service: "s", amount: "6000", basis: "monthly" },
      5,
      "",
    );
    expect(cells[2]).toBe("=D5*12");
    expect(cells[3]).toBe(6000);
  });
});

// --- crawlRoom with deformed ROOM payloads (own mock harness) ---
const mockRange = {
  clearContent: vi.fn(),
  setValues: vi.fn(),
  getValues: vi.fn(() => []),
};
const mockSheet = {
  getRange: vi.fn().mockReturnValue(mockRange),
  getLastRow: vi.fn().mockReturnValue(1),
  getDataRange: vi.fn().mockReturnValue(mockRange),
};
const mockSpreadsheet = {
  getSheetByName: vi.fn().mockReturnValue(mockSheet),
};
const mockUrlFetchApp = { fetch: vi.fn() };
global.SpreadsheetApp = { openById: vi.fn().mockReturnValue(mockSpreadsheet) };
global.UrlFetchApp = mockUrlFetchApp;
global.Logger = { log: vi.fn() };
global.PropertiesService = {
  getScriptProperties: vi.fn(() => ({
    getProperty: vi.fn(
      (name) =>
        ({ DATA_BASE_URL: "http://home.example", ROOM_USER_ID: "u1" })[name] ||
        null,
    ),
  })),
};
global.CacheService = {
  getScriptCache: vi.fn(() => ({ get: vi.fn(() => null), removeAll: vi.fn() })),
};

const json = (obj) => ({ getContentText: () => JSON.stringify(obj) });
const raw = (text) => ({ getContentText: () => text });
const collect = (name, price = 100) => ({
  item: { name, url: `http://example.com/${name}`, price },
  content: "",
});

describe("crawlRoom with deformed ROOM payloads", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockSheet.getLastRow.mockReturnValue(1);
  });

  it("leaves the sheet untouched when the fallback payload is not an array", () => {
    mockSheet.getLastRow.mockReturnValue(5);
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(raw("403 forbidden"))
      .mockReturnValueOnce(json({ data: { nope: 1 } }));
    expect(() => crawlRoom()).toThrow();
    expect(mockRange.clearContent).not.toHaveBeenCalled();
    expect(mockRange.setValues).not.toHaveBeenCalled();
  });

  it("skips null entries inside the fallback collects", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(raw("403 forbidden"))
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(
        json({ data: [collect("a"), null, collect("b")] }),
      );
    expect(() => crawlRoom()).not.toThrow();
    expect(mockRange.setValues).toHaveBeenCalledTimes(1);
    expect(mockRange.setValues.mock.calls[0][0]).toHaveLength(2);
  });

  it("skips null and non-object entries inside the fallback collections", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(raw("403 forbidden"))
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }, null, "junk"] }))
      .mockReturnValueOnce(json({ data: [collect("a")] }));
    expect(() => crawlRoom()).not.toThrow();
    expect(mockUrlFetchApp.fetch).toHaveBeenCalledTimes(3);
    expect(mockRange.setValues.mock.calls[0][0]).toHaveLength(1);
  });

  it("skips non-object collect entries instead of writing blank rows", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(
        json({ data: [collect("a"), "junk", 42, collect("b")] }),
      );
    expect(() => crawlRoom()).not.toThrow();
    expect(mockRange.setValues.mock.calls[0][0]).toHaveLength(2);
  });

  it("uses the API directly when its collects carry a null entry", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(json({ data: [collect("a"), null] }));
    expect(() => crawlRoom()).not.toThrow();
    expect(mockUrlFetchApp.fetch).toHaveBeenCalledTimes(2);
    expect(mockRange.setValues.mock.calls[0][0]).toHaveLength(1);
  });

  it("falls back when the API collects payload is not an array", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(json({ data: "oops" }))
      .mockReturnValueOnce(json({ data: [{ id: 2, name: "c2" }] }))
      .mockReturnValueOnce(json({ data: [collect("a")] }));
    expect(() => crawlRoom()).not.toThrow();
    expect(mockRange.setValues).toHaveBeenCalledTimes(1);
  });

  it("writes an empty name cell for an unstringifiable item name", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(
        json({
          data: [
            { item: { name: evil, url: "http://x", price: 100 }, content: "" },
          ],
        }),
      );
    expect(() => crawlRoom()).not.toThrow();
    expect(mockRange.setValues.mock.calls[0][0][0][0]).toBe("");
  });

  it("writes an empty price cell for deformed prices", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(
        json({
          data: [
            { item: { name: "a", url: "http://x", price: evil }, content: "" },
            { item: { name: "b", url: "http://y", price: [500] }, content: "" },
            { item: { name: "c", url: "http://z", price: true }, content: "" },
          ],
        }),
      );
    expect(() => crawlRoom()).not.toThrow();
    const rows = mockRange.setValues.mock.calls[0][0];
    expect(rows.map((r) => r[3])).toEqual(["", "", ""]);
  });

  it("writes an empty genre cell for an unstringifiable collection name", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: evil }] }))
      .mockReturnValueOnce(json({ data: [collect("a")] }));
    expect(() => crawlRoom()).not.toThrow();
    expect(mockRange.setValues.mock.calls[0][0][0][5]).toBe("");
  });

  it("writes an empty comment cell for unstringifiable content", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(
        json({
          data: [
            { item: { name: "a", url: "http://x", price: 1 }, content: evil },
          ],
        }),
      );
    expect(() => crawlRoom()).not.toThrow();
    expect(mockRange.setValues.mock.calls[0][0][0][6]).toBe("");
  });

  it("stores a numeric picture URL as text", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce(json({ data: [{ id: 1, name: "c1" }] }))
      .mockReturnValueOnce(
        json({
          data: [
            {
              item: {
                name: "a",
                picture: { url: 123 },
                url: "http://x",
                price: 1,
              },
              content: "",
            },
          ],
        }),
      );
    expect(() => crawlRoom()).not.toThrow();
    expect(mockRange.setValues.mock.calls[0][0][0][1]).toBe("123");
  });
});
