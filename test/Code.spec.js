import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { doGet } from "../src/main.js";
import { crawlRoom } from "../src/room.js";

// Mockup of Google Apps Script global objects
const mockRange = {
  clearContent: vi.fn(),
  setValues: vi.fn(),
  getValues: vi.fn().mockImplementation(() => [
    ["header1", "header2"],
    ["value1", "value2"],
  ]),
};

const mockSheet = {
  getRange: vi.fn().mockReturnValue(mockRange),
  getLastRow: vi.fn().mockReturnValue(1),
  getDataRange: vi.fn().mockReturnValue(mockRange),
};

const mockSpreadsheet = {
  getSheetByName: vi.fn().mockReturnValue(mockSheet),
};

const mockContentService = {
  createTextOutput: vi.fn().mockReturnThis(),
  setMimeType: vi.fn().mockReturnThis(),
  setContent: vi.fn(),
  MimeType: {
    JSON: "JSON",
    JAVASCRIPT: "JAVASCRIPT",
  },
};

const mockUrlFetchApp = {
  fetch: vi.fn().mockReturnValue({
    getContentText: vi.fn(),
  }),
};

// Assign mocks to global scope
global.SpreadsheetApp = {
  openById: vi.fn().mockReturnValue(mockSpreadsheet),
};
global.ContentService = mockContentService;
global.UrlFetchApp = mockUrlFetchApp;
global.Logger = {
  log: vi.fn(),
};
global.PropertiesService = {
  getScriptProperties: vi.fn(() => ({
    getProperty: vi.fn(
      (name) =>
        ({
          DATA_BASE_URL: "http://home-api.example",
          ROOM_USER_ID: "room-user-1",
        })[name] || null
    ),
  })),
};

// In-memory stand-in for the script cache used by the doGet response cache
const createMockCache = () => {
  const store = new Map();
  return {
    store,
    get: vi.fn((key) => (store.has(key) ? store.get(key) : null)),
    getAll: vi.fn((keys) => {
      const values = {};
      keys.forEach((key) => {
        if (store.has(key)) {
          values[key] = store.get(key);
        }
      });
      return values;
    }),
    putAll: vi.fn((entries) => {
      Object.entries(entries).forEach(([key, value]) => store.set(key, value));
    }),
    removeAll: vi.fn((keys) => {
      keys.forEach((key) => store.delete(key));
    }),
  };
};

let mockCache;
global.CacheService = {
  getScriptCache: vi.fn(() => mockCache),
};

beforeEach(() => {
  mockCache = createMockCache();
});

describe("doGet", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockSpreadsheet.getSheetByName.mockImplementation(sheetName => {
      if (sheetName === "room") {
        return {
          ...mockSheet,
          getDataRange: () => ({
            getValues: () => [
              ["header1", "header2"],
              ["value1", "value2"],
            ],
          }),
        };
      }
      return {
        ...mockSheet,
        getDataRange: () => ({
          getValues: () => [["header1", "header2"]], // Empty data for other sheets
        }),
      };
    });
  });

  it("should return JSON when no callback is provided", () => {
    const e = { parameter: {} };
    doGet(e);
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JSON");
    expect(mockContentService.setContent).toHaveBeenCalledWith(
      JSON.stringify([{ header1: "value1", header2: "value2", shortTitle: "", props: "" }])
    );
  });

  it("should return JSONP when a callback is provided", () => {
    const e = { parameter: { callback: "myCallback" } };
    doGet(e);
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JAVASCRIPT");
    expect(mockContentService.setContent).toHaveBeenCalledWith(
      'myCallback&&myCallback(' +
        JSON.stringify([{ header1: "value1", header2: "value2", shortTitle: "", props: "" }]) +
        ');'
    );
  });

  it("should fall back to JSON for a callback that is not a plain identifier", () => {
    const e = { parameter: { callback: "alert(1)//" } };
    doGet(e);
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JSON");
    expect(mockContentService.setContent).toHaveBeenCalledWith(
      JSON.stringify([{ header1: "value1", header2: "value2", shortTitle: "", props: "" }])
    );
  });
});

describe("doGet response cache", () => {
  const META_KEY = "items-response:v2:meta";
  const CHUNK_KEY = (index) => `items-response:v2:${index}`;

  const stubSheetValues = (roomValues) => {
    mockSpreadsheet.getSheetByName.mockImplementation((sheetName) => ({
      ...mockSheet,
      getDataRange: () => ({
        getValues: () => (sheetName === "room" ? roomValues : [["h"]]),
      }),
    }));
  };

  beforeEach(() => {
    vi.clearAllMocks();
    stubSheetValues([["h"], ["v"]]);
  });

  it("serves the second request from the cache without reopening the spreadsheet", () => {
    const e = { parameter: {} };
    doGet(e);
    doGet(e);
    expect(global.SpreadsheetApp.openById).toHaveBeenCalledTimes(1);
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(
      JSON.stringify([{ h: "v", shortTitle: "", props: "" }])
    );
  });

  it("wraps cached data as JSONP when a callback is provided", () => {
    doGet({ parameter: {} });
    doGet({ parameter: { callback: "cb" } });
    expect(global.SpreadsheetApp.openById).toHaveBeenCalledTimes(1);
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(
      "cb&&cb(" + JSON.stringify([{ h: "v", shortTitle: "", props: "" }]) + ");"
    );
  });

  it("splits a large response into multiple cache chunks", () => {
    const big = "x".repeat(45000);
    stubSheetValues([["h"], [big]]);
    doGet({ parameter: {} });

    const json = JSON.stringify([{ h: big, shortTitle: "", props: "" }]);
    const meta = JSON.parse(mockCache.store.get(META_KEY));
    expect(meta).toEqual({ chunks: 3, length: json.length });
    const parts = [0, 1, 2].map((i) => mockCache.store.get(CHUNK_KEY(i)));
    parts.forEach((part) => expect(part.length).toBeLessThanOrEqual(20000));
    expect(parts.join("")).toBe(json);
  });

  it("does not split a surrogate pair across a chunk boundary", () => {
    // The JSON prefix `[{"h":"` is 7 chars, so 19992 filler chars place
    // the emoji's high surrogate exactly at the end of the first chunk.
    const value = "a".repeat(19992) + "😀" + "b".repeat(10);
    stubSheetValues([["h"], [value]]);
    doGet({ parameter: {} });

    const json = JSON.stringify([{ h: value, shortTitle: "", props: "" }]);
    const meta = JSON.parse(mockCache.store.get(META_KEY));
    expect(meta.chunks).toBe(2);
    expect(mockCache.store.get(CHUNK_KEY(0)).length).toBe(20001);
    const parts = [0, 1].map((i) => mockCache.store.get(CHUNK_KEY(i)));
    expect(parts.join("")).toBe(json);

    doGet({ parameter: {} });
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(json);
  });

  it("rebuilds from the sheets when a cache chunk is missing", () => {
    doGet({ parameter: {} });
    mockCache.store.delete(CHUNK_KEY(0));
    stubSheetValues([["h"], ["changed"]]);

    doGet({ parameter: {} });
    expect(global.SpreadsheetApp.openById).toHaveBeenCalledTimes(2);
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(
      JSON.stringify([{ h: "changed", shortTitle: "", props: "" }])
    );
  });

  it("rebuilds from the sheets when the cache manifest is corrupt", () => {
    mockCache.store.set(META_KEY, "not-json");
    doGet({ parameter: {} });
    expect(global.SpreadsheetApp.openById).toHaveBeenCalledTimes(1);
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(
      JSON.stringify([{ h: "v", shortTitle: "", props: "" }])
    );
  });

  it("rebuilds from the sheets when the cached length does not match", () => {
    doGet({ parameter: {} });
    const meta = JSON.parse(mockCache.store.get(META_KEY));
    mockCache.store.set(
      META_KEY,
      JSON.stringify({ chunks: meta.chunks, length: meta.length + 1 })
    );
    stubSheetValues([["h"], ["changed"]]);

    doGet({ parameter: {} });
    expect(global.SpreadsheetApp.openById).toHaveBeenCalledTimes(2);
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(
      JSON.stringify([{ h: "changed", shortTitle: "", props: "" }])
    );
  });

  it("still serves the response when reading the cache throws", () => {
    mockCache.get.mockImplementationOnce(() => {
      throw new Error("cache read failed");
    });
    doGet({ parameter: {} });
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(
      JSON.stringify([{ h: "v", shortTitle: "", props: "" }])
    );
  });

  it("still serves the response when writing the cache throws", () => {
    mockCache.putAll.mockImplementationOnce(() => {
      throw new Error("cache write failed");
    });
    doGet({ parameter: {} });
    expect(mockContentService.setContent).toHaveBeenLastCalledWith(
      JSON.stringify([{ h: "v", shortTitle: "", props: "" }])
    );
  });
});

describe("crawlRoom", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockSheet.getLastRow.mockReturnValue(1);
    mockSpreadsheet.getSheetByName.mockReturnValue(mockSheet);
  });

  it("should fetch data and write to the sheet", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({
            data: [
              {
                item: {
                  key: "test-item;123",
                  name: "Test Item",
                  picture: { url: "http://example.com/test.jpg" },
                  url: "http://example.com/item",
                  price: 1000,
                },
                content: "Test comment",
              },
            ],
          }),
      });

    crawlRoom();
    const expectedRow = [
      "Test Item",
      "http://example.com/test.jpg",
      "http://example.com/item",
      1000,
      1,
      "collection1",
      "Test comment",
    ];
    expect(mockRange.setValues).toHaveBeenCalledWith([expectedRow]);
    expect(Logger.log).toHaveBeenCalledWith("collection1");
    expect(Logger.log).toHaveBeenCalledWith("Test Item");
    expect(Logger.log).toHaveBeenCalledWith("Fetched 1 items.");
  });

  it("should store Rakuten picture URLs with a resize query", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({
            data: [
              {
                item: {
                  name: "Rakuten Item",
                  picture: { url: "https://tshop.r10s.jp/shop/cabinet/item.jpg" },
                  url: "https://item.rakuten.co.jp/shop/item/",
                  price: 2000,
                },
                content: "comment",
              },
              {
                item: {
                  name: "Already Queried",
                  picture: { url: "https://shop.r10s.jp/shop/cabinet/item2.jpg?_ex=100x100" },
                  url: "https://item.rakuten.co.jp/shop/item2/",
                  price: 3000,
                },
                content: "comment2",
              },
            ],
          }),
      });

    crawlRoom();

    expect(mockRange.setValues).toHaveBeenCalledWith([
      [
        "Rakuten Item",
        "https://tshop.r10s.jp/shop/cabinet/item.jpg?_ex=400x400",
        "https://item.rakuten.co.jp/shop/item/",
        2000,
        1,
        "collection1",
        "comment",
      ],
      [
        "Already Queried",
        "https://shop.r10s.jp/shop/cabinet/item2.jpg?_ex=100x100",
        "https://item.rakuten.co.jp/shop/item2/",
        3000,
        1,
        "collection1",
        "comment2",
      ],
    ]);
  });

  it("should clear the doGet response cache after rewriting the room sheet", () => {
    mockCache.store.set(
      "items-response:v2:meta",
      JSON.stringify({ chunks: 1, length: 2 })
    );
    mockCache.store.set("items-response:v2:0", "[]");
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [] }),
      });

    crawlRoom();

    expect(mockCache.store.has("items-response:v2:meta")).toBe(false);
    expect(mockCache.store.has("items-response:v2:0")).toBe(false);
  });

  it("should not fail when clearing the response cache throws", () => {
    mockCache.removeAll.mockImplementationOnce(() => {
      throw new Error("cache remove failed");
    });
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [] }),
      });

    expect(() => crawlRoom()).not.toThrow();
  });

  it("should handle empty collections", () => {
    // Empty collections from the ROOM API count as a direct-fetch
    // failure, so the fallback is asked too; it is empty as well.
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [] }),
      })
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [] }),
      });

    crawlRoom();

    expect(mockRange.clearContent).not.toHaveBeenCalled();
    expect(mockRange.setValues).not.toHaveBeenCalled();
  });

  it("should handle empty collects", () => {
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () =>
        JSON.stringify({
          data: [{ id: 1, name: "collection1" }],
        }),
    });
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () => JSON.stringify({ data: [] }),
    });

    crawlRoom();

    expect(mockRange.clearContent).not.toHaveBeenCalled();
    expect(mockRange.setValues).not.toHaveBeenCalled();
  });

  it("should handle items with no key", () => {
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () =>
        JSON.stringify({
          data: [{ id: 1, name: "collection1" }],
        }),
    });
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () =>
        JSON.stringify({
          data: [
            {
              item: {
                name: "Test Item",
                picture: { url: "http://example.com/test.jpg" },
                url: "http://example.com/item",
                price: 1000,
              },
              content: "Test comment",
            },
          ],
        }),
    });

    crawlRoom();

    expect(mockRange.clearContent).not.toHaveBeenCalled();
    expect(mockRange.setValues).toHaveBeenCalledWith([
      [
        "Test Item",
        "http://example.com/test.jpg",
        "http://example.com/item",
        1000,
        1,
        "collection1",
        "Test comment",
      ],
    ]);
  });

  it("should handle collections API with no data property", () => {
    // No data property means no collections, which fails the direct
    // fetch; the fallback answers the same way.
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ message: "error" }),
      })
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ message: "error" }),
      });
    crawlRoom();
    expect(mockRange.setValues).not.toHaveBeenCalled();
  });

  it("should handle collects API with no data property", () => {
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () => JSON.stringify({ data: [{ id: 1, name: "c1" }] }),
    });
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () => JSON.stringify({ message: "error" }),
    });
    crawlRoom();
    expect(mockRange.setValues).not.toHaveBeenCalled();
  });

  it("should handle collect with no item property", () => {
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () => JSON.stringify({ data: [{ id: 1, name: "c1" }] }),
    });
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () => JSON.stringify({ data: [{ content: "comment" }] }),
    });
    crawlRoom();
    expect(mockRange.setValues).toHaveBeenCalledWith([
      ["", "", "", "", 1, "c1", "comment"],
    ]);
  });

  it("should handle missing properties in API responses", () => {
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () => JSON.stringify({ data: [{ id: 1 /*no name*/ }] }),
    });
    mockUrlFetchApp.fetch.mockReturnValueOnce({
      getContentText: () =>
        JSON.stringify({
          data: [
            {
              item: {
                /* no properties */
              },
              /* no content */
            },
          ],
        }),
    });

    crawlRoom();
    expect(mockRange.setValues).toHaveBeenCalledWith([
      ["", "", "", "", 1, "", ""],
    ]);
  });

  it("should clear content when sheet has existing data", () => {
    mockSheet.getLastRow.mockReturnValue(2); // Simulate existing data
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [] }),
      })
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [] }),
      });

    crawlRoom();

    expect(mockRange.clearContent).toHaveBeenCalled();
  });

  it("should leave the sheet untouched when a fetch fails", () => {
    mockSheet.getLastRow.mockReturnValue(2); // Simulate existing data
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        // The data server once answered a collects fetch with a plain
        // "OK" instead of JSON (2026-11-03); parsing that must fail
        // before the sheet is cleared, not after.
        getContentText: () => "OK",
      });

    expect(() => crawlRoom()).toThrow();
    expect(mockRange.clearContent).not.toHaveBeenCalled();
    expect(mockRange.setValues).not.toHaveBeenCalled();
  });

  it("should fetch directly from the ROOM API when it works", () => {
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({
            data: [
              {
                item: {
                  name: "Test Item",
                  picture: { url: "http://example.com/test.jpg" },
                  url: "http://example.com/item",
                  price: 1000,
                },
                content: "Test comment",
              },
            ],
          }),
      });

    crawlRoom();

    const urls = mockUrlFetchApp.fetch.mock.calls.map(([url]) => url);
    expect(urls).toEqual([
      "https://room.rakuten.co.jp/api/room-user-1/collections",
      "https://room.rakuten.co.jp/api/1/collects?limit=100",
    ]);
    expect(mockRange.setValues).toHaveBeenCalledWith([
      [
        "Test Item",
        "http://example.com/test.jpg",
        "http://example.com/item",
        1000,
        1,
        "collection1",
        "Test comment",
      ],
    ]);
    expect(Logger.log).toHaveBeenCalledWith("source: room-api");
  });

  it("should fall back to the home data server when the ROOM API fetch fails", () => {
    mockUrlFetchApp.fetch
      .mockImplementationOnce(() => {
        throw new Error("network error");
      })
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({
            data: [
              {
                item: {
                  name: "Test Item",
                  picture: { url: "http://example.com/test.jpg" },
                  url: "http://example.com/item",
                  price: 1000,
                },
                content: "Test comment",
              },
            ],
          }),
      });

    crawlRoom();

    const urls = mockUrlFetchApp.fetch.mock.calls.map(([url]) => url);
    expect(urls).toEqual([
      "https://room.rakuten.co.jp/api/room-user-1/collections",
      "http://home-api.example/collections.json",
      "http://home-api.example/collects_1.json",
    ]);
    expect(mockRange.setValues).toHaveBeenCalledWith([
      [
        "Test Item",
        "http://example.com/test.jpg",
        "http://example.com/item",
        1000,
        1,
        "collection1",
        "Test comment",
      ],
    ]);
    expect(Logger.log).toHaveBeenCalledWith("source: fallback (a.ze.gs)");
  });

  it("should leave the sheet untouched when both the ROOM API and the fallback fail", () => {
    mockSheet.getLastRow.mockReturnValue(2); // Simulate existing data
    mockUrlFetchApp.fetch
      .mockImplementationOnce(() => {
        throw new Error("ROOM API down");
      })
      .mockImplementationOnce(() => {
        throw new Error("fallback down");
      });

    expect(() => crawlRoom()).toThrow();
    expect(mockRange.clearContent).not.toHaveBeenCalled();
    expect(mockRange.setValues).not.toHaveBeenCalled();
  });

  it("should use the fallback when ROOM_USER_ID is not set", () => {
    global.PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: vi.fn(
        (name) => ({ DATA_BASE_URL: "http://home-api.example" })[name] || null
      ),
    });
    mockUrlFetchApp.fetch
      .mockReturnValueOnce({
        getContentText: () =>
          JSON.stringify({ data: [{ id: 1, name: "collection1" }] }),
      })
      .mockReturnValueOnce({
        getContentText: () => JSON.stringify({ data: [] }),
      });

    crawlRoom();

    const urls = mockUrlFetchApp.fetch.mock.calls.map(([url]) => url);
    expect(urls).toEqual([
      "http://home-api.example/collections.json",
      "http://home-api.example/collects_1.json",
    ]);
    expect(Logger.log).toHaveBeenCalledWith("source: fallback (a.ze.gs)");
  });
});

describe("spreadsheet id override (staging)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockSpreadsheet.getSheetByName.mockImplementation(() => ({
      getDataRange: vi.fn().mockReturnValue({
        getValues: vi.fn().mockReturnValue([["h"], ["v"]]),
      }),
    }));
  });

  it("falls back to the production spreadsheet id when the property is unset", () => {
    doGet({ parameter: {} });
    expect(global.SpreadsheetApp.openById).toHaveBeenCalledWith(
      "1yxnHhfAT2cEF50hvajHoLMsW_sf7JlxnjFKeDFLNAkc"
    );
  });

  it("uses the SPREADSHEET_ID script property when it is set", () => {
    global.PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: vi.fn((name) =>
        name === "SPREADSHEET_ID" ? "staging-sheet-id" : null
      ),
    });
    doGet({ parameter: {} });
    expect(global.SpreadsheetApp.openById).toHaveBeenCalledWith("staging-sheet-id");
  });
});
