const { test } = require("node:test");
const assert = require("node:assert");
const { saveToBuffer, saveToFileSync } = require("../index.js");
const fs = require("fs");

test("save to buffer", async (t) => {
  const buffer = await saveToBuffer({
    name: "test",
    sheets: [
      {
        name: "test",
        cells: [
          {
            col: 0,
            row: 0,
            value: "header 1",
            celType: "string",
          },
          {
            col: 1,
            row: 0,
            value: "header 2",
            celType: "string",
          },
          {
            col: 0,
            row: 1,
            value: 1,
            celType: "number",
          },
          {
            col: 1,
            row: 1,
            value: 2,
            celType: "number",
          },
          {
            col: 0,
            row: 2,
            value: "https://example.com",
            celType: "link",
          },
        ],
      },
    ],
  });

  assert.ok(buffer instanceof Buffer);
  assert.ok(buffer.length > 0);

  fs.writeFileSync("./temp/test_node.xlsx", buffer);
});

test("save to file sync", async (t) => {
  saveToFileSync(
    {
      name: "test",
      sheets: [
        {
          name: "test",
          cells: [
            {
              col: 0,
              row: 0,
              value: "header 1",
              celType: "string",
            },
            {
              col: 1,
              row: 0,
              value: "header 2",
              celType: "string",
            },
            {
              col: 0,
              row: 1,
              value: 1,
              celType: "number",
            },
            {
              col: 1,
              row: 1,
              value: 2,
              celType: "number",
            },
            {
              col: 0,
              row: 2,
              value: "https://example.com",
              celType: "link",
            },
          ],
        },
      ],
    },
    "./temp/test_node_sync.xlsx"
  );
});
