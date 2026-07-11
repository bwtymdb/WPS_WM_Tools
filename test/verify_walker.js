// verify_walker.js - exercise the REAL docToMarkdown with a mocked WPS document object.
// Confirms the JS walker logic (headings via OutlineLevel, table dedup via Range.Start, paragraphs)
// before the user tests in WPS. Node can't run WPS COM, so we mock the doc object.
const fs = require("fs"), vm = require("vm");
const sandbox = { console: console };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync("js/finalize.js", "utf8"), sandbox, { filename: "finalize.js" });

// mock a 2x2 table (cells carry the  cell-end marker, like real WPS)
const mockTable = {
    Range: { Start: 100 },
    Rows: { Count: 2, Item: function (r) {
        const rows = [
            { Cells: { Count: 2, Item: function (c) { return { Range: { Text: "H" + c + "" } }; } } },
            { Cells: { Count: 2, Item: function (c) { return { Range: { Text: "v" + c + "" } }; } } }
        ];
        return rows[r - 1];
    } }
};

const paras = [
    { OutlineLevel: 1, Range: { Text: "Title\r", Tables: { Count: 0 } } },
    { OutlineLevel: 10, Range: { Text: "Body para one.\r", Tables: { Count: 0 } } },
    { OutlineLevel: 2, Range: { Text: "Section Two\r", Tables: { Count: 0 } } },
    { OutlineLevel: 10, Range: { Text: "c1", Tables: { Count: 1, Item: function () { return mockTable; } } } },
    { OutlineLevel: 10, Range: { Text: "c2", Tables: { Count: 1, Item: function () { return mockTable; } } } },
    { OutlineLevel: 10, Range: { Text: "Trailing paragraph.\r", Tables: { Count: 0 } } }
];
const mockDoc = { Paragraphs: { Count: paras.length, Item: function (i) { return paras[i - 1]; } } };

const out = sandbox.docToMarkdown(mockDoc);
console.log("=== docToMarkdown output (raw, with escapes) ===");
console.log(JSON.stringify(out));
console.log("=== rendered ===");
console.log(out);
