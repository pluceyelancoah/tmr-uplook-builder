// Regression check against real farm reports kept in raw/ (git-ignored). Run `npm run snapshot` to accept new output.
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const { sampleFiles, summarise, snapshotPath } = require("./samples.js");

const files = sampleFiles().filter(f => fs.existsSync(snapshotPath(f)));
if (!files.length) test("private samples", { skip: "no raw/*.expected.json snapshots on this machine" }, () => {});

for (const f of files) {
  test(`sample: ${f}`, () => {
    const expected = JSON.parse(fs.readFileSync(snapshotPath(f), "utf8"));
    assert.deepEqual(JSON.parse(JSON.stringify(summarise(f))), expected);
  });
}
