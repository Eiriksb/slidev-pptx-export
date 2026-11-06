import { test } from "node:test";
import assert from "node:assert/strict";
import { writeAtomic } from "../src/write.mjs";

test("the file appears under its name only once it is complete", async (t) => {
  const calls = [];
  const fs = { writeFile: async (p) => calls.push(["write", p]), rename: async (a, b) => calls.push(["rename", a, b]), rm: async () => {} };
  await writeAtomic("/d/out.pptx", Buffer.from("x"), { fs });
  assert.deepEqual(calls, [["write", "/d/out.pptx.partial"], ["rename", "/d/out.pptx.partial", "/d/out.pptx"]]);
});
test("a write that fails leaves no .partial behind", async (t) => {
  const removed = [];
  const fs = { writeFile: async () => {}, rename: async () => { throw new Error("EXDEV"); }, rm: async (p) => removed.push(p) };
  await assert.rejects(writeAtomic("/d/out.pptx", Buffer.from("x"), { fs }), /EXDEV/);
  assert.deepEqual(removed, ["/d/out.pptx.partial"]);
});
