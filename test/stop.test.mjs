import { test } from "node:test";
import assert from "node:assert/strict";
import { onStop, runStop, isStopping } from "../src/stop.mjs";
import { writeAtomic } from "../src/write.mjs";

test("runStop runs what is registered once, and survives a failing cleanup", async () => {
  const order = [];
  onStop(() => order.push("first"));
  onStop(() => {
    throw new Error("boom");
  });
  const forget = onStop(() => order.push("forgotten"));
  onStop(async () => order.push("last"));
  forget();
  assert.equal(isStopping(), false);
  await runStop();
  assert.deepEqual(order.sort(), ["first", "last"]);
  await runStop();
  assert.equal(order.length, 2);
});

test("runStop gives up on a cleanup that hangs after the deadline", async () => {
  // (runs after the first test, so the registry is already closed: a hanging one runs at once)
  const started = Date.now();
  onStop(() => new Promise(() => {}));
  await runStop(100);
  assert.ok(Date.now() - started < 2000);
});

// These two come last: they depend on the stopped state the tests above left behind.
test("once stopping, a cleanup registered later runs at once", async () => {
  assert.equal(isStopping(), true);
  let ran = false;
  onStop(() => {
    ran = true;
  });
  assert.equal(ran, true);
});

test("once stopping, writeAtomic refuses to write", async () => {
  const calls = [];
  const fs = { writeFile: async () => calls.push("write"), rename: async () => calls.push("rename"), rm: async () => {} };
  await assert.rejects(writeAtomic("/nowhere/out.pptx", Buffer.from("x"), { fs }), /stopped/);
  assert.deepEqual(calls, []);
});
