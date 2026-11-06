import { test } from "node:test";
import assert from "node:assert/strict";
import { onStop, runStop } from "../src/stop.mjs";

test("runStop runs what is registered, newest first, once, and survives a failing cleanup", async () => {
  const order = [];
  onStop(() => order.push("first"));
  onStop(() => {
    throw new Error("boom");
  });
  const forget = onStop(() => order.push("forgotten"));
  onStop(async () => order.push("last"));
  forget();
  await runStop();
  assert.deepEqual(order, ["last", "first"]);
  await runStop();
  assert.deepEqual(order, ["last", "first"]);
});
