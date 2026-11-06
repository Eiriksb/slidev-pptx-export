import { test } from "node:test";
import assert from "node:assert/strict";
import { parseArgs, applyFlags } from "../src/cli-args.mjs";

test("--no-animations gives animations: false, and the last of --animations / --no-animations wins", () => {
  assert.equal(parseArgs(["a.md"]).animations, null);
  assert.equal(parseArgs(["a.md", "--animations"]).animations, true);
  assert.equal(parseArgs(["a.md", "--no-animations"]).animations, false);
  assert.equal(parseArgs(["--animations", "a.md", "--no-animations"]).animations, false);
  assert.equal(parseArgs(["--no-animations", "--animations"]).animations, true);
});

test("--no-animations doesn't imply --render and doesn't turn it off", () => {
  assert.equal(parseArgs(["--no-animations"]).render, false);
  assert.equal(parseArgs(["--render", "--no-animations"]).render, true);
});

test("--no-animations wins over a config with render.animations: true, which --animations alone would not override", () => {
  const config = { render: { animations: true, browser: "/b" } };
  applyFlags(config, parseArgs(["--render", "--no-animations"]));
  assert.deepEqual(config.render, { animations: false, browser: "/b" });
  const untouched = { render: { animations: true } };
  applyFlags(untouched, parseArgs(["--render"]));
  assert.equal(untouched.render.animations, true);
  const on = {};
  applyFlags(on, parseArgs(["--animations", "--videos", "native"]));
  assert.deepEqual(on.render, { videos: "native", animations: true });
});
