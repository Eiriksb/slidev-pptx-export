# Animation Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** An opt-in render-mode export whose slide transitions, click reveals and script-driven motion play in a PowerPoint for the web slide show, driven by keyboard or clicker, in Slidev's click order.

**Architecture:** Each click state of a slide is captured and diffed. A click with no script motion becomes native Fade effects on the current PowerPoint slide; a click with script motion starts a new PowerPoint slide whose MP4 clip starts by itself. The writer gains a timing module and otherwise knows nothing about Slidev; everything else is new modules beside the existing render pipeline, used only when the flag is on.

**Tech Stack:** Node ES modules, `node:test`, Playwright Chromium, hand-written OOXML, `ffmpeg` or GStreamer found at run time.

**Spec:** `docs/superpowers/specs/2026-11-05-animation-export-design.md`

## Global Constraints

- Flag: `--animations` on the CLI, `render: { animations: true }` in config. It implies render mode and `render.videos: "native"`.
- With the flag off, output is unchanged, except the media extension id fix in Task 1.
- Media extension id: `{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}`.
- Click effects: Fade, 400 ms. Clips: width 1920, 30 fps, CRF 23, muted, poster is the first frame.
- Nothing is added to the root `dependencies` or `devDependencies`. No binary is downloaded at install time.
- The test deck and its organisations are never named in the repo, commits or PRs. Refer to it as "the test deck". Its local copy and its install are outside the repo.
- Anything that could not be animated prints a line starting `Notice:` through `console.warn`, as the rest of render mode does.
- Commits end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`. PR bodies end with `🤖 Generated with [Claude Code](https://claude.com/claude-code)`.
- `npm test`, `npm run test:render` and `lintPackage` (from `test/lint.mjs`, returning `[]`) pass after every task.
- The feature branch is based on `render-map-slides` while Eiriksb/slidev-pptx-export#2 is open.

Facts checked on Slidev 53 and safe to rely on: `window.__slidev__.nav.clicksTotal` is the click total of the current slide (read it through `.value` if it is a ref); a slide's transition is `route.meta?.transition ?? __slidev__.configs.transition`, a string such as `"slide-left | fade"` where the first half is forward; it belongs to the slide being arrived at, as `p:transition` does; the extractor already leaves out `v-click` items that are still hidden.

## Review Focus

- A deck that is not 16:9 keeps its shape in clips, with even pixel dimensions, because H.264 requires them. Pinned in Task 5.
- A slide with no clicks exports as one PowerPoint slide with no `p:timing` and no empty extra slide. Pinned in Task 8.
- A deck's own video on a slide with clicks stays a video and keeps its own autoplay and loop, with On Click effects after it. Pinned in Tasks 2 and 8.
- An element that hides and later returns, and two identical elements on one slide, each export correctly. Pinned in Task 6.
- A page whose script never stops gives up at the time cap, the click becomes a Fade with a notice, and no temporary files are left. Pinned in Tasks 5 and 7.

---

### Task 1: Fix the media extension id in its own PR

**Files:**
- Modify: `src/ooxml/shapes.mjs:95`
- Test: `test/video.test.mjs`

**Interfaces:**
- Consumes: nothing.
- Produces: `MEDIA_EXT` with the correct id, on `main` through its PR and on the feature branch by cherry-pick.

- [ ] **Step 1: Create the fix branch in a separate worktree**

Run: `git worktree add "$TMPDIR/fix-media-id" -b fix-video-extension-id origin/main`, using the session scratchpad as `$TMPDIR`, then `npm ci` there.
Expected: a clean checkout of `origin/main`.

- [ ] **Step 2: Write the failing test in `test/video.test.mjs`**

```js
test("a video's embedded copy uses the extension id PowerPoint writes", async () => {
  const { z } = await build([video()]);
  assert.match(z.text("ppt/slides/slide1.xml"), /<p:ext uri="\{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230\}"><p14:media /);
});
```

- [ ] **Step 3: Run it to verify it fails**

Run: `node --test test/video.test.mjs`
Expected: FAIL, the slide has `3DA1BC7CA4C3`.

- [ ] **Step 4: Change `MEDIA_EXT` in `src/ooxml/shapes.mjs`**

- [ ] **Step 5: Run the tests**

Run: `npm test`
Expected: PASS.

- [ ] **Step 6: Commit, push and open the PR against `main`**

Commit message: `Native videos: use the extension id PowerPoint writes, so PowerPoint for the web can play them`. The PR body says this affects the default export, how it was found (a video the web inserted itself, compared with ours) and what was checked: the web only, in the app's browser pane.

- [ ] **Step 7: Bring the fix to the feature branch**

Run: `git cherry-pick <fix commit>` on the feature branch, then remove the temporary worktree with `git worktree remove`.
Expected: `npm test` passes on the feature branch.

### Task 2: Click effects, clip timing and transitions in the writer

**Files:**
- Create: `src/ooxml/timing.mjs`
- Modify: `src/ooxml/shapes.mjs` (`slideXml`), `src/ooxml/index.mjs` (`shapeToXml`, `writePptx`)
- Test: `test/timing.test.mjs`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `animationTimingXml({ videos, effects }) -> string`, `""` when both are empty. `videos`: the writer's existing `[{ spid, autoplay, loop, muted, duration }]`. `effects`: `[{ spid, kind, fill, click, dir }]`, `click` numbered from 1 and `dir` `"in"` or `"out"`.
  - `transitionXml(t) -> string`. `t`: `"fade"`, `{ push: "l" | "r" | "u" | "d" }` or `undefined`.
  - `slideXml(shapesXml, bgXml = "", timingXml = "", transitionXml = "")`.
  - SlideSpec fields, ignored unless set: `spec.animated` (use `animationTimingXml`), `spec.transition`, and on any shape `shape.effects: [{ click, dir }]`.

- [ ] **Step 1: Write the failing tests in `test/timing.test.mjs`**

```js
const count = (xml, re) => (xml.match(re) || []).length;
const fx = (spid, click, dir, kind = "image", fill = false) => ({ spid, kind, fill, click, dir });

test("effects of one click: the first is On Click, the rest With Previous", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(4, 1, "in"), fx(5, 1, "out")] });
  assert.equal(count(xml, /nodeType="clickEffect"/g), 1);
  assert.equal(count(xml, /nodeType="withEffect"/g), 1);
  assert.match(xml, /presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"/);
  assert.match(xml, /<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="\d+" dur="400"\/>/);
  assert.match(xml, /presetID="10" presetClass="exit"/);
  assert.match(xml, /<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="\d+" dur="400"\/>/);
  assert.match(xml, /<p:strVal val="hidden"\/>/);
});

test("each click is its own step, in order", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(5, 2, "in"), fx(4, 1, "in"), fx(6, 3, "in")] });
  assert.equal(count(xml, /nodeType="clickEffect"/g), 3);
  assert.equal(count(xml, /<p:cond delay="indefinite"\/>/g), 3);
  assert.ok(xml.indexOf('spid="4"') < xml.indexOf('spid="5"') && xml.indexOf('spid="5"') < xml.indexOf('spid="6"'));
});

test("a text shape with two effects gets grpId 0 and 1 and a bldP for each; a fill adds animBg", () => {
  const xml = animationTimingXml({ videos: [], effects: [fx(7, 1, "in", "text", true), fx(7, 2, "out", "text", true), fx(8, 1, "in")] });
  assert.equal(count(xml, /grpId="0"/g), 2);
  assert.equal(count(xml, /grpId="1"/g), 2);
  assert.match(xml, /<p:bldLst><p:bldP spid="7" grpId="0" animBg="1"\/><p:bldP spid="7" grpId="1" animBg="1"\/><\/p:bldLst>/);
  assert.doesNotMatch(xml, /spid="8" grpId/);
});

test("a video that starts by itself is written as PowerPoint for the web writes it", () => {
  const xml = animationTimingXml({ videos: [{ spid: 6, autoplay: true, loop: false, muted: true, duration: 3 }], effects: [] });
  assert.match(xml, /<p:cond evt="onBegin" delay="0"><p:tn val="2"\/><\/p:cond>/);
  assert.match(xml, /presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"/);
  assert.match(xml, /cmd="playFrom\(0\.0\)"><p:cBhvr><p:cTn id="\d+" dur="1" fill="hold"\/>/);
  assert.match(xml, /<p:video><p:cMediaNode mute="1"><p:cTn id="\d+" fill="hold" display="0">/);
  assert.doesNotMatch(xml, /togglePause|repeatCount/);
});

test("a looping video keeps repeatCount; one that does not autoplay has only its media node", () => {
  const looping = animationTimingXml({ videos: [{ spid: 6, autoplay: true, loop: true, muted: true }], effects: [] });
  assert.match(looping, /repeatCount="indefinite"/);
  const waiting = animationTimingXml({ videos: [{ spid: 6, autoplay: false, muted: false }], effects: [] });
  assert.doesNotMatch(waiting, /mediacall/);
  assert.match(waiting, /<p:video><p:cMediaNode>/);
});

test("the clip starts first and On Click effects follow in the same main sequence, with unique ids", () => {
  const xml = animationTimingXml({ videos: [{ spid: 6, autoplay: true, muted: true }], effects: [fx(9, 1, "in", "text")] });
  assert.equal(count(xml, /nodeType="mainSeq"/g), 1);
  assert.ok(xml.indexOf("mediacall") < xml.indexOf("clickEffect"));
  const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length);
});

test("transitions", () => {
  assert.equal(transitionXml("fade"), "<p:transition><p:fade/></p:transition>");
  assert.equal(transitionXml({ push: "l" }), '<p:transition><p:push dir="l"/></p:transition>');
  assert.equal(transitionXml(undefined), "");
});

test("the writer: effects reach the timing, the transition comes before it, the package is clean", async () => {
  const shapes = [{ kind: "image", x: 1, y: 1, w: 2, h: 1, path: PIC, effects: [{ click: 1, dir: "in" }] }];
  const buffer = await writePptx([{ layout: "blank", shapes, animated: true, transition: "fade" }], DECK);
  assert.deepEqual(await lintPackage(buffer), []);
  const slide = (await unzip(buffer)).text("ppt/slides/slide1.xml");
  assert.match(slide, /<\/p:clrMapOvr><p:transition><p:fade\/><\/p:transition><p:timing>/);
  assert.match(slide, /nodeType="clickEffect"/);
});
```

- [ ] **Step 2: Run them to verify they fail**

Run: `node --test test/timing.test.mjs`
Expected: FAIL, module not found.

- [ ] **Step 3: Implement `src/ooxml/timing.mjs`**

Fade in: a `p:set` of `style.visibility` to `visible`, then `p:animEffect transition="in" filter="fade"` of 400 ms. Fade out: `p:animEffect transition="out"` of 400 ms, then a `p:set` to `hidden` with `delay="399"`. Number `cTn` ids in document order. Only `kind: "text"` shapes get `grpId`, counted per `spid`. Do not reuse or change the existing `timingXml` in `shapes.mjs`: the default export stays as it is.

- [ ] **Step 4: Wire it into the writer**

`slideXml` writes `transitionXml` between `p:clrMapOvr` and the timing. `shapeToXml` pushes each `shape.effects` entry to `env.effects` with `spid: id`, `kind: shape.kind`, `fill: !!shape.fill`. `writePptx` uses `spec.animated ? animationTimingXml({ videos: env.videos, effects: env.effects }) : timingXml(env.videos)`.

- [ ] **Step 5: Run the tests**

Run: `npm test`
Expected: PASS, with every existing test unchanged.

- [ ] **Step 6: Commit**

Message: `Writer: click effects, a clip that starts by itself and slide transitions, for animated exports`.

### Task 3: Early web checkpoint (needs the user)

**Files:**
- Create outside the repo, in the scratchpad: a throwaway generator using Task 2's writer and the spike's clips.
- Modify: the spec, with the results.

**Interfaces:**
- Consumes: Task 2's `writePptx` fields.
- Produces: a recorded answer on whether a full-slide picture may enter above a held clip, used in Task 6 Step 5.

This must be run from the session that holds the conversation, because it needs the user and the browser pane.

- [ ] **Step 1: Generate a small test file with labelled slides**

Slides: a map picture as JPEG followed by a clip slide with a JPEG poster and no click trigger; a clip slide with an On Click full-slide JPEG picture Fade above it; a clip slide with later On Click native text and a table.

- [ ] **Step 2: Ask the user to upload it in the browser pane and leave it open**

Copy it to the user's Downloads folder first. Expected: the file open in the editor.

- [ ] **Step 3: Run the show from the start with the keyboard and read the Animation Pane**

Record: how long the blank lasts before the poster, on first visit and again after going back; whether each clip starts by itself; whether a key press during a clip moves on; whether the picture enters above the held clip; whether On Click steps follow the clip in order.

- [ ] **Step 4: Update the spec's table and "Planning a slide" with the results, and commit**

Message: `Design: results of the early web checkpoint`. If a clip with later reveals misbehaves, stop and tell the user before Task 4: the design's core has changed.

### Task 4: Clip encoder

**Files:**
- Create: `src/render/encode.mjs`
- Test: `test/encode.test.mjs`

**Interfaces:**
- Consumes: a folder of JPEG frames named `f00000.jpg`, `f00001.jpg`, ...
- Produces:
  - `findEncoder({ run = spawnSync } = {}) -> { kind: "ffmpeg" | "gstreamer" } | null`
  - `encoderArgs(kind, { dir, count, fps, out }) -> { bin, args }`
  - `encodeClip(encoder, { dir, count, fps }, { run = spawnSync } = {}) -> Buffer`, throwing with the tool's first error line.
  - `ENCODER_HINT`: the text naming what to install, used in notices.

- [ ] **Step 1: Write the failing tests**

```js
test("ffmpeg with libx264 is preferred, then GStreamer with x264enc and mp4mux, then nothing", () => {
  const fake = (present) => (bin, args) => ({ status: present[bin] ? 0 : 1, stdout: bin === "ffmpeg" && present.x264 ? " V..... libx264 " : "" });
  assert.deepEqual(findEncoder({ run: fake({ ffmpeg: true, x264: true }) }), { kind: "ffmpeg" });
  assert.deepEqual(findEncoder({ run: fake({ ffmpeg: true, "gst-inspect-1.0": true }) }), { kind: "gstreamer" });
  assert.equal(findEncoder({ run: fake({}) }), null);
});

test("ffmpeg arguments: H.264, CRF 23, yuv420p, no B-frames, faststart", () => {
  const { bin, args } = encoderArgs("ffmpeg", { dir: "/d", count: 3, fps: 30, out: "/o.mp4" });
  assert.equal(bin, "ffmpeg");
  for (const a of ["libx264", "23", "yuv420p", "+faststart"]) assert.ok(args.includes(a), a);
});

test("GStreamer arguments ask x264enc for avc, so h264parse is not needed", () => {
  const { args } = encoderArgs("gstreamer", { dir: "/d", count: 3, fps: 30, out: "/o.mp4" });
  assert.ok(args.includes("video/x-h264,profile=high,stream-format=avc,alignment=au"));
  assert.ok(args.includes("quantizer=23") && args.includes("pass=qual"));
  assert.ok(!args.includes("h264parse"));
});

test("a real encode gives an MP4, when an encoder is installed", { skip: findEncoder() ? false : "no encoder installed" }, () => {
  // Writes 10 copies of a small JPEG from the fixtures to a temporary folder.
  const data = encodeClip(findEncoder(), makeFrames(10), {});
  assert.equal(data.subarray(4, 8).toString(), "ftyp");
});
```

- [ ] **Step 2: Run them to verify they fail**

Run: `node --test test/encode.test.mjs`
Expected: FAIL, module not found.

- [ ] **Step 3: Implement `src/render/encode.mjs`**

ffmpeg detection reads `ffmpeg -hide_banner -encoders` for `libx264`. GStreamer detection runs `gst-inspect-1.0 x264enc` and `gst-inspect-1.0 mp4mux`. GStreamer pipeline: `multifilesrc location=<dir>/f%05d.jpg index=0 stop-index=<count-1> caps=image/jpeg,framerate=<fps>/1 ! jpegdec ! videoconvert ! video/x-raw,format=I420 ! x264enc speed-preset=medium pass=qual quantizer=23 key-int-max=15 bframes=0 ! <caps above> ! mp4mux faststart=true ! filesink location=<out>`. Use the `spawnSync` argument array, never a shell string.

- [ ] **Step 4: Run the tests**

Run: `node --test test/encode.test.mjs`
Expected: PASS.

- [ ] **Step 5: Commit**

Message: `Render mode: encode recorded frames to H.264 with ffmpeg or GStreamer, whichever is installed`.

### Task 5: Stepped clock and click recorder

**Files:**
- Create: `src/render/clock.js`, `src/render/record.mjs`
- Modify: `src/render/extract.js` (optional `paths` option)
- Test: `test/render/record.test.mjs`, `test/render/extract.test.mjs`

**Interfaces:**
- Consumes: `.slidev-page-<n>`, `window.__slidev__.nav.next()`.
- Produces:
  - `window.__pptxExtract(pageEl, canvas, { paths: true })` adds `node.elPath` to each node: a selector from the page root of `tag.class.class` steps joined by ` > `, with the pseudo-element appended for a pseudo-element's node. Without the option, the result is exactly as today.
  - `clipViewport(canvas) -> { width: 1920, height }`, `height` rounded to even.
  - `recordClick({ browser, url, sel, canvas, hide, keep = null, show = [], tmpRoot, fps = 30, maxSeconds = 40, timeout = 30000, setup = async () => {} }) -> RecordResult`. `keep` is the selector of the picture's own element.
  - `RecordResult`: `{ dir, count, fps, motionFrames, settled, changedText: string[], changedVisibly: boolean }`. `dir` holds `f%05d.jpg`, frame 0 being the state before the click. The caller removes `dir`.
  - `MOTION_FRAMES = 12` (0.4 s at 30 fps).

- [ ] **Step 1: Write the failing tests in `test/render/record.test.mjs`**

Each uses a stub page served through `setup` (`context.route`) with a `.slidev-page-1` and a `window.__slidev__ = { nav: { next() { ... } } }`.

```js
test("a requestAnimationFrame animation is recorded at an exact frame rate, however slowly it paints", opts, async (t) => {
  const r = await recordStub(t, ANIMATE_RAF_1S);
  assert.ok(r.settled);
  assert.ok(r.motionFrames >= 28 && r.motionFrames <= 32, `motionFrames ${r.motionFrames}`);
  assert.ok(r.changedVisibly);
  assert.equal(readdirSync(r.dir).length, r.count);
  assert.equal(frameCount(r.dir), r.count);
});

test("a CSS transition after the click is drawn but is not script motion", opts, async (t) => {
  const r = await recordStub(t, CSS_TRANSITION_500MS);
  assert.ok(r.settled);
  assert.ok(r.motionFrames < MOTION_FRAMES);
  assert.ok(r.changedVisibly);
});

test("what `hide` names stays out of the frames, including an element that only appears after the click", opts, async (t) => {
  const r = await recordStub(t, CAPTION_APPEARS, { hide: [".caption"] });
  assert.ok(!frameHasColor(r.dir, r.count - 1, CAPTION_COLOR));
});

test("a selector that matches the picture's own element or an ancestor of it is not applied", opts, async (t) => {
  const r = await recordStub(t, ANIMATE_RAF_1S, { hide: ["div", ".caption"], keep: ".map" });
  assert.ok(r.changedVisibly);
});

test("a hidden text that changes after the click is reported", opts, async (t) => {
  const r = await recordStub(t, CAPTION_CHANGES_AT_500MS, { hide: [".caption"] });
  assert.deepEqual(r.changedText, [".caption"]);
});

test("a page whose script never stops gives up at the cap", opts, async (t) => {
  const r = await recordStub(t, RAF_FOREVER, { maxSeconds: 1 });
  assert.equal(r.settled, false);
  assert.ok(r.count <= 31);
});

test("a clip keeps a 4:3 deck's shape, with even dimensions", () => {
  assert.deepEqual(clipViewport({ w: 980, h: 552 }), { width: 1920, height: 1082 });
  assert.deepEqual(clipViewport({ w: 1024, h: 768 }), { width: 1920, height: 1440 });
  assert.deepEqual(clipViewport({ w: 1440, h: 810 }), { width: 1920, height: 1080 });
});
```

Add to `test/render/extract.test.mjs`: with `{ paths: true }` a node has `elPath`, a pseudo-element's ends in `::after`, and without the option no node has `elPath`.

- [ ] **Step 2: Run them to verify they fail**

Run: `node --test test/render/record.test.mjs test/render/extract.test.mjs`
Expected: FAIL, module not found and missing `elPath`.

- [ ] **Step 3: Implement `src/render/clock.js`, loaded with `addInitScript` before the page's own scripts**

Real time until `window.__startVirtualClock(rootSelector)`. Override `requestAnimationFrame`, `cancelAnimationFrame`, `performance.now`, `Date` (subclass, so `new Date()` and `Date.now()` follow), `setTimeout` and `clearTimeout`, using negative ids for virtual callbacks so a real id still cancels correctly. `window.__stepClock(dtMs)` advances the clock, runs due timers then due animation-frame callbacks with the virtual time, pauses every `document.getAnimations()` entry and moves its `currentTime` by `dtMs` (`finish()` when past its end), and returns `{ at, frames, timers, running, images, sinceMutation }`. A `MutationObserver` on the root records `lastMutation` and, past the first 5 steps, the targets of text changes.

- [ ] **Step 4: Implement `recordClick` in `src/render/record.mjs`**

New context with `clipViewport(canvas)` and `deviceScaleFactor: 1`. Go to `url`, wait for `sel`, fonts and quiet using the existing wait. Add the hide stylesheet (`visibility:hidden!important`), first dropping any selector whose `closest` matches the kept picture element. Always hide `[data-pptx-layer="top"]` and leave out names in `show`. Start the clock, take frame 0 with CDP `Page.captureScreenshot` (`jpeg`, quality 92), call `nav.next()`, then step: if `images` is non-zero wait 60 ms of real time, take the frame, count `motionFrames` when `frames > 0`, stop when `frames === 0 && running === 0 && images === 0 && sinceMutation > 600`, or at `maxSeconds`. `changedVisibly`: frame 0 and the last frame differ. Close the context in a `finally`.

- [ ] **Step 5: Add the `paths` option to `extract.js`**

`sources.get(n)` already holds each node's element and pseudo-element. Compute `elPath` only when the option is set.

- [ ] **Step 6: Run the tests**

Run: `npm run test:render`
Expected: PASS, skipping only what needs a missing tool.

- [ ] **Step 7: Commit**

Message: `Render mode: record a click frame by frame on a stepped clock, with native content left out`.

### Task 6: Slide plan

**Files:**
- Create: `src/render/animation-plan.mjs`
- Test: `test/animation-plan.test.mjs`

**Interfaces:**
- Consumes: scene nodes; `MOTION_FRAMES`.
- Produces:
  - `stateSignature(node) -> string`, for every kind. Includes kind, rect rounded to 0.5 px, and every other field except `id`, `above`, `elPath` and `label`, with `Buffer` values replaced by their sha1.
  - `mergeRuns(states: Node[][]) -> Run[]`, `Run = { node, from, to }`, in merged z-order.
  - `clipTarget(before: Node[], after: Node[]) -> Node | null`: the largest `raster` node in `after` whose signature is not in `before`.
  - `clickKind({ target, record, encoder }) -> { kind: "reveal" | "motion", reason: string | null }`.
  - `planSlide(states, kinds) -> Segment[]`, `kinds[k]` being click `k`'s kind (`kinds[0]` unused). `Segment = { first, last, runs }`, where `first > 0` means the segment starts at a motion click. `from` and `to` in its runs are state numbers of the whole slide, as `effectsOf` expects.
  - `effectsOf(run, segment) -> [{ click, dir }]`.

- [ ] **Step 1: Write the failing tests**

```js
const box = (id, x = 0) => ({ kind: "box", id, rect: { x, y: 0, w: 10, h: 10 }, fill: { hex: id } });
const raster = (hash, w = 1000) => ({ kind: "raster", rect: { x: 0, y: 0, w, h: 500 }, png: Buffer.from(hash) });

test("a signature covers every kind, ignores ids and paths, and hashes bytes", () => {
  const table = { kind: "table", id: 1, rect: { x: 0, y: 0, w: 1, h: 1 }, rows: [["a"]], elPath: "x" };
  assert.equal(stateSignature(table), stateSignature({ ...table, id: 2, elPath: "y" }));
  assert.notEqual(stateSignature(table), stateSignature({ ...table, rows: [["b"]] }));
  assert.equal(stateSignature(raster("a")), stateSignature(raster("a")));
  assert.notEqual(stateSignature(raster("a")), stateSignature(raster("b")));
});

test("runs: what stays spans the states; what enters is inserted after its predecessor", () => {
  const runs = mergeRuns([[box("a"), box("c")], [box("a"), box("b"), box("c")]]);
  assert.deepEqual(runs.map((r) => [r.node.fill.hex, r.from, r.to]), [["a", 0, 1], ["b", 1, 1], ["c", 0, 1]]);
});

test("an element that hides and returns is two runs", () => {
  const runs = mergeRuns([[box("a")], [], [box("a")]]);
  assert.deepEqual(runs.map((r) => [r.from, r.to]), [[0, 0], [2, 2]]);
});

test("two identical elements are two runs, not one", () => {
  assert.equal(mergeRuns([[box("a"), box("a")]]).length, 2);
});

test("when two states disagree on order, the later order wins by repeating the node", () => {
  const runs = mergeRuns([[box("a"), box("b")], [box("b"), box("a")]]);
  assert.equal(runs.length, 3);
  const inState1 = runs.filter((r) => r.from <= 1 && r.to >= 1).map((r) => r.node.fill.hex);
  assert.deepEqual(inState1, ["b", "a"]);
});

test("a clip replaces the largest picture that changed", () => {
  const big = raster("new");
  assert.equal(clipTarget([raster("old")], [raster("logo", 50), big]), big);
  assert.equal(clipTarget([raster("same")], [raster("same")]), null);
});

test("click kinds", () => {
  const motion = { motionFrames: 40, settled: true, changedVisibly: true };
  assert.deepEqual(clickKind({ target: null, record: null, encoder: {} }), { kind: "reveal", reason: null });
  assert.deepEqual(clickKind({ target: {}, record: { ...motion, motionFrames: 3 }, encoder: {} }), { kind: "reveal", reason: null });
  assert.equal(clickKind({ target: {}, record: motion, encoder: {} }).kind, "motion");
  assert.match(clickKind({ target: {}, record: motion, encoder: null }).reason, /no encoder/);
  assert.match(clickKind({ target: {}, record: { ...motion, settled: false }, encoder: {} }).reason, /never settled/);
  assert.match(clickKind({ target: {}, record: { ...motion, changedVisibly: false }, encoder: {} }).reason, /no visible change/);
});

test("a motion click starts a new segment; later reveals stay on it", () => {
  const states = [[box("a")], [box("a"), box("b")], [raster("x")], [raster("x"), box("c")]];
  const segs = planSlide(states, [null, "reveal", "motion", "reveal"]);
  assert.deepEqual(segs.map((s) => [s.first, s.last]), [[0, 1], [2, 3]]);
});

test("effects are numbered within the segment", () => {
  const seg = { first: 2, last: 4 };
  assert.deepEqual(effectsOf({ from: 3, to: 3 }, seg), [{ click: 1, dir: "in" }, { click: 2, dir: "out" }]);
  assert.deepEqual(effectsOf({ from: 2, to: 4 }, seg), []);
});
```

- [ ] **Step 2: Run them to verify they fail**

Run: `node --test test/animation-plan.test.mjs`
Expected: FAIL, module not found.

- [ ] **Step 3: Implement `src/render/animation-plan.mjs`**

`mergeRuns` walks the states in order and keeps the merged list. For each node of a state, number duplicate signatures as `signaturesOf` does. If that numbered signature has a run that ended in the previous state and sits after the previous node's run in the merged list, extend it. Otherwise start a new run and insert it directly after the previous node's run, or at the start of the list. `clickKind` checks in order: no target, too few motion frames, no encoder, not settled, no visible change.

- [ ] **Step 4: Run the tests**

Run: `node --test test/animation-plan.test.mjs`
Expected: PASS.

- [ ] **Step 5: Apply the early checkpoint's result**

Only if Task 3 found a full-slide picture cannot enter above a held clip: add a test that a click changing a full-slide picture on a segment after a clip starts a new segment, and implement it in `planSlide`. Otherwise skip this step.

- [ ] **Step 6: Commit**

Message: `Render mode: plan an animated slide from its click states`.

### Task 7: Capture every click state

**Files:**
- Modify: `src/render/capture.mjs`, `src/render/scene.mjs`
- Test: `test/render/capture.test.mjs`, `test/video.test.mjs` or `test/render-scene.test.mjs`

**Interfaces:**
- Consumes: `recordClick`, `findEncoder`, `encodeClip`, `clipTarget`, `clickKind`, `compareImages`.
- Produces: `captureClickStates({ baseUrl, browser, timeout, slidevVersion, record = recordClick, encoder = findEncoder(), notice = console.warn }) -> { canvas, slides }`. Each slide is one of:
  - `{ index, sourceIndex, filepath, notes, transition, states: Scene[], clicks: Click[] }`, `clicks[k]` for `k >= 1`, `Click = { kind: "reveal" } | { kind: "motion", target, drop: Node[], clip: { data, poster, duration, natural } }`.
  - the existing `{ index, sourceIndex, filepath, error }`.

- [ ] **Step 1: Write the failing tests in `test/render/capture.test.mjs`**

Extend the file's existing stub server so `/<n>?clicks=k` serves per-click HTML and `window.__slidev__.nav` has `clicksTotal`.

```js
test("every click state is captured, and neighbouring slides stay out", opts, async (t) => {
  const { slides } = await captureStates(t, TWO_CLICKS);
  assert.equal(slides[0].states.length, 3);
  assert.deepEqual(slides[0].states.map((s) => s.nodes.length), [1, 2, 3]);
});

test("the transition comes from the slide, else the deck", opts, async (t) => {
  const { slides } = await captureStates(t, TRANSITIONS);
  assert.deepEqual(slides.map((s) => s.transition), ["slide-left | fade", "fade"]);
});

test("in this mode a large picture is JPEG; a small one stays PNG; the default capture keeps PNG", opts, async (t) => {
  const { slides } = await captureStates(t, BIG_AND_SMALL_PICTURE);
  const [big, small] = slides[0].states[0].nodes.filter((n) => n.kind === "raster");
  assert.equal(big.png.subarray(0, 2).toString("hex"), "ffd8");
  assert.equal(big.pictureExt, "jpg");
  assert.equal(small.png.subarray(0, 4).toString("hex"), "89504e47");
  const plain = await captureStub(t, BIG_AND_SMALL_PICTURE);
  assert.equal(plain.slides[0].nodes.find((n) => n.kind === "raster").png.subarray(0, 4).toString("hex"), "89504e47");
});

test("no recording happens when no picture changed", opts, async (t) => {
  let recorded = 0;
  await captureStates(t, TWO_CLICKS, { record: async () => { recorded++; } });
  assert.equal(recorded, 0);
});

for (const [name, over, reason] of [
  ["with no encoder", { encoder: null }, /no encoder/],
  ["that never settles", { encoder: FAKE_ENCODER, settled: false }, /never settled/],
]) {
  test(`a click ${name} becomes a reveal with a notice, and its frames are removed`, opts, async (t) => {
    const lines = [];
    const dir = fakeFrames();
    const { encoder, settled = true } = over;
    const { slides } = await captureStates(t, PICTURE_CHANGES, { encoder, record: async () => ({ dir, count: 2, fps: 30, motionFrames: 40, settled, changedText: [], changedVisibly: true }), notice: (l) => lines.push(l) });
    assert.equal(slides[0].clicks[1].kind, "reveal");
    assert.match(lines[0], new RegExp(`^Notice: slide 1 click 1: .*${reason.source}`));
    assert.equal(existsSync(dir), false);
  });
}

test("a hidden text that changes during motion is re-recorded in the clip and dropped from the native layer", opts, async (t) => {
  const calls = [];
  const { slides } = await captureStates(t, PICTURE_AND_CAPTION, { encoder: FAKE_ENCODER, record: fakeRecordReportingChange(calls, ".caption") });
  assert.equal(calls.length, 2);
  assert.deepEqual(calls[1].show, [".caption"]);
  assert.equal(slides[0].clicks[1].drop.length, 1);
});

test("a clip's last frame unlike the picture is kept, with a notice", opts, async (t) => {
  const lines = [];
  const { slides } = await captureStates(t, PICTURE_CHANGES, { encoder: FAKE_ENCODER, record: fakeRecordWithOtherLastFrame, notice: (l) => lines.push(l) });
  assert.equal(slides[0].clicks[1].kind, "motion");
  assert.match(lines.join("\n"), /^Notice: slide 1 click 1: the clip's last frame differs/m);
});

test("a slide whose click states cannot be captured exports its final state, with a notice", opts, async (t) => {
  const lines = [];
  const { slides } = await captureStates(t, CLICK_STATE_FAILS, { notice: (l) => lines.push(l) });
  assert.equal(slides[0].states.length, 1);
  assert.deepEqual(slides[0].clicks, []);
  assert.match(lines[0], /^Notice: slide 1: /);
});
```

Add a scene test: a raster node with `pictureExt: "jpg"` gives an image shape with `ext: "jpg"`; without it, `ext: "png"`; a node with `effects` passes them to each of its shapes.

- [ ] **Step 2: Run them to verify they fail**

Run: `node --test test/render/capture.test.mjs test/render-scene.test.mjs`
Expected: FAIL, `captureClickStates` is not exported.

- [ ] **Step 3: Implement the capture changes**

`captureSlide` takes a final `{ url, jpeg, paths } = {}`; with none of them set it behaves exactly as today. `elementShot` with `jpeg` uses `{ type: "jpeg", quality: 85 }` and sets `node.pictureExt = "jpg"` when the part shot covers at least half the canvas. `captureClickStates` sets up its page as `captureDeck` does and first opens `/1?print` and waits for the deck, because the first page of a fresh dev server can reload while Vite prepares its dependencies (a spike got a nearly empty first capture without this). It then opens `/<n>?clicks=0` for `clicksTotal` and captures each state, with `{ url, jpeg: true, paths: true }`. For each `k >= 1` with a `clipTarget` it calls `record` with `keep` set to the target's `elPath` and `hide` set to the `elPath` of every node after the target in state `k`, plus every node after the largest raster in state `k-1`, since native content that exists only before the click must stay out of the clip too. If it reports `changedText`, it records again with those as `show` and puts the matching nodes in `drop`. Then it applies `clickKind`, encodes and removes `dir` in a `finally`. The poster is frame 0; `duration` is `count / fps`; `natural` is `clipViewport(canvas)`. Compare the last frame with `target.png` through `compareImages` only when the target covers at least half the canvas: below similarity 0.95, print `Notice: slide N click K: the clip's last frame differs from the slide's picture (the page's live state differs from a fresh load, or a recording is wrong).` For a motion click, print `Notice: slide N click K: clip, S s, M MB.`

- [ ] **Step 4: Implement the scene changes**

In `shapesOf`: `ext: node.pictureExt || "png"` for rasters. In `sceneToSpec`: after `shapesOf(node, ...)`, set `node.effects` on each resulting shape when present.

- [ ] **Step 5: Run the tests**

Run: `npm test && npm run test:render`
Expected: PASS, with existing capture and scene tests unchanged.

- [ ] **Step 6: Commit**

Message: `Render mode: capture every click state, and turn a click with script motion into a clip`.

### Task 8: Animated specs, the render flow and the flag

**Files:**
- Create: `src/render/animated.mjs`
- Modify: `src/render/index.mjs`, `src/index.mjs`, `bin/cli.mjs`
- Test: `test/animated.test.mjs`, `test/render-index.test.mjs`

**Interfaces:**
- Consumes: `planSlide`, `mergeRuns`, `effectsOf`, `sceneToSpec`, `captureClickStates`, Task 2's spec fields.
- Produces:
  - `slidevTransition(value) -> "fade" | { push } | null`.
  - `slideSpecs(slide, sceneStyle, { notice }) -> SlideSpec[]`.
  - `renderSpecs` with `config.render.animations` set: uses `deps.captureStates ?? captureClickStates` and returns `masterShapes: []`.

- [ ] **Step 1: Write the failing tests in `test/animated.test.mjs`**

```js
test("transitions: the forward half maps; anything else is null", () => {
  assert.equal(slidevTransition("fade"), "fade");
  assert.equal(slidevTransition("fade-out"), "fade");
  assert.deepEqual(slidevTransition("slide-left | fade"), { push: "l" });
  assert.deepEqual(slidevTransition("slide-up"), { push: "u" });
  assert.equal(slidevTransition("view-transition"), null);
  assert.equal(slidevTransition(undefined), null);
});

// Helpers: slideOf(states, clicks = [], meta = {}) builds a captured slide; title, text, table, box and raster build nodes; `fail` throws if called.
test("a slide with no clicks is one spec with no effects, not marked animated", () => {
  const specs = slideSpecs(slideOf([[title("T"), box("a")]]), STYLE, { notice: fail });
  assert.equal(specs.length, 1);
  assert.ok(!specs[0].animated);
  assert.ok(specs[0].shapes.every((s) => !s.effects));
});

test("reveals stay on one spec: a table per state, swapped on the click", () => {
  const states = [[title("T"), table("a")], [title("T"), table("b")]];
  const [spec, ...rest] = slideSpecs(slideOf(states, [null, { kind: "reveal" }]), STYLE, { notice: fail });
  assert.equal(rest.length, 0);
  assert.ok(spec.animated);
  assert.deepEqual(spec.shapes.filter((s) => s.kind === "table").map((s) => s.effects), [[{ click: 1, dir: "out" }], [{ click: 1, dir: "in" }]]);
  assert.ok(spec.placeholders.title, "a title that stays stays a placeholder");
});

test("a title that changes on a reveal becomes plain text shapes", () => {
  const states = [[title("One")], [title("Two")]];
  const [spec] = slideSpecs(slideOf(states, [null, { kind: "reveal" }]), STYLE, { notice: fail });
  assert.ok(!spec.placeholders.title);
  assert.equal(spec.shapes.filter((s) => s.effects?.length).length, 2);
});

test("a motion click: a new spec, the clip where the picture was, native content above, no transition", () => {
  const target = raster("map");
  const states = [[raster("old"), title("T")], [target, title("T")]];
  const clicks = [null, { kind: "motion", target, drop: [], clip: CLIP }];
  const specs = slideSpecs(slideOf(states, clicks, { transition: "fade", notes: "N" }), STYLE, { notice: fail });
  assert.equal(specs.length, 2);
  assert.equal(specs[0].transition, "fade");
  assert.equal(specs[1].transition, undefined);
  assert.equal(specs[1].shapes[0].kind, "video");
  assert.equal(specs[1].shapes[0].autoplay, true);
  assert.equal(specs[1].shapes[0].muted, true);
  assert.equal(specs[1].shapes[0].loop, false);
  assert.equal(specs[1].shapes[0].poster.ext, "jpg");
  assert.deepEqual(specs.map((s) => s.notes), ["N", "N"]);
});

test("a text dropped for changing during motion is not on the clip's slide", () => {
  const target = raster("map");
  const caption = text("Searching");
  const clicks = [null, { kind: "motion", target, drop: [caption], clip: CLIP }];
  const specs = slideSpecs(slideOf([[raster("old")], [target, caption]], clicks), STYLE, { notice: fail });
  assert.ok(specs[1].shapes.every((s) => s.kind !== "text"));
});

test("a deck's own video on a slide with clicks stays a video with its own autoplay and loop", () => {
  const own = { kind: "video", rect: RECT, natural: { w: 16, h: 9 }, data: MP4, ext: "mp4", posterImage: POSTER, autoplay: true, loop: true, muted: true };
  const states = [[own], [own, box("a")]];
  const [spec] = slideSpecs(slideOf(states, [null, { kind: "reveal" }]), STYLE, { notice: fail });
  const video = spec.shapes.find((s) => s.kind === "video");
  assert.equal(video.loop, true);
  assert.ok(!video.effects?.length);
});

test("an unknown transition gets a notice and no transition", () => {
  const lines = [];
  const [spec] = slideSpecs(slideOf([[box("a")]], [], { transition: "view-transition" }), STYLE, { notice: (l) => lines.push(l) });
  assert.equal(spec.transition, undefined);
  assert.match(lines[0], /^Notice: slide 1: transition "view-transition"/);
});
```

Add to `test/render-index.test.mjs`, using its `setup`:

```js
test("with animations on: the states capture is used, nothing goes on the master, and videos stay native", async () => { /* deps.captureStates returns two states; assert specs[0].animated, masterShapes is [], deps.gif never called */ });
test("with animations off: the states capture is never called", async () => { /* deps.captureStates throws if called */ });
test("a slide that failed still falls back to its markdown", async () => { /* a slide with error in the states capture */ });
```

- [ ] **Step 2: Run them to verify they fail**

Run: `node --test test/animated.test.mjs test/render-index.test.mjs`
Expected: FAIL, module not found and new cases failing.

- [ ] **Step 3: Implement `src/render/animated.mjs`**

For each segment: merge the runs, set `node.effects = effectsOf(run, segment)` on a copy of each run's node, delete `role` from a title whose run does not span the segment, replace the motion click's `target` run with a full-canvas video node (`fit: "fill"`, `natural: clip.natural`, `posterImage: { data: clip.poster, ext: "jpg" }`, `autoplay: true`, `loop: false`, `muted: true`), remove `drop` nodes, and call `sceneToSpec`. Set `spec.animated` when a segment has more than one state or a clip, `spec.transition` on the first segment only, and `spec.notes` on all.

- [ ] **Step 4: Wire up the flow**

`renderSpecs`: when `config.render?.animations`, capture with the states capture, skip `animateVideos`, keep `resolveChart` on every state's nodes, skip the master computation, and build specs per slide with `slideSpecs`, or `markdownFallback` for `error`. The default branch stays byte-for-byte as it is. `src/index.mjs`: `options.render?.animations` or `config.render?.animations` implies render mode and sets `config.render.videos = "native"`. `bin/cli.mjs`: parse `--animations`, set `config.render = { ...config.render, animations: true }`, and document it in the header comment.

- [ ] **Step 5: Run the tests**

Run: `npm test && npm run test:render`
Expected: PASS.

- [ ] **Step 6: Commit**

Message: `Render mode: --animations exports click reveals, clips and slide transitions`.

### Task 9: Fixture deck, end-to-end test and midway checkpoint

**Files:**
- Create: `test/fixtures/animated-deck/package.json`, `package-lock.json`, `slides.md`, `components/Stepper.vue`, `components/Mover.vue`
- Create: `test/render/animations.test.mjs`

**Interfaces:**
- Consumes: the whole feature through `convert`.
- Produces: `npm --prefix test/fixtures/animated-deck ci`, documented for Task 10's README.

- [ ] **Step 1: Create the fixture deck**

`@slidev/cli` pinned to `53.0.0` in the fixture's own `package.json`, with `headmatter transition: fade`, `theme: none`, and no network resources. Slide 1: two `v-click` blocks. Slide 2: `clicks: 2` with `Stepper`, a table that gains a column per click through `$clicks`, and a caption that changes per click. Slide 3: `clicks: 1` with `Mover`, a full-slide inline SVG redrawn by `requestAnimationFrame` for 1 s after the click.

- [ ] **Step 2: Install it**

Run: `npm --prefix test/fixtures/animated-deck ci`
Expected: the fixture's Slidev installed; `git status` shows no `node_modules`.

- [ ] **Step 3: Write `test/render/animations.test.mjs`**

```js
const FIXTURE = join(ROOT, "test/fixtures/animated-deck");
const ready = existsSync(join(FIXTURE, "node_modules/@slidev/cli"));
const opts = { skip: !ready ? "run npm --prefix test/fixtures/animated-deck ci" : !chromium ? "Playwright/Chromium not available" : false, timeout: 300000 };

test("the fixture exports in Slidev's click order", opts, async (t) => {
  const { buffer, z } = await exportFixture(t);
  assert.deepEqual(await lintPackage(buffer), []);
  const slides = slideTexts(z);
  assert.equal(slides.length, findEncoder() ? 4 : 3);
  assert.deepEqual(slides.slice(0, 3).map(clickSteps), [2, 2, findEncoder() ? 0 : 1]);
  for (const s of slides) assert.doesNotMatch(s, /3DA1BC7CA4C3/);
});

test("slide 1's v-clicks and slide 2's table each step one at a time", opts, async (t) => {
  const { z } = await exportFixture(t);
  const [one, two] = slideTexts(z);
  assert.ok(entrances(one) >= 2);
  assert.equal(tables(two), 3);
  assert.equal(entrances(two), 4);
  assert.equal(exits(two), 4);
  assert.match(one, /<\/p:clrMapOvr><p:transition><p:fade\/><\/p:transition>/);
});

test("the motion click is a clip that starts by itself, with no click step and no transition", { ...opts, skip: opts.skip || (findEncoder() ? false : "no encoder installed") }, async (t) => {
  const { z } = await exportFixture(t);
  const clip = slideTexts(z)[3];
  assert.match(clip, /nodeType="afterEffect"/);
  assert.doesNotMatch(clip, /clickEffect|<p:transition/);
  const mp4 = z.names.find((n) => n.endsWith(".mp4"));
  assert.equal(z.buffer(mp4).subarray(4, 8).toString(), "ftyp");
});

test("without the flag the same deck has one slide each and no effects", opts, async (t) => {
  const { z } = await exportFixture(t, { animations: false });
  const slides = slideTexts(z);
  assert.equal(slides.length, 3);
  for (const s of slides) assert.doesNotMatch(s, /<p:timing|<p:transition/);
});
```

- [ ] **Step 4: Run it**

Run: `node --test test/render/animations.test.mjs`
Expected: PASS. If it cannot have failed first, because earlier tasks made it pass, break `planSlide` temporarily to see it fail, then restore it.

- [ ] **Step 5: Commit**

Message: `Tests: an animated fixture deck, exported end to end`.

- [ ] **Step 6: Midway checkpoint (needs the user)**

From the session that holds the conversation: export a copy of the test deck trimmed to its table slide and its route search slide with `--animations`, ask the user to upload it in the browser pane, and step through with the keyboard. Expected: the table steps one column per click with no flash, and the clip starts by itself. Fix what fails, with a test, before Task 10.

### Task 10: README, full deck on the web, and the PR

**Files:**
- Modify: `README.md`, `bin/cli.mjs` (usage line if not done), `package.json` (version only if the maintainer's pattern calls for it)

**Interfaces:**
- Consumes: the finished feature.
- Produces: the open PR.

- [ ] **Step 1: Write the README section "Animations (`--animations`)"**

Cover: the flag and that it implies render mode; what becomes what; that it targets PowerPoint for the web, and that desktop PowerPoint and Impress should open the file but were not the target; to use a keyboard or clicker, because a mouse click on a full-slide clip restarts it; the brief blank at a motion click; the slide count; the need for `ffmpeg` or GStreamer with H.264, and the Fade fallback; what cannot carry over (hover highlights, live sliders, table scrolling, motion with no click); the risks from the spec, in plain words; that the default export stays the editable copy. Change the line near `README.md:144` that says animations and transitions are not carried over. Add `npm --prefix test/fixtures/animated-deck ci` to the development section.

- [ ] **Step 2: Check the default export is unchanged**

Export the test deck with `--render` from `render-map-slides` plus the Task 1 fix, and from this branch. Compare every unzipped part.
Expected: identical, apart from any timestamp the writer already varies. If a map picture differs between two runs, run the same export twice on one build to see whether that is only tile-loading noise. Record the command and result for the PR.

- [ ] **Step 3: Check the full deck in LibreOffice**

Export the test deck with `--animations`. Convert to `.odp` and PDF with `soffice --headless`.
Expected: both convert; the number of `anim:` effects per page matches the plan.

- [ ] **Step 4: Run the final web check (needs the user)**

From the session that holds the conversation: ask the user to upload the export in the browser pane. Step through all 15 clicks with the keyboard, compare each with Slidev's state for that click, and read the Animation Pane on any slide that misbehaves. Check the clips on slides 2, 4, 7 and 10, the blank, and going back with the left arrow. Ask the user to repeat the show in their usual browser. Record the file size.

- [ ] **Step 5: Run everything**

Run: `npm test && npm run test:render`
Expected: PASS, with the fixture tests having run, not skipped.

- [ ] **Step 6: Check nothing names the deck**

Run: `git log origin/render-map-slides..HEAD --format=%B | grep -i -E '<names of the test deck and its organisations>'; git diff origin/render-map-slides --stat`, with the names supplied at run time and not written into the repo.
Expected: no matches.

- [ ] **Step 7: Commit, push and open the PR**

Base it on `render-map-slides` if Eiriksb/slidev-pptx-export#2 is still open, else `main`. The body says plainly what was checked on the web, what was checked in LibreOffice, and what was not checked: desktop PowerPoint at least.
