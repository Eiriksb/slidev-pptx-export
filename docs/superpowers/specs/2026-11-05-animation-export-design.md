# Animation export for render mode

Status: approved, then changed by the early web checkpoint: see "Clips hold everything". Date: 2026-11-05.

## Intent

Render mode runs a deck in headless Chromium and converts what it draws, but it captures each slide in its
final click state, with no animation. A deck whose story unfolds click by click loses that story in the
`.pptx`. This feature adds an opt-in export that plays the deck's slide transitions, click reveals and
script-driven motion in the slide show, in the same click order as Slidev.

The person presenting uses **PowerPoint for the web**, in a browser, and advances with a **keyboard or
clicker**. Those two facts shaped this design. Several things that work in desktop PowerPoint do not work on
the web, so the design follows what was measured there rather than what the file format allows.

## Who decided what

- **Decided by the user:** presenting from PowerPoint for the web; advancing with a keyboard or clicker; a brief
  blank before the clip at each motion click is acceptable; a few test uploads during the build; testing a start
  delay on clips before giving up native text above them, with everything going into the clip if that failed.
- **Set by the task:** opt-in (`--animations` / `render: { animations: true }`); the default export stays as it
  is; the test deck steps through its 15 clicks in Slidev's order; new tests first; `npm test`,
  `npm run test:render` and the package lint pass; README updated; PR stacked on `render-map-slides` while that
  PR is open.
- **The author's calls, open to correction:** everything under "Behaviour" that is not a measured fact,
  including that the animated file is a copy for presenting while the default export stays the editable one.

## What was measured

Test files were run as real slide shows in PowerPoint for the web, in the Claude desktop app's browser pane
(Chromium 152 on Linux), signed in to a Microsoft 365 work account. The Animation Pane shows how the web parsed
a slide's timing, which makes it the quickest way to check a timeline.

| Behaviour on the web | Result |
|---|---|
| Video with the correct media extension id, set to start by itself | Plays on slide entry, every time the generated XML was used |
| That video when it ends | Holds its last frame |
| Native text and shapes above a video | Above the first clip of a show they stay above it, sharp, including a caption fading in during playback. In a full run from slide 1, on later clip slides with their own files, every native shape vanished, including those due on a key press, while the clip played on. Three runs agree: the second clip slide lost them in two of them and the third in the other, and static slides between were fine. A 1 s start delay on the clip did not help. Not explained: five clip slides sharing one file kept them, and starting the show at a failing slide kept them |
| Key press during a playing video | Moves on |
| On Click video in the main sequence, key press | Step is consumed, the video does not start, even when the web's own UI writes it |
| Mouse click on a video | Starts it if not started, restarts it if playing; never advances |
| Video that is `withEffect` in a click | Starts on slide entry, not on the click |
| Native click effects on one slide: On Click, With Previous, Appear, Fade, Disappear | Work, across several clicks, on tables, text and pictures |
| The map during those click effects | Does not flash |
| Changing slide | The map blanks briefly and draws last; a JPEG map (about 500 KB) appeared within about one screenshot, the PNG (about 4 MB) took visibly longer. If the slide before shows the same JPEG, as a clip's poster does, the map did not blank at all and only the text redrew |
| Animated GIF, including a one-shot GIF | Stays on its first frame |
| 40 timed full-slide pictures on one slide | The slide show hung and needed a reload |
| Morph | Box ended in the right place; the motion itself could not be sampled by screenshot |
| The finished export of the test deck, run from slide 1 with the Right arrow | All 24 key presses (9 slide changes, 15 clicks) gave Slidev's states in order; all 7 clips started by themselves and held their last frame; no native shape vanished |
| Going back with the Left arrow | Works; each clip slide takes one extra press, as the first goes back through the clip's own step. Reveals undo one per press |
| Video colours | GStreamer's default conversion plays brighter than the pictures around it. In real Chrome, whole-frame means against a source of 201.6, 203.8, 193.6: default 213.5, 215.6, 205.3; `colorimetry=bt709` 200.7, 205.3, 194.1; ffmpeg's plain `yuv420p` 201.6, 203, 193.6 |

Found along the way:

- **Bug in the current code.** `MEDIA_EXT` in `src/ooxml/shapes.mjs` is `{DAA4B4D4-6D71-4841-9C94-3DA1BC7CA4C3}`.
  PowerPoint writes `{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}`. With the wrong id the web shows no play button,
  no media bar and a greyed-out Trim Video, and the video never plays. This is on `main` and on
  `render-map-slides`, and it affects the default export too.
- **Each animation on a text shape needs its own `grpId`** (0, 1, ...) and its own `p:bldP` entry. With both on
  `grpId="0"` the web drops the second one. If the dropped effect was the first in a click, that click's other
  effects merge into the previous click.
- PowerPoint for the web can insert video itself and its Present menu offers Presenter View, so Microsoft's
  pages saying otherwise are out of date.

Not explained, so treat as risks:

- A stock video inserted by the web did not start by itself in the one run tried, though it played from its
  own control.
- After the clip's Start setting was changed in the web UI to On Click and back, that clip no longer started
  by itself. The generated, untouched XML kept working in the same session.
- Pressing Escape during playback can leave the clip's frame over the editor until the page is reloaded.

Not checked at all: the user's usual browser, desktop PowerPoint and LibreOffice Impress with the corrected id,
and Presenter View.

## Goals

- `--animations` / `render: { animations: true }` produces a `.pptx` that, in a PowerPoint for the web slide
  show driven by keyboard or clicker, steps through the deck's clicks in Slidev's order.
- Works for both `v-click` and `$clicks`-driven components.
- Still slides keep text and tables native and sharp. Clip slides hold everything in the clip, because native shapes
  above a clip did not survive a full run on the web.

## Non-goals

- Motion that runs when a slide opens, with no click. The slide shows its settled state, as it does today.
- Hover highlights, live sliders and table scrolling. These cannot carry over; the README says so.
- Making the animated file pleasant to edit by hand. The default export is the editable copy.
- Morph.
- Advancing with mouse clicks over a clip. The README says to use a keyboard or clicker.
- Tuning for desktop PowerPoint or Impress. The file should open there, but they are not the target.

## Anti-goals

- Any change to default-export output other than the `MEDIA_EXT` fix.
- A dependency that downloads and runs a binary at install time.
- Naming the test deck or its organisations anywhere in the repo, commits or PRs. The repo looks public.
- Silently exporting something misleading. Anything that could not be animated gets a `Notice:` line, as the
  rest of render mode does.

## Behaviour

### The flag

`--animations` on the CLI, `render: { animations: true }` in config. It implies render mode. It also implies
`render.videos: "native"`, because the GIF path in `src/render/animate.mjs` would turn clips and the deck's own
videos into GIFs, which do not animate on the web. With the flag off, the code paths and output are unchanged.

### Planning a slide

For each Slidev slide, capture every click state `k` from 0 to the slide's click total, using
`/<n>?clicks=k` without `print` (`?print&clicks=1` ignores `clicks`). Scope every selector to
`.slidev-page-<n>`, because Slidev also mounts neighbouring slides. A throwaway spike confirmed the existing
capture and extraction work on these pages, including the wait for scripts to settle.

Each click `k` is then one of two kinds:

- **Reveal:** the change has no script-driven motion. It becomes native click effects on the current still slide.
- **Motion:** scripts animate for a while after the click and a picture changes. It becomes a PowerPoint slide
  holding only a clip of the whole slide, which starts by itself, because that is the only start that works from a
  keyboard on the web.

So a Slidev slide becomes a run of PowerPoint slides of two sorts:

- A **still slide** shows one or more states natively, its later states as On Click effects. The first one opens in
  state 0. One that follows a clip opens in the state of the reveal that follows the clip, because the key press that
  reached it was that reveal. It uses a Fade transition in place of that reveal's fade.
- A **clip slide** holds the clip and nothing else, and holds its last frame until the next key press. It is state `k`
  until then. It has no state of its own afterwards: the next click goes to a new slide.

### Clips hold everything

The approved design kept text and tables native above the clip. On the web that did not hold in a full run from the
first slide: every native shape on later clip slides vanished, while the clip played on (see "What was measured"). A
1 s start delay did not help. The user chose to test the delay first, on the stated basis that if it failed, a clip
slide would hold everything.

What that costs: on a clip slide, text and tables are pixels in a video, a little softer than native text, and they
are fixed if someone edits the still slides. What it gives: there is nothing left to vanish, the recording needs no
hide lists, and a text that changes during motion simply plays. Key presses still match Slidev one for one, and every
still slide stays native.

Motion detection (the builder may refine the rule, and the test deck is the reference): after the click, script
motion means page scripts keep requesting animation frames for at least about 0.4 s, and a picture node differs
between states `k-1` and `k`. CSS transitions alone are a reveal.

On the test deck, by reading its code and not yet measured, the motion clicks are slide 2 click 1 (a route being
drawn), slide 4 click 1 (a map flight), slide 7 click 1 (a route search of about 10.7 s) and slide 10 click 1
(another route search).

### Reveals: diffing click states

Compare the node lists of states `k-1` and `k`. A node in both stays. A node only in `k` enters on click `k`. A
node only in `k-1` leaves on click `k`. Everything for one click plays together: the first effect is On Click,
the rest With Previous.

- **Identity.** `nodeSignature` in `src/render/index.mjs` is the starting point, but it returns null for tables,
  lists, code, charts, videos and titles. This needs a signature for every kind, built from the node's content
  and position with bytes hashed. A spike saw byte-identical map screenshots across page loads for slides whose
  map did not change, so hashing pictures is viable. If noise makes a picture look changed, the result is a
  harmless cross-fade.
- **Effect.** Fade, 400 ms, in and out. Both Fade and Appear were measured working. Fade is closest to the CSS
  transitions these decks use.
- **Stacking.** All states of a PowerPoint slide sit on it at once, in one z-order merged from each state's DOM
  order so that each state still paints correctly. If two states disagree on order, repeat a node as a separate
  shape rather than fail. A table that gains a column is one native table per state, swapped on the click.
- **`box` nodes** can expand to several shapes through `boxShapes`; all of them get the same effect.
- **Titles.** A title placeholder cannot be duplicated on a slide. If the title changes on a reveal, the later
  ones are plain text shapes.

### Motion: recording clips

Record state `k-1` to `k` as H.264 MP4 and make it the only shape on its own PowerPoint slide, filling the slide.

- **Recording.** Load `/<n>?clicks=k-1`, let it settle, advance with
  `window.__slidev__.nav.next()`, and step a virtual clock one frame at a time until the page settles. Real-time
  capture is not good enough: Chromium's screencast managed about 9.5 fps on the map flight with gaps of about
  0.5 s, while stepping gave an exact 30 fps. The clock overrides `requestAnimationFrame`, `performance.now`,
  `Date` and timers, and advances CSS animations through `document.getAnimations()`. The spike's version recorded
  342 frames in 46 s. Wait briefly for images still loading before taking a frame, so map tiles appear.
- **Size.** 1920x1080 at 30 fps, with the slide scaled up by Slidev. CRF 23 gave about 0.8 MB for the 3 s flight
  and 2.8 MB for the 11.4 s search.
- **What is in the clip.** The whole slide as the page draws it: pictures, text, tables and global layers. Nothing is
  hidden, and nothing is drawn natively above it. A text that changes while the page moves, such as the test deck's
  slide 10 caption, needs no special handling.
- **Poster.** The clip's first frame, which is the state before the click. A spike confirmed the map lines up pixel
  for pixel with the picture on the slide before it (best match at zero offset), and only the text redraws, as
  Chromium's rendering gives way to PowerPoint's. On the test deck's flight the map jumps to full width on the click,
  as it does in Slidev.
- **Which state is the reference.** Still states come from a direct load of `/<n>?clicks=k`, as the default export
  does. A clip is what the page does when it is clicked. The two can legitimately differ, so neither is checked
  against the other as a condition for keeping the clip.
- **Guard.** Drop the clip, making the click a Fade reveal with a `Notice:` saying why, only if the page never
  settles, no encoder is available, or the recording shows no visible change. If the clip's last frame differs from
  a screenshot of the direct load beyond a threshold, keep the clip and print a `Notice:` naming the slide and click,
  so a wrong recording is visible without costing a good one.
- **Known quirk.** The test deck's slide 4 leaves its map a stale size after the flight, because its
  `ResizeObserver` skips while the map is moving: a blank band beside the map, and the map off-centre. Slidev does the
  same live, so the clip records it faithfully and prints the notice above. The next state, captured directly, does
  not have it, so the map visibly corrects itself on that click. This is the deck's bug and is not worked around
  here. Fixing it in the deck removes both the band and the notice.

### The clip's timing

Exactly the form below, as PowerPoint for the web writes it and as measured working: a `p:seq` main sequence with
one `p:par` that starts on `onBegin` of the sequence, holding a `mediacall` with `nodeType="afterEffect"`,
`cmd="playFrom(0.0)"` and `dur="1"`; plus `p:video` with `p:cMediaNode mute="1"` and no `repeatCount`. Clips are
muted. Do not add the `togglePause` interactive sequence: it was not needed for starting, and it only serves mouse
clicks. A clip slide has no other steps in its sequence.

### Slide transitions

Map the deck's `transition`, from the slide's frontmatter or the headmatter default, taking the forward half of
`a | b`: `fade` and `fade-out` to `<p:transition><p:fade/></p:transition>`, `slide-*` to `p:push` with the matching
direction. Anything else, such as `view-transition` or a custom transition, gets no transition and a `Notice:`.
Only `fade` was measured on the web; Microsoft's service description says only Fade and Wipe play there. The new
slide at a motion click gets no transition, and its poster lines up with the slide before it. On the web the map did
not blank at all there, and only the text redraws. A still slide that follows a clip opens with `p:fade`, standing in
for the 400 ms fade of the reveal that opened it.

### Pictures

Screenshot pictures that cover most of the slide become JPEG in this mode only: about 500 KB against about 4 MB,
and they appear faster after a slide change. Screenshots are opaque, so nothing is lost. Small pictures and logos
stay PNG, and SVG-backed images keep their transparency. The default export keeps PNG.

### Encoder

Look for `ffmpeg` with `libx264` on PATH, as `src/render/animate.mjs` already does, then GStreamer's `gst-launch-1.0`
with `x264enc` and `mp4mux`. `h264parse` may be missing, and is not needed if `x264enc` is asked for
`stream-format=avc,alignment=au`. GStreamer must also be told `colorimetry=bt709` after `videoconvert`, or the clip
plays brighter than the pictures around it. With neither tool, motion clicks become Fade reveals with a `Notice:`
naming what to install. Playwright's bundled ffmpeg writes only VP8 WebM, so it cannot be used. Nothing is added to
`dependencies`.

### Notes and failures

Speaker notes repeat on every PowerPoint slide of a Slidev slide, so the presenter always has them. A slide whose
click states cannot be captured exports as it does today, in its final state, with a `Notice:`.

## XML rules that must not regress

- The media extension id is `{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}`, asserted exactly by a test.
- Each animation on a text shape has its own `grpId` and `p:bldP`; a shape with a fill also has `animBg="1"`.
- `p:transition` comes after `p:clrMapOvr` and before `p:timing`.
- `cTn` ids are unique within a slide, and every `spTgt` names a shape on that slide.

## Worked example: a table that gains columns

A slide with `clicks: 3`, no script motion and a table that gains a column per click stays one PowerPoint slide,
holding four tables and four captions. Click 1 fades out table 0 and caption 0 and fades in table 1 and caption 1,
and so on. The map picture is in all four states, so it stays and never flashes. This was run on the web and worked.

## Worked example: a motion click with reveals after it

A slide with `clicks: 3` whose click 1 is a route being drawn, and whose clicks 2 and 3 each add a card, becomes three
PowerPoint slides. The first shows state 0 natively. The second holds only the clip of state 0 to state 1, with the
page's text and table inside it, and holds its last frame. Click 2 moves to the third slide, which opens natively in
state 2 with a Fade transition. Click 3 fades in the second card on that slide. Three key presses, as in Slidev.

## Verification

- **Unit tests, written first:** the diff and z-order merge, the timing XML (including the rules above), the
  transition map, the slide plan, the encoder's arguments.
- **Colour:** in real Chrome, when it is installed, a clip's frames match the pictures they were made from, because
  Playwright's own Chromium cannot decode H.264.
- **Browser tests** (`test/render/`), on a small fixture deck with `v-click`, a `$clicks` component and a
  `requestAnimationFrame` animation, since the test deck uses only `$clicks`. The repo has no Slidev of its own and
  `render/server.mjs` needs one, so the fixture gets its own `package.json` under `test/fixtures/`, installed with
  `npm ci` there, and keeps Slidev out of the root `devDependencies`. The tests skip with a message saying how to
  install it, as the other render tests do when Chromium is missing, but they must have run, and been seen failing
  first, on the machine that builds this. The clock and motion detection also get stub-page tests that need no Slidev.
- **Default export unchanged:** with the flag off, output on the existing fixtures matches what `main` produces,
  except for the corrected id.
- **LibreOffice:** the file converts, and the number of effects per slide matches the plan.
- **PowerPoint for the web, the real check:** upload the test deck's export and step through all 15 clicks in the
  browser pane with the keyboard, comparing each step with Slidev's state for that click, and reading the Animation
  Pane on any slide that misbehaves. Check the risks above, the first reveal after a held clip, and the clips on
  slides 2, 4, 7 and 10. The user repeats it in their usual browser, which cannot be checked from here.
- **Report plainly** what was not checked: desktop PowerPoint at least.

## Checkpoints on the web

PowerPoint for the web can only be tested with a file the user uploads, because the browser pane has no file picker.
The user has agreed to about three uploads. Each time they upload and leave the file open, and the slide show is then
run in the pane.

1. **Early, before capture and recording are built. Done, four uploads in the end.** The results are in the table
   under "What was measured". Clips start by themselves with a key press during playback moving on, and there is no
   blank map before a clip. Native shapes above clips did not hold, so "Planning a slide" changed: a clip slide holds
   only the clip.
2. **Midway.** A real export of one multi-state slide and one motion click, made by the exporter itself.
3. **At the end.** The full test deck, as described under "Verification".

## Delivery

- The `MEDIA_EXT` fix goes first as its own small PR against `main`, with a test, because it also fixes the default
  export. It is the one exception to "flag off stays unchanged".
- The feature goes on this branch, based on `render-map-slides` (Eiriksb/slidev-pptx-export#2) while that PR is open,
  and needs the fix too.
- The README gets a section covering the flag, what becomes what, what cannot carry over, the need for a keyboard or
  clicker, the need for an encoder, and the risks above.

## Left to the builder

File layout, the z-order merge algorithm, the signature's exact form, the guard's threshold, the JPEG quality and
size cutoff, and whether detection and recording share one pass.
