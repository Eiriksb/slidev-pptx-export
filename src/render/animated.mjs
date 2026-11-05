// A captured slide's click states as SlideSpecs for an animated export (see animation-plan.mjs for what
// becomes what).

import { effectsOf, planSlide } from "./animation-plan.mjs";
import { sceneNotices, sceneToSpec } from "./scene.mjs";

const PUSH = { "slide-left": "l", "slide-right": "r", "slide-up": "u", "slide-down": "d" };

// The PowerPoint transition for a Slidev one: its forward half, as in "slide-left | fade". Null for
// anything PowerPoint has no form of.
export function slidevTransition(value) {
  const forward = String(value ?? "").split("|")[0].trim();
  if (forward === "fade" || forward === "fade-out") return "fade";
  return PUSH[forward] ? { push: PUSH[forward] } : null;
}

// Native shapes above a later clip vanish on the web, so the clip is all there is on its slide.
function clipSpec(slide, click, canvas, sceneStyle) {
  const { clip } = slide.clicks[click];
  const node = {
    kind: "video",
    rect: { x: 0, y: 0, w: canvas.w, h: canvas.h },
    fit: "fill",
    natural: clip.natural,
    data: clip.data,
    ext: "mp4",
    posterImage: { data: clip.poster, ext: "jpg" },
    autoplay: true,
    loop: false,
    muted: true,
    duration: clip.duration,
    alt: `Slide ${slide.index}, click ${click}`,
  };
  const spec = sceneToSpec({ index: slide.index, canvas, notes: slide.notes, background: null, nodes: [node] }, sceneStyle);
  spec.layout = "blank";
  return spec;
}

function stillSpec(slide, segment, sceneStyle) {
  const nodes = segment.runs.map((run) => {
    const node = { ...run.node, effects: effectsOf(run, segment) };
    // A placeholder can't be there twice, or come and go: a title that changes is plain text.
    if (node.role === "title" && (run.from > segment.first || run.to < segment.last)) delete node.role;
    return node;
  });
  return sceneToSpec({ ...slide.states[segment.first], nodes, notes: slide.notes }, sceneStyle);
}

// `slide` as captureClickStates gives it. Every spec is animated, so a deck's own video uses the timing
// the web plays too. The first spec has the deck's transition. A still slide after a clip fades in, in
// place of the fade of the reveal that opens it; a clip's poster lines up with what is before it.
export function slideSpecs(slide, sceneStyle, { notice = console.warn } = {}) {
  const said = new Set(slide.states.flatMap((state) => sceneNotices(state, slide.index)));
  for (const line of said) notice(line);
  const transition = slidevTransition(slide.transition);
  if (slide.transition && !transition && slide.transition !== "none") notice(`Notice: slide ${slide.index}: transition "${slide.transition}" has no form in PowerPoint - left out.`);

  const kinds = slide.clicks.length ? slide.clicks.map((c) => c?.kind ?? null) : [null];
  return planSlide(slide.states.map((s) => s.nodes), kinds).map((segment, i) => {
    const spec = segment.kind === "clip" ? clipSpec(slide, segment.click, slide.states[0].canvas, sceneStyle) : stillSpec(slide, segment, sceneStyle);
    spec.animated = true;
    const arrival = i === 0 ? transition : segment.kind === "still" ? "fade" : null;
    if (arrival) spec.transition = arrival;
    return spec;
  });
}
