// Timing for animated exports: native click effects, videos that start by themselves, and slide
// transitions, in the form PowerPoint for the web plays. What it needs was measured there:
//
// - A video starts by itself when its mediacall is an afterEffect in a main-sequence step that begins
//   with the sequence (onBegin). An On Click mediacall takes a key press and doesn't start the video,
//   even when the web's own UI writes it, so clips have no On Click form here.
// - Each animation on a text shape needs a grpId of its own and a p:bldP. With two on grpId 0 the web
//   drops the second, and if that one led its click, the click's other effects merge into the click before.
//
// The default export's video timing is timingXml in shapes.mjs, which this doesn't touch.

const FADE_MS = 400;

const target = (spid) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;

// The slide's timing. `videos`: the writer's [{ spid, autoplay, loop, muted }]. `effects`:
// [{ spid, kind, fill, click, dir }], with `click` numbered from 1 and `dir` "in" or "out"; those of one
// click play together, in the order given. Empty when there is nothing to time.
export function animationTimingXml({ videos = [], effects = [] }) {
  if (!videos.length && !effects.length) return "";
  let n = 1;
  const next = () => n++;
  const root = next();

  const groups = new Map();
  const builds = [];
  const groupOf = (e) => {
    if (e.kind !== "text") return "";
    const g = groups.get(e.spid) || 0;
    groups.set(e.spid, g + 1);
    builds.push(`<p:bldP spid="${e.spid}" grpId="${g}"${e.fill ? ' animBg="1"' : ""}/>`);
    return ` grpId="${g}"`;
  };

  const visibility = (spid, value, delay) => {
    const id = next();
    return `<p:set><p:cBhvr><p:cTn id="${id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${target(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${value}"/></p:to></p:set>`;
  };
  const fade = (spid, direction) => {
    const id = next();
    return `<p:animEffect transition="${direction}" filter="fade"><p:cBhvr><p:cTn id="${id}" dur="${FADE_MS}"/>${target(spid)}</p:cBhvr></p:animEffect>`;
  };
  const effectXml = (e, nodeType) => {
    const id = next();
    const body = e.dir === "out" ? fade(e.spid, "out") + visibility(e.spid, "hidden", FADE_MS - 1) : visibility(e.spid, "visible", 0) + fade(e.spid, "in");
    return (
      `<p:par><p:cTn id="${id}" presetID="10" presetClass="${e.dir === "out" ? "exit" : "entr"}" presetSubtype="0" fill="hold"${groupOf(e)} nodeType="${nodeType}">` +
      `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`
    );
  };
  const playXml = (v, nodeType) => {
    const id = next();
    const call = next();
    return (
      `<p:par><p:cTn id="${id}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="${nodeType}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
      `<p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${call}" dur="1" fill="hold"/>${target(v.spid)}</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`
    );
  };
  const step = (conditions, inner) => {
    const outer = next();
    const held = next();
    return (
      `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst>${conditions}</p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${held}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner()}</p:childTnLst></p:cTn></p:par>` +
      "</p:childTnLst></p:cTn></p:par>"
    );
  };

  const auto = videos.filter((v) => v.autoplay);
  const clicks = [...new Set(effects.map((e) => e.click))].sort((a, b) => a - b);
  let seq = "";
  if (auto.length || clicks.length) {
    const main = next();
    const steps = [];
    if (auto.length) {
      steps.push(step(`<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="${main}"/></p:cond>`, () => auto.map((v, i) => playXml(v, i ? "withEffect" : "afterEffect")).join("")));
    }
    for (const click of clicks) {
      const list = effects.filter((e) => e.click === click);
      steps.push(step('<p:cond delay="indefinite"/>', () => list.map((e, i) => effectXml(e, i ? "withEffect" : "clickEffect")).join("")));
    }
    seq =
      `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${main}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${steps.join("")}</p:childTnLst></p:cTn>` +
      '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
      '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>';
  }
  const media = videos
    .map((v) => {
      const id = next();
      return (
        `<p:video><p:cMediaNode${v.muted ? ' mute="1"' : ""}><p:cTn id="${id}"${v.loop ? ' repeatCount="indefinite"' : ""} fill="hold" display="0">` +
        `<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>${target(v.spid)}</p:cMediaNode></p:video>`
      );
    })
    .join("");
  const buildList = builds.length ? `<p:bldLst>${builds.join("")}</p:bldLst>` : "";
  return `<p:timing><p:tnLst><p:par><p:cTn id="${root}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${seq}${media}</p:childTnLst></p:cTn></p:par></p:tnLst>${buildList}</p:timing>`;
}

// A slide's transition. `t`: "fade", { push: "l" | "r" | "u" | "d" } or undefined for none.
export function transitionXml(t) {
  if (t === "fade") return "<p:transition><p:fade/></p:transition>";
  if (t?.push) return `<p:transition><p:push dir="${t.push}"/></p:transition>`;
  return "";
}
