<!-- A full-slide drawing moved by requestAnimationFrame for 1 s after the click. Loaded directly at a
     click state, it draws that state's end with no motion, as a slide that was already clicked does. -->
<script setup>
import { ref, watch } from "vue";
import { useSlideContext } from "@slidev/client";

const END = { 0: 100, 1: 860 };
const { $clicks } = useSlideContext();
const x = ref(END[$clicks.value] ?? END[1]);

watch($clicks, (k) => {
  const from = x.value;
  const to = END[k] ?? END[1];
  const start = performance.now();
  requestAnimationFrame(function step(now) {
    const p = Math.min(1, (now - start) / 1000);
    x.value = from + (to - from) * p;
    if (p < 1) requestAnimationFrame(step);
  });
});
</script>

<template>
  <svg viewBox="0 0 980 552" style="position: absolute; left: 0; top: 0; width: 980px; height: 552px; z-index: -1">
    <rect width="980" height="552" fill="#dfe8ef" />
    <circle :cx="x" cy="276" r="60" fill="#c0392b" />
  </svg>
</template>
