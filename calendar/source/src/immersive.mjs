// Visual preferences do not touch calendar storage or Microsoft authentication.
const motionButton = document.getElementById("motion-toggle");
const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)");
let pausedByUser = false;

function updateMotion() {
  const paused = pausedByUser || reducedMotion.matches;
  document.documentElement.classList.toggle("ambient-paused", paused);
  motionButton.disabled = reducedMotion.matches;
  const label = reducedMotion.matches
    ? "Animation disabled by reduced motion preference"
    : paused
      ? "Resume ambient animation"
      : "Pause ambient animation";
  motionButton.setAttribute("aria-label", label);
  motionButton.title = label;
}

motionButton.addEventListener("click", () => {
  pausedByUser = !pausedByUser;
  updateMotion();
});
reducedMotion.addEventListener("change", updateMotion);
updateMotion();
