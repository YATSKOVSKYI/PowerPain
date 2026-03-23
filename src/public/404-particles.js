(() => {
  const container = document.getElementById("particles");
  const count = 15;
  for (let i = 0; i < count; i++) {
    const p = document.createElement("div");
    p.className = "particle";
    const size = Math.random() * 3 + 1;
    p.style.cssText = [
      "width:" + size + "px",
      "height:" + size + "px",
      "left:" + (Math.random() * 100) + "%",
      "opacity:" + (Math.random() * 0.3 + 0.1),
      "animation-duration:" + (Math.random() * 10 + 8) + "s",
      "animation-delay:" + (Math.random() * 10) + "s",
    ].join(";");
    container.appendChild(p);
  }
})();
