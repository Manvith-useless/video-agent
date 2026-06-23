// Global utilities available on all pages

document.addEventListener("DOMContentLoaded", () => {
  // Highlight active nav link
  const path = window.location.pathname;
  document.querySelectorAll(".nav-links a").forEach(a => {
    if (a.getAttribute("href") === path) {
      a.style.background = "rgba(255,255,255,.15)";
      a.style.color = "#fff";
    }
  });
});
