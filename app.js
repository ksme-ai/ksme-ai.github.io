"use strict";

const galleryImages = {
  electric: [
    {src:"assets/photos/01-hd-electric-group.webp", alt:"HD현대 표지 앞에 함께 선 연구회 방문단 6명", caption:"HD현대일렉트릭 방문단", full:"2026.04.03 · HD현대일렉트릭 방문 · 단체사진 (밝기·기울기 보정)"},
    {src:"assets/photos/01-hd-electric-meeting.webp", alt:"HD현대일렉트릭 회의실에서 전력기기와 AI 연구를 논의하는 연구진", caption:"전력기기의 AI 활용을 논의한 회의 현장", full:"2026.04.03 · HD현대일렉트릭 · 전력설비 AI 연구 교류"}
  ],
  mobis: [
    {src:"assets/photos/03-mobis-meeting.webp?v=privacy-20261004", alt:"참석자의 얼굴을 흐리게 처리한 현대모비스 마북연구소 세미나 사진", caption:"현대모비스 연구진과의 세미나", full:"2026.07.28 · 현대모비스 마북연구소 · 연구회·기업 관계자들과의 교류"},
    {src:"assets/photos/03-mobis-presentation.webp", alt:"현대모비스 마북연구소에서 발표 화면을 보며 연구 내용을 공유하는 참석자들", caption:"AI 연구개발 사례를 공유한 발표 현장", full:"2026.07.28 · 현대모비스 마북연구소 · AI 연구 발표"}
  ]
};

document.querySelectorAll("[data-gallery]").forEach(gallery => {
  const items = galleryImages[gallery.dataset.gallery];
  if (!items?.length) return;
  let current = 0;
  const show = direction => {
    current = (current + direction + items.length) % items.length;
    const item = items[current];
    const trigger = gallery.querySelector(".photo-open");
    const img = trigger.querySelector("img");
    img.src = item.src;
    img.alt = item.alt;
    trigger.dataset.image = item.src;
    trigger.dataset.caption = item.full;
    gallery.querySelector(".gallery-caption").textContent = item.caption;
    gallery.querySelector("[data-counter]").textContent = `${current + 1} / ${items.length}`;
  };
  gallery.querySelector("[data-prev]").addEventListener("click", () => show(-1));
  gallery.querySelector("[data-next]").addEventListener("click", () => show(1));
});

const lightbox = document.querySelector(".lightbox");
const closeButton = lightbox.querySelector(".lightbox-close");
let previousFocus;
document.querySelectorAll(".photo-open").forEach(trigger => {
  trigger.addEventListener("click", () => {
    previousFocus = trigger;
    const photo = lightbox.querySelector("img");
    photo.src = trigger.dataset.image;
    photo.alt = trigger.querySelector("img").alt;
    lightbox.querySelector(".lightbox-caption").textContent = trigger.dataset.caption;
    lightbox.showModal();
    document.body.classList.add("modal-open");
    closeButton.focus();
  });
});
closeButton.addEventListener("click", () => lightbox.close());
lightbox.addEventListener("click", event => { if (event.target === lightbox) { const r=lightbox.getBoundingClientRect(); if(event.clientX<r.left||event.clientX>r.right||event.clientY<r.top||event.clientY>r.bottom) lightbox.close(); } });
lightbox.addEventListener("close", () => {
  document.body.classList.remove("modal-open");
  previousFocus?.focus({preventScroll:true});
});

if ("IntersectionObserver" in window) {
  const indexLinks = [...document.querySelectorAll(".journey-index a")];
  const observer = new IntersectionObserver(entries => {
    for (const entry of entries) {
      if (entry.isIntersecting) {
        indexLinks.forEach(link => {
          const active = link.hash === `#${entry.target.id}`;
          link.classList.toggle("is-current", active);
          if (active) link.setAttribute("aria-current", "location");
          else link.removeAttribute("aria-current");
        });
      }
    }
  }, {rootMargin:"-15% 0px -55% 0px", threshold:0});
  document.querySelectorAll(".visit").forEach(visit => observer.observe(visit));
}

// Magnify only the currently displayed image, including its existing privacy blur.
document.querySelectorAll(".visit-gallery .photo-open").forEach(trigger => {
  const img = trigger.querySelector("img");
  const lens = document.createElement("span");
  lens.className = "photo-magnifier";
  lens.setAttribute("aria-hidden", "true");
  trigger.append(lens);
  const hide = () => lens.classList.remove("is-visible");
  trigger.addEventListener("pointermove", event => {
    if (event.pointerType !== "mouse" || !window.matchMedia("(hover:hover) and (pointer:fine)").matches || !img.complete || !img.naturalWidth) return hide();
    const rect = img.getBoundingClientRect();
    const parent = trigger.getBoundingClientRect();
    const x = Math.max(0, Math.min(rect.width, event.clientX - rect.left));
    const y = Math.max(0, Math.min(rect.height, event.clientY - rect.top));
    const zoom = 2.5;
    const size = Math.min(220, rect.width * .6, rect.height * .8);
    const fit = getComputedStyle(img).objectFit;
    const scale = (fit === "contain" ? Math.min : Math.max)(rect.width / img.naturalWidth, rect.height / img.naturalHeight);
    const width = img.naturalWidth * scale;
    const height = img.naturalHeight * scale;
    const offsetX = (rect.width - width) / 2;
    const offsetY = (rect.height - height) / 2;
    lens.style.width = lens.style.height = `${size}px`;
    lens.style.left = `${rect.left - parent.left + Math.max(0, Math.min(rect.width - size, x - size / 2))}px`;
    lens.style.top = `${rect.top - parent.top + Math.max(0, Math.min(rect.height - size, y - size / 2))}px`;
    lens.style.backgroundImage = `url("${img.currentSrc || img.src}")`;
    lens.style.backgroundSize = `${width * zoom}px ${height * zoom}px`;
    lens.style.backgroundPosition = `${size / 2 - (x - offsetX) * zoom}px ${size / 2 - (y - offsetY) * zoom}px`;
    lens.classList.add("is-visible");
  });
  trigger.addEventListener("pointerleave", hide);
  trigger.addEventListener("click", hide);
  img.addEventListener("load", hide);
  window.addEventListener("scroll", hide, {passive:true});
  window.addEventListener("resize", hide);
});
