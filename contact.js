"use strict";

(() => {
  const frame = document.querySelector("#inquiry-frame");
  const notice = document.querySelector("#inquiry-connection");
  const openPanel = document.querySelector("#inquiry-open");
  const openLink = document.querySelector("#inquiry-open-link");
  const openHelp = document.querySelector("#inquiry-open-help");
  if (!frame) return;
  const endpoint = window.KSME_CONTACT_CONFIG?.webAppUrl?.trim() || "";
  const connected = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(endpoint);
  const formUrl = connected ? endpoint : "contact-form.html";
  if (frame.getAttribute("src") !== formUrl) frame.src = formUrl;
  notice.hidden = connected;
  let receivedHeight = false;
  if (connected) {
    openLink.href = endpoint;
    openPanel.hidden = false;
    setTimeout(() => {
      if (receivedHeight) return;
      frame.hidden = true;
      openHelp.hidden = false;
      openPanel.classList.add("standalone");
      openLink.textContent = "문의·방문 신청 작성 ↗";
    }, 12000);
  }
  if (endpoint && !connected) {
    notice.textContent = "접수 창구 연결을 확인하고 있습니다. 현재 화면은 미리보기이며 입력한 정보는 전송되지 않습니다.";
  }
  window.addEventListener("message", event => {
    const local = !connected && event.source === frame.contentWindow && event.origin === window.location.origin;
    const google = connected && (event.origin === "https://script.google.com" || /^https:\/\/[a-z0-9-]+-script\.googleusercontent\.com$/.test(event.origin));
    if ((!local && !google) || event.data?.channel !== "ksme-inquiry-height") return;
    const height = Number(event.data.height);
    if (Number.isFinite(height)) {
      receivedHeight = true;
      frame.hidden = false;
      openHelp.hidden = true;
      openPanel.classList.remove("standalone");
      openLink.textContent = "문의·신청 양식 새 창에서 열기 ↗";
      frame.style.height = `${Math.min(6000, Math.max(600, Math.ceil(height)))}px`;
    }
  });
})();
