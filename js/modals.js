// Universal Escape key handler for overlays
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;

  const dataEditorOverlay = document.getElementById("dataEditorOverlay");
  if (dataEditorOverlay && dataEditorOverlay.style.display !== "none") {
    dataEditorOverlay.style.display = "none";
  }

  const sheetPickerOverlay = document.getElementById("sheetPickerOverlay");
  if (sheetPickerOverlay && sheetPickerOverlay.style.display !== "none") {
    sheetPickerOverlay.style.display = "none";
  }
});

// -----------------------------
// Help section toggle
// -----------------------------
function toggleHelpSection(forceOpen) {
  const modal = document.getElementById("helpModal");
  if (!modal) return;

  const isOpen = modal.classList.contains("visible");
  const shouldOpen = typeof forceOpen === "boolean" ? forceOpen : !isOpen;

  modal.classList.toggle("visible", shouldOpen);
  modal.setAttribute("aria-hidden", shouldOpen ? "false" : "true");
  document.body.classList.toggle("modal-open", shouldOpen);

  if (shouldOpen) {
    const closeBtn = modal.querySelector(".modal-close");
    if (closeBtn) closeBtn.focus();
  }
}

(function initHelpModal() {
  const modal = document.getElementById("helpModal");
  if (!modal) return;

  // Click outside (backdrop) closes
  modal.addEventListener("click", (e) => {
    if (e.target && e.target.classList.contains("modal-backdrop")) {
      toggleHelpSection(false);
    }
  });

  // Escape closes
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && modal.classList.contains("visible")) {
      toggleHelpSection(false);
    }
  });
})();

const spcHelperCloseBtn = document.getElementById("spcHelperCloseBtn");

if (spcHelperCloseBtn) {
  spcHelperCloseBtn.addEventListener("click", () => {
    if (spcHelperPanel) {
      spcHelperPanel.classList.remove("visible");
    }
  });
}

function showDataEditorDeleteHelp() {
  if (!dataEditorDeleteHelpBtn || !dataEditorDeleteHelpPopup) return;
  dataEditorDeleteHelpPopup.classList.add("show");
  dataEditorDeleteHelpBtn.setAttribute("aria-expanded", "true");
  dataEditorDeleteHelpPopup.setAttribute("aria-hidden", "false");
}

function hideDataEditorDeleteHelp() {
  if (!dataEditorDeleteHelpBtn || !dataEditorDeleteHelpPopup) return;
  dataEditorDeleteHelpPopup.classList.remove("show");
  dataEditorDeleteHelpBtn.setAttribute("aria-expanded", "false");
  dataEditorDeleteHelpPopup.setAttribute("aria-hidden", "true");
}

function toggleDataEditorDeleteHelp() {
  if (!dataEditorDeleteHelpPopup) return;
  if (dataEditorDeleteHelpPopup.classList.contains("show")) {
    hideDataEditorDeleteHelp();
  } else {
    showDataEditorDeleteHelp();
  }
}
