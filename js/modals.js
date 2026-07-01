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