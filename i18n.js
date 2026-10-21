
/* SimpleSPC language support — initial framework.
 * Language dictionaries and interface translation will be added in later steps.
 * This file does not change any existing interface text by itself.
 */
(function () {
  "use strict";

  const STORAGE_KEY = "simplespc.language";
  const dictionaries = Object.create(null);
  let currentLanguage = "en";

  // English is the default; honour a previous manual choice when one exists.
  try {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (saved && /^[a-z]{2}(?:-[A-Za-z]{2})?$/.test(saved)) {
      currentLanguage = saved;
    }
  } catch (error) {
    // Portable HTML files may not have access to browser storage.
  }

  function registerLanguage(code, entries) {
    if (!code || !entries || typeof entries !== "object") return;
    dictionaries[code] = { ...(dictionaries[code] || {}), ...entries };
    if (code === currentLanguage) document.documentElement.lang = code;
  }

  function translate(key, values = {}) {
    const text = dictionaries[currentLanguage]?.[key]
      ?? dictionaries.en?.[key]
      ?? key;

    return String(text).replace(/\{([a-zA-Z0-9_]+)\}/g, (match, name) =>
      Object.prototype.hasOwnProperty.call(values, name)
        ? String(values[name])
        : match
    );
  }

  function setLanguage(code, options = {}) {
    if (!Object.prototype.hasOwnProperty.call(dictionaries, code)) return false;

    currentLanguage = code;
    document.documentElement.lang = code;

    if (options.remember !== false) {
      try {
        window.localStorage.setItem(STORAGE_KEY, code);
      } catch (error) {
        // Continue without persistence if browser storage is unavailable.
      }
    }

    document.dispatchEvent(new CustomEvent("simplespc:languagechange", {
      detail: { language: code }
    }));
    return true;
  }


  // Update text in HTML elements marked with data-i18n.
  function updatePageText() {
    document.querySelectorAll("[data-i18n]").forEach(element => {
      const key = element.getAttribute("data-i18n");
      if (key) {
        element.textContent = translate(key);
      }
    });
  }

  // Translate marked elements when the page first loads.
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", updatePageText);
  } else {
    updatePageText();
  }

  // Refresh translated elements when the language changes.
  document.addEventListener(
    "simplespc:languagechange",
    updatePageText
  );

  window.SimpleSPCi18n = {
    registerLanguage,
    translate,
    setLanguage,
    getLanguage: () => currentLanguage
  };
})();

