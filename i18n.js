
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

    dictionaries[code] = {
      ...(dictionaries[code] || {}),
      ...entries
    };

    if (code === currentLanguage) {
      document.documentElement.lang = code;
    }

    // Also supports translations loaded after the page is ready,
    // including when scripts are restored from portable HTML.
    if (document.readyState !== "loading") {
      updatePageText();
    }
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


  
  // Update interface text and translated attributes.
  function updatePageText() {

    // Simple text elements only: do not mark elements
    // containing icons, inputs or other nested controls.
    document.querySelectorAll("[data-i18n]").forEach(element => {
      const key = element.getAttribute("data-i18n");

      if (key) {
        element.textContent = translate(key);
      }
    });

    // Translate tooltips, placeholders and accessibility labels.
    ["title", "placeholder", "aria-label"].forEach(attribute => {
      const marker = "data-i18n-" + attribute;

      document.querySelectorAll("[" + marker + "]").forEach(element => {
        const key = element.getAttribute(marker);

        if (key) {
          element.setAttribute(attribute, translate(key));
        }
      });
    });

    // Keep the language selector synchronised.
    const selector = document.getElementById("languageSelect");

    if (selector) {
      selector.value = currentLanguage;
    }
  }

  function initialisePageLanguage() {

    // Use English if a previously remembered language
    // is not available in this version of the application.
    if (
      !Object.prototype.hasOwnProperty.call(dictionaries, currentLanguage) &&
      Object.prototype.hasOwnProperty.call(dictionaries, "en")
    ) {
      currentLanguage = "en";
    }

    document.documentElement.lang = currentLanguage;

    const selector = document.getElementById("languageSelect");

    if (selector) {
      selector.value = currentLanguage;

      selector.addEventListener("change", () => {
        if (!setLanguage(selector.value)) {
          selector.value = currentLanguage;
        }
      });
    }

    updatePageText();
  }

  // Initialise once the HTML has loaded.
  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      initialisePageLanguage
    );
  } else {
    initialisePageLanguage();
  }

  // Update the interface whenever the language changes.
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

