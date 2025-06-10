(function (window, undefined) {
  // Keep lastParams accessible to onOpen and highlightMore handler
  let lastParams = null;

  // Utility: check if document has any non-empty paragraph
  function documentHasText() {
    const doc = Api.GetDocument();
    const paragraphs = doc.GetAllParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      if (paragraphs[i].GetText().trim() !== "") {
        return true;
      }
    }
    return false;
  }

  // Initialize the sidebar UI once DOM is loaded
  function initPluginUI() {
    // Retrieve elements by ID
    const searchInput       = document.getElementById("searchText");
    const ignoreCaseBox     = document.getElementById("ignoreCase");
    const highlightSel      = document.getElementById("highlightColor");
    const textColorPicker   = document.getElementById("textColor");
    const boldCheckbox      = document.getElementById("boldCheckbox");
    const italicCheckbox    = document.getElementById("italicCheckbox");
    const underlineCheckbox = document.getElementById("underlineCheckbox");
    const strikeCheckbox    = document.getElementById("strikeCheckbox");
    const applyBtn          = document.getElementById("applyBtn");
    const noMatchesDiv      = document.getElementById("noMatches");
    const highlightMoreLink = document.getElementById("highlightMore");

    // If essential elements are missing, abort initialization
    if (!searchInput || !ignoreCaseBox || !highlightSel || !textColorPicker || !applyBtn) {
      console.error("Sidebar UI elements not found.");
      return;
    }

    // Enable or disable “Apply” based on document content and search input
    function refreshApplyButton() {
      const hasDocText = documentHasText();
      const hasSearch  = searchInput.value.trim() !== "";
      applyBtn.disabled = !(hasDocText && hasSearch);
    }

    // Bind input/change events to re-check Apply button
    searchInput.addEventListener("input", refreshApplyButton);
    ignoreCaseBox.addEventListener("change", refreshApplyButton);

    // MAIN “Apply” click handler
    applyBtn.addEventListener("click", function () {
      // Hide “No matches” message when re-applying
      if (noMatchesDiv) {
        noMatchesDiv.style.display = "none";
      }

      const rawSearch = searchInput.value.trim();
      if (!rawSearch) {
        return;
      }
      const ignoreCase    = ignoreCaseBox.checked;
      const highlightColor = highlightSel.value;  // e.g. "yellow" or "NoFill"
      const textColor     = textColorPicker.value; // "#rrggbb"
      const doBold        = boldCheckbox.checked;
      const doItalic      = italicCheckbox.checked;
      const doUnderline   = underlineCheckbox.checked;
      const doStrike      = strikeCheckbox.checked;

      // Save parameters for "Highlight more"
      lastParams = { rawSearch, ignoreCase, highlightColor, textColor, doBold, doItalic, doUnderline, doStrike };

      const doc = Api.GetDocument();
      const paragraphs = doc.GetAllParagraphs();
      let matchesFound = 0;

      // Helper for comparing segments
      function equalsWord(a, b) {
        return ignoreCase ? a.toLowerCase() === b.toLowerCase() : a === b;
      }

      // Build regex that escapes special characters
      const escaped = rawSearch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const flags = ignoreCase ? "gi" : "g";
      const regex = new RegExp("(" + escaped + ")", flags);

      paragraphs.forEach(function (paragraph) {
        const text = paragraph.GetText();
        if (!text) return;
        if (!regex.test(text)) {
          return; // no match
        }
        matchesFound++;
        // Split into segments around the matches
        const segments = text.split(regex);
        paragraph.RemoveAllElements();

        segments.forEach(function (segment) {
          const run = Api.CreateRun();
          run.AddText(segment);

          if (equalsWord(segment, rawSearch)) {
            // Apply highlight if not "NoFill"
            if (highlightColor !== "NoFill") {
              run.SetHighlight(highlightColor);
            }
            // Convert hex to RGB
            const r = parseInt(textColor.slice(1, 3), 16);
            const g = parseInt(textColor.slice(3, 5), 16);
            const b = parseInt(textColor.slice(5, 7), 16);
            run.SetColor(r, g, b, false);

            if (doBold)      run.SetBold(true);
            if (doItalic)    run.SetItalic(true);
            if (doUnderline) run.SetUnderline(true);
            if (doStrike)    run.SetStrikeout(true);
          }
          paragraph.AddElement(run);
        });
      });

      if (matchesFound === 0 && noMatchesDiv) {
        noMatchesDiv.style.display = "block";
      }
      if (highlightMoreLink) {
        highlightMoreLink.style.display = "block";
      }
      // Close sidebar to show results
      Api.CloseSidebar();
    });

    // “Highlight more” click → reopen sidebar and repopulate fields
    if (highlightMoreLink) {
      highlightMoreLink.addEventListener("click", function () {
        if (!lastParams) return;
        Api.OpenSidebar("index.html");
        // onOpen handler below will refill values
      });
    }

    // Called by ONLYOFFICE after sidebar HTML loads
    window.onOpen = function () {
      if (!lastParams) return;
      const si = document.getElementById("searchText");
      if (si) si.value = lastParams.rawSearch;
      const ic = document.getElementById("ignoreCase");
      if (ic) ic.checked = lastParams.ignoreCase;
      const hs = document.getElementById("highlightColor");
      if (hs) hs.value = lastParams.highlightColor;
      const tc = document.getElementById("textColor");
      if (tc) tc.value = lastParams.textColor;
      const bc = document.getElementById("boldCheckbox");
      if (bc) bc.checked = lastParams.doBold;
      const icb = document.getElementById("italicCheckbox");
      if (icb) icb.checked = lastParams.doItalic;
      const ub = document.getElementById("underlineCheckbox");
      if (ub) ub.checked = lastParams.doUnderline;
      const sb = document.getElementById("strikeCheckbox");
      if (sb) sb.checked = lastParams.doStrike;
      refreshApplyButton();
    };

    // When sidebar DOM is ready
    document.addEventListener("DOMContentLoaded", function () {
      refreshApplyButton();
    });
  }

  // Entry point for ONLYOFFICE plugin: when sidebar is opened
  window.Asc.plugin.init = function init(text) {
    // Initialize the UI in the loaded sidebar
    initPluginUI();
  };

  // Handle plugin button (OK/Close); typical pattern closes sidebar
  window.Asc.plugin.button = function button(id) {
    this.executeCommand("close", "");
  };

})(window, undefined);
