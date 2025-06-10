(function(window, undefined) {
  // Keep lastParams in a scope accessible to onOpen and highlightMore handler
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

  function initPluginUI() {
    // Now that the sidebar DOM is loaded, get elements by ID
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

    // Ensure elements exist before binding
    if (!searchInput || !ignoreCaseBox || !highlightSel || !textColorPicker || !applyBtn) {
      console.error("Some UI elements not found in sidebar DOM");
      return;
    }

    // Enable or disable “Apply” based on conditions
    function refreshApplyButton() {
      const hasDocText = documentHasText();
      const hasSearch  = searchInput.value.trim() !== "";
      applyBtn.disabled = !(hasDocText && hasSearch);
    }

    // Bind events now that elements exist
    searchInput.addEventListener("input", refreshApplyButton);
    ignoreCaseBox.addEventListener("change", refreshApplyButton);

    // MAIN “Apply” click handler
    applyBtn.addEventListener("click", function () {
      noMatchesDiv.style.display = "none"; // hide previous
      const rawSearch     = searchInput.value.trim();
      if (!rawSearch) return;

      const ignoreCase    = ignoreCaseBox.checked;
      const highlightColor = highlightSel.value; // e.g. "yellow" or "NoFill"
      const textColor     = textColorPicker.value; // hex like "#ff0000"
      const doBold        = boldCheckbox.checked;
      const doItalic      = italicCheckbox.checked;
      const doUnderline   = underlineCheckbox.checked;
      const doStrike      = strikeCheckbox.checked;

      lastParams = { rawSearch, ignoreCase, highlightColor, textColor, doBold, doItalic, doUnderline, doStrike };

      const doc = Api.GetDocument();
      const paragraphs = doc.GetAllParagraphs();
      let matchesFound = 0;

      function equalsWord(a, b) {
        return ignoreCase ? a.toLowerCase() === b.toLowerCase() : a === b;
      }

      // Build regex for splitting, escaping special chars
      const escaped = rawSearch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const flags = ignoreCase ? "gi" : "g";
      const regex = new RegExp("(" + escaped + ")", flags);

      paragraphs.forEach(function(paragraph) {
        const text = paragraph.GetText();
        if (!text) return;
        if (!regex.test(text)) return; // no match

        matchesFound++;
        const segments = text.split(regex);
        paragraph.RemoveAllElements();
        segments.forEach(function(segment) {
          const run = Api.CreateRun();
          run.AddText(segment);
          if (equalsWord(segment, rawSearch)) {
            // Highlight if not NoFill
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

      if (matchesFound === 0) {
        noMatchesDiv.style.display = "block";
      }
      highlightMoreLink.style.display = "block";
      Api.CloseSidebar();
    });

    // “Highlight more” click handler: reopen sidebar and repopulate
    highlightMoreLink.addEventListener("click", function () {
      if (!lastParams) return;
      Api.OpenSidebar("index.html");
      // onOpen handler (below) will pick up lastParams and fill fields
    });

    // ONLYOFFICE calls window.Asc.plugin.onOpen or window.onOpen after sidebar loads.
    window.onOpen = function() {
      if (!lastParams) return;
      const si = document.getElementById("searchText");
      if (si) si.value = lastParams.rawSearch;
      const ic = document.getElementById("ignoreCase");
      if (ic) ic.checked = lastParams.ignoreCase;
      const hs = document.getElementById("highlightColor");
      if (hs) hs.value = lastParams.highlightColor;
      const tc = document.getElementById("textColor");
      if (tc) tc.value = lastParams.textColor;
      document.getElementById("boldCheckbox").checked    = lastParams.doBold;
      document.getElementById("italicCheckbox").checked  = lastParams.doItalic;
      document.getElementById("underlineCheckbox").checked = lastParams.doUnderline;
      document.getElementById("strikeCheckbox").checked  = lastParams.doStrike;
      refreshApplyButton();
    };

    // On first load of sidebar HTML
    document.addEventListener("DOMContentLoaded", function() {
      refreshApplyButton();
    });
  }

  // Entry point for ONLYOFFICE plugin
  window.Asc.plugin.init = function(text) {
    // When sidebar is opened, this is called.
    // Now we can safely init our UI.
    initPluginUI();
  };

  window.Asc.plugin.button = function(id) {
    // Called when user clicks OK/Close in plugin
    this.executeCommand("close", "");
  };

})(window, undefined);
