(function (window) {
  // Keep lastParams in scope for onOpen and “Highlight more”
  let lastParams = null;

  // Utility: does the current document have any non‐empty paragraph?
  function documentHasText() {
    const doc = Api.GetDocument();
    return doc.GetAllParagraphs().some(p => p.GetText().trim() !== "");
  }

  // Initialize and wire up the sidebar UI
  function initPluginUI() {
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

    if (!searchInput || !applyBtn) {
      console.error("Required sidebar elements missing");
      return;
    }

    // Toggle Apply button enabled state
    function refreshApplyButton() {
      applyBtn.disabled = !(documentHasText() && searchInput.value.trim() !== "");
    }

    searchInput.addEventListener("input", refreshApplyButton);
    ignoreCaseBox.addEventListener("change", refreshApplyButton);

    applyBtn.addEventListener("click", () => {
      // Hide “no matches” if showing
      if (noMatchesDiv) noMatchesDiv.style.display = "none";

      const rawSearch     = searchInput.value.trim();
      if (!rawSearch) return;

      const params = {
        rawSearch,
        ignoreCase:    ignoreCaseBox.checked,
        highlightColor: highlightSel.value,
        textColor:     textColorPicker.value,
        doBold:        boldCheckbox.checked,
        doItalic:      italicCheckbox.checked,
        doUnderline:   underlineCheckbox.checked,
        doStrike:      strikeCheckbox.checked
      };

      lastParams = params;

      // Defer document changes into the editor context
      this.callCommand(() => {
        const {
          rawSearch, ignoreCase,
          highlightColor, textColor,
          doBold, doItalic, doUnderline, doStrike
        } = Asc.scope.lastParams;

        const doc        = Api.GetDocument();
        const paragraphs = doc.GetAllParagraphs();
        let matchesFound = 0;

        const escaped = rawSearch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const flags   = ignoreCase ? "gi" : "g";
        const regex   = new RegExp("(" + escaped + ")", flags);

        function equals(a, b) {
          return ignoreCase
            ? a.toLowerCase() === b.toLowerCase()
            : a === b;
        }

        paragraphs.forEach(paragraph => {
          const text = paragraph.GetText();
          if (!text || !regex.test(text)) return;

          matchesFound++;
          const segments = text.split(regex);
          paragraph.RemoveAllElements();

          segments.forEach(segment => {
            const run = Api.CreateRun();
            run.AddText(segment);

            if (equals(segment, rawSearch)) {
              if (highlightColor !== "NoFill") run.SetHighlight(highlightColor);
              const r = parseInt(textColor.slice(1,3), 16);
              const g = parseInt(textColor.slice(3,5), 16);
              const b = parseInt(textColor.slice(5,7), 16);
              run.SetColor(r, g, b, false);

              if (doBold)      run.SetBold(true);
              if (doItalic)    run.SetItalic(true);
              if (doUnderline) run.SetUnderline(true);
              if (doStrike)    run.SetStrikeout(true);
            }

            paragraph.AddElement(run);
          });
        });

        // After processing, show “no matches” or “highlight more” link
        if (matchesFound === 0 && noMatchesDiv) {
          noMatchesDiv.style.display = "block";
        }
        if (highlightMoreLink) {
          highlightMoreLink.style.display = "block";
        }
      }, true);

      // Close the sidebar so the user sees the updated doc
      Api.CloseSidebar();
    });

    // “Highlight more” → reopen and repopulate
    if (highlightMoreLink) {
      highlightMoreLink.addEventListener("click", () => {
        if (!lastParams) return;
        Api.OpenSidebar("index.html");
      });
    }

    // Called by ONLYOFFICE after sidebar HTML is injected
    window.onOpen = () => {
      if (!lastParams) return;

      document.getElementById("searchText").value       = lastParams.rawSearch;
      document.getElementById("ignoreCase").checked     = lastParams.ignoreCase;
      document.getElementById("highlightColor").value   = lastParams.highlightColor;
      document.getElementById("textColor").value        = lastParams.textColor;
      document.getElementById("boldCheckbox").checked   = lastParams.doBold;
      document.getElementById("italicCheckbox").checked = lastParams.doItalic;
      document.getElementById("underlineCheckbox").checked = lastParams.doUnderline;
      document.getElementById("strikeCheckbox").checked = lastParams.doStrike;

      refreshApplyButton();
    };

    // Finally, initialize button state when the panel’s DOM is ready
    document.addEventListener("DOMContentLoaded", refreshApplyButton);
  }

  // Plugin entry point → ONLYOFFICE calls this when the user opens the panel
  window.Asc.plugin.init = function () {
    initPluginUI();
  };

  // Handle “OK” / “Close” buttons in a modal variant (if you set one up)
  window.Asc.plugin.button = function (id) {
    this.executeCommand("close", "");
  };

  // Expose lastParams into the editor context for callCommand
  Asc.scope.lastParams = () => lastParams;
})(window);
