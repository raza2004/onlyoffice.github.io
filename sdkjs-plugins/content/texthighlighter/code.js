(function(window) {
  "use strict";

  // === 1) Dropdown toggle bindings (run once on DOM load) ===
  document.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll(".dropdown-header").forEach(header => {
      header.addEventListener("click", () => {
        header.parentElement.classList.toggle("open");
      });
    });
  });

  // 2) Theme change handler
  function onThemeChanged(theme) {
    // Let OnlyOffice apply its base styling
    window.Asc.plugin.onThemeChangedBase(theme);

    // Log to console
    console.log("Detected theme:", theme.type);

    // Toggle our dark-mode class
    document.body.classList.toggle("dark-mode", theme.type === "dark");
  }
  window.Asc.plugin.attachEvent?.("onThemeChanged", onThemeChanged);
  window.Asc.plugin.onThemeChanged = onThemeChanged;

  // === 2) ONLYOFFICE plugin code ===

  // UI elements (filled in init)
  let searchInput, ignoreCaseBox, applyBtn;
  let stateInput, stateNo, stateDone;
  let loader, foundCountSpan;
  let highlightMore1, highlightMore2, revertBtn;

  window.Asc.plugin.init = function(text) {
    // Cache DOM nodes
    searchInput    = document.getElementById("searchText");
    ignoreCaseBox  = document.getElementById("ignoreCase");
    applyBtn       = document.getElementById("applyButton");
    stateInput     = document.getElementById("state-input");
    stateNo        = document.getElementById("state-no-results");
    stateDone      = document.getElementById("state-done");
    loader         = document.getElementById("loader");
    foundCountSpan = document.getElementById("foundCount");
    highlightMore1 = document.getElementById("highlightMore1");
    highlightMore2 = document.getElementById("highlightMore2");
    revertBtn      = document.getElementById("revertButton");

    // Simple state-switchers
    function showInput() {
      stateInput.style.display = "";
      stateNo.style.display    = "none";
      stateDone.style.display  = "none";
      loader.style.display     = "none";
    }
    function showNoResults() {
      stateInput.style.display = "none";
      stateNo.style.display    = "";
      stateDone.style.display  = "none";
      loader.style.display     = "none";
    }
    function showDone(count) {
      stateInput.style.display = "none";
      stateNo.style.display    = "none";
      stateDone.style.display  = "";
      loader.style.display     = "none";
      foundCountSpan.textContent = count;
    }

    // Wire up UI
    applyBtn.disabled = true;
    searchInput.addEventListener("input", () => {
      applyBtn.disabled = !searchInput.value.trim();
    });
    applyBtn.addEventListener("click", onApply);
    highlightMore1.addEventListener("click", showInput);
    highlightMore2.addEventListener("click", showInput);
    revertBtn.addEventListener("click", onRevert);

    // If the plugin was opened with a selection, use it
    if (text && text.trim()) {
      searchInput.value = text.trim();
      applyBtn.disabled = false;
    }

    // React when user changes selection in the document
    if (window.Asc.plugin.attachEvent) {
      window.Asc.plugin.attachEvent("onSelectionChanged", sel => {
        if (sel && sel.text) {
          searchInput.value = sel.text;
          applyBtn.disabled = false;
        }
      });
    }

    // Initialize last-term storage
    Asc.scope.lastTerm     = "";
    Asc.scope.lastCaseSens = false;

    // Show the initial state
    showInput();
  };

  // 3) Apply highlights
  function onApply() {
    const term    = searchInput.value.trim();
    const caseSens= !ignoreCaseBox.checked;
    
     const hlColor    = document.getElementById("highlightColor").value;
    const txtColor   = document.getElementById("textColor") .value;
    const doBold     = document.getElementById("boldCheckbox") .checked;
    const doItalic   = document.getElementById("italicCheckbox").checked;
    const doUnder    = document.getElementById("underlineCheckbox").checked;
    const doStrike   = document.getElementById("strikeCheckbox").checked;

    // remember for revert
    Asc.scope.lastTerm      = term;
    Asc.scope.lastCaseSens  = caseSens;
    Asc.scope.lastHlColor   = hlColor;
    Asc.scope.lastTxtColor  = txtColor;
    Asc.scope.lastDoBold    = doBold;
    Asc.scope.lastDoItalic  = doItalic;
    Asc.scope.lastDoUnderline = doUnder;
    Asc.scope.lastDoStrike  = doStrike;


    // transition UI
    stateInput.style.display = "none";
    stateNo.style.display    = "none";
    stateDone.style.display  = "none";
    loader.style.display     = "";

    window.Asc.plugin.callCommand(function() {
      const results = Api.GetDocument().Search(Asc.scope.lastTerm, Asc.scope.lastCaseSens);
      results.forEach(function(range) {
        range.SetHighlight(Asc.scope.lastHlColor);
        if (Asc.scope.lastDoBold)      range.SetBold(true);
        if (Asc.scope.lastDoItalic)    range.SetItalic(true);
        if (Asc.scope.lastDoUnderline) range.SetUnderline(true);
        if (Asc.scope.lastDoStrike)    range.SetStrikeout(true);
        if (Asc.scope.lastTxtColor !== "#000000") {
          var rgb = Asc.scope.lastTxtColor
            .slice(1)
            .match(/.{2}/g)
            .map(h => parseInt(h, 16));
          range.SetColor(rgb[0], rgb[1], rgb[2], false);
        }
      });
      return results.length;
    }, false);
  }

  // 4) Revert highlights
  function onRevert() {
    loader.style.display = "";
    window.Asc.plugin.callCommand(function() {
      const results = Api.GetDocument().Search(Asc.scope.lastTerm, Asc.scope.lastCaseSens);
      results.forEach(r => {
        r.SetHighlight("none");
        r.SetBold(false);
        r.SetItalic(false);
        r.SetUnderline(false);
        r.SetStrikeout(false);
        r.SetColor(0,0,0,false);
      });
      return results.length;
    }, false);
  }

  // 5) After each callCommand
  window.Asc.plugin.onCommandCallback = function(count) {
    const n = Number(count) || 0;
    if (!Asc.scope.lastTerm) {
      // fallback
      document.getElementById("state-input").style.display = "";
    } else if (n === 0) {
      document.getElementById("state-no-results").style.display = "";
    } else {
      document.getElementById("state-done").style.display = "";
      foundCountSpan.textContent = n;
    }
    loader.style.display = "none";
  };

})(window);



