var RefDetectDialog = {
  _io: null,
  _detector: null,
  _fileData: null,
  _fileName: null,
  _results: null,

  init: function() {
    this._io = window.arguments?.[0] || {};

    // FIX: get Detector from io object (passed by refdetect.js openDialog)
    this._detector = this._io.Detector;
    if (!this._detector && this._io.Zotero && this._io.Zotero.RefDetect) {
      this._detector = this._io.Zotero.RefDetect.Detector;
    }

    this._setupDropZone();
    this._setupFileInput();
  },

  // ---- FILE HANDLING ----

  _setupDropZone: function() {
    var dropZone = document.getElementById("rd-drop-zone");
    var self = this;

    dropZone.addEventListener("click", function() {
      document.getElementById("rd-file-input").click();
    });
    dropZone.addEventListener("dragover", function(e) {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add("rd-dragover");
    });
    dropZone.addEventListener("dragleave", function(e) {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("rd-dragover");
    });
    dropZone.addEventListener("drop", function(e) {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("rd-dragover");
      var files = e.dataTransfer?.files;
      if (files && files.length > 0) self._handleFile(files[0]);
    });
  },

  _setupFileInput: function() {
    var input = document.getElementById("rd-file-input");
    var self = this;
    input.addEventListener("change", function(e) {
      if (e.target.files && e.target.files.length > 0) self._handleFile(e.target.files[0]);
    });
  },

  _handleFile: function(file) {
    var name = file.name || "";
    var ext = name.toLowerCase().split(".").pop();
    if (ext !== "docx" && ext !== "odt") {
      if (this._io.Zotero) this._io.Zotero.alert(window, "RefDetect", "Please select a .docx or .odt file.");
      return;
    }
    this._fileName = name;
    var self = this;
    var reader = new FileReader();
    reader.onload = function(evt) {
      self._fileData = new Uint8Array(evt.target.result);
      self._showFileInfo();
    };
    reader.readAsArrayBuffer(file);
  },

  _showFileInfo: function() {
    document.getElementById("rd-file-name").value = this._fileName;
    document.getElementById("rd-file-info").hidden = false;
    document.getElementById("rd-drop-zone").hidden = true;
    document.getElementById("rd-btn-detect").disabled = false;
  },

  clearFile: function() {
    this._fileData = null;
    this._fileName = null;
    document.getElementById("rd-file-info").hidden = true;
    document.getElementById("rd-drop-zone").hidden = false;
    document.getElementById("rd-btn-detect").disabled = true;
    document.getElementById("rd-file-input").value = "";
  },

  // ---- DETECTION ----

  runDetection: async function() {
    if (!this._fileData) return;
    if (!this._detector) {
      if (this._io.Zotero) {
        this._io.Zotero.debug("RefDetect: Detector not available in dialog scope");
        this._io.Zotero.alert(window, "RefDetect Error", "Internal error: Detector module not available.");
      }
      return;
    }

    this._showScreen("progress");
    try {
      document.getElementById("rd-progress-label").value = "Parsing document structure...";
      await new Promise(function(r) { setTimeout(r, 50); });

      this._results = await this._detector.detect(this._fileData, this._fileName);

      document.getElementById("rd-progress-label").value = "Analysis complete.";
      await new Promise(function(r) { setTimeout(r, 200); });
      this._showResults();
    } catch (err) {
      if (this._io.Zotero) {
        this._io.Zotero.debug("RefDetect error: " + err.message + "\n" + (err.stack || ""));
        this._io.Zotero.alert(window, "RefDetect Error", "Error analyzing document:\n" + err.message);
      }
      this._showScreen("picker");
    }
  },

  // ---- RESULTS ----

  _showResults: function() {
    this._showScreen("results");
    var r = this._results;
    document.getElementById("rd-result-file").value = "File: " + this._fileName + " (" + r.format + ")";

    if (r.managers.length === 0) {
      document.getElementById("rd-primary-result").hidden = true;
      document.getElementById("rd-additional-managers").hidden = true;
      document.getElementById("rd-signatures-section").hidden = true;
      document.getElementById("rd-no-results").hidden = false;
      return;
    }

    document.getElementById("rd-no-results").hidden = true;
    document.getElementById("rd-primary-result").hidden = false;
    document.getElementById("rd-signatures-section").hidden = false;

    var primary = r.managers[0];
    document.getElementById("rd-detected-manager").value = primary.name;
    document.getElementById("rd-confidence").value = primary.confidence + "%";
    document.getElementById("rd-citation-count").value = String(primary.citations);
    document.getElementById("rd-has-bibliography").value = primary.hasBibliography ? "Yes" : "No";
    document.getElementById("rd-mechanism").value = primary.mechanisms.join(", ");

    var fill = document.getElementById("rd-confidence-fill");
    fill.style.width = primary.confidence + "%";
    var confLabel = document.getElementById("rd-confidence");
    if (primary.confidence >= 80) {
      fill.style.background = "#28a745"; confLabel.style.color = "#28a745";
    } else if (primary.confidence >= 50) {
      fill.style.background = "#ffc107"; confLabel.style.color = "#b8860b";
    } else {
      fill.style.background = "#dc3545"; confLabel.style.color = "#dc3545";
    }

    // Additional managers
    if (r.managers.length > 1) {
      document.getElementById("rd-additional-managers").hidden = false;
      var list = document.getElementById("rd-additional-list");
      while (list.firstChild) list.removeChild(list.firstChild);
      for (var i = 1; i < r.managers.length; i++) {
        var m = r.managers[i];
        var item = document.createXULElement("hbox");
        item.className = "rd-additional-item";
        var nl = document.createXULElement("label");
        nl.className = "rd-add-name";
        nl.value = m.name;
        item.appendChild(nl);
        var dl = document.createXULElement("label");
        dl.value = " (" + m.citations + " citations, " + m.confidence + "% confidence)";
        item.appendChild(dl);
        list.appendChild(item);
      }
    } else {
      document.getElementById("rd-additional-managers").hidden = true;
    }

    // Signatures
    var sigList = document.getElementById("rd-signatures-list");
    while (sigList.firstChild) sigList.removeChild(sigList.firstChild);
    for (var j = 0; j < r.raw.length; j++) {
      var sig = r.raw[j];
      var div = document.createElementNS("http://www.w3.org/1999/xhtml", "html:div");
      div.className = "rd-sig-item";
      div.textContent = sig.signature;
      var mech = document.createElementNS("http://www.w3.org/1999/xhtml", "html:span");
      mech.className = "rd-sig-mechanism";
      mech.textContent = " [" + sig.mechanism + ", " + sig.role + "]";
      div.appendChild(mech);
      sigList.appendChild(div);
    }
  },

  // ---- COPY REPORT ----

  copyReport: function() {
    if (!this._results) return;
    var r = this._results;
    var rpt = "RefDetect Report\n================\n";
    rpt += "File: " + this._fileName + "\nFormat: " + r.format + "\n\n";

    if (r.managers.length === 0) {
      rpt += "Result: No reference manager detected.\n";
    } else {
      rpt += "Detected managers:\n";
      for (var i = 0; i < r.managers.length; i++) {
        var m = r.managers[i];
        rpt += "  " + m.name + "\n";
        rpt += "    Confidence: " + m.confidence + "%\n";
        rpt += "    Citations: " + m.citations + "\n";
        rpt += "    Bibliography: " + (m.hasBibliography ? "Yes" : "No") + "\n";
        rpt += "    Mechanisms: " + m.mechanisms.join(", ") + "\n\n";
      }
      rpt += "All signatures:\n";
      for (var j = 0; j < r.raw.length; j++) {
        var s = r.raw[j];
        rpt += "  " + s.signature + " [" + s.mechanism + ", " + s.role + "]\n";
      }
    }
    rpt += "\nField codes: " + r.fieldCodes.total + " total, " + r.fieldCodes.citation + " citations, " + r.fieldCodes.bibliography + " bibliography\n";
    rpt += "Content controls: " + r.contentControls.total + " total, " + r.contentControls.citation + " citations\n";

    var clipHelper = Cc["@mozilla.org/widget/clipboardhelper;1"].getService(Ci.nsIClipboardHelper);
    clipHelper.copyString(rpt);
    if (this._io.Zotero) this._io.Zotero.alert(window, "RefDetect", "Report copied to clipboard.");
  },

  // ---- NAVIGATION ----

  _showScreen: function(name) {
    var screens = ["picker", "progress", "results"];
    for (var i = 0; i < screens.length; i++) {
      var el = document.getElementById("rd-screen-" + screens[i]);
      if (el) el.hidden = (screens[i] !== name);
    }
  },

  reset: function() {
    this._fileData = null;
    this._fileName = null;
    this._results = null;
    document.getElementById("rd-file-info").hidden = true;
    document.getElementById("rd-drop-zone").hidden = false;
    document.getElementById("rd-btn-detect").disabled = true;
    document.getElementById("rd-file-input").value = "";
    this._showScreen("picker");
  }
};
