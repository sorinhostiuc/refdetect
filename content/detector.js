// Attaches to RefDetect (loaded before this file via bootstrap.js)
RefDetect.Detector = {

  /**
   * Main entry point. Detects reference manager(s) in a .docx or .odt file.
   * @param {Uint8Array} fileData - raw bytes
   * @param {string} fileName - for extension detection
   * @returns {Promise<object>} detection results
   */
  detect: async function(fileData, fileName) {
    if (fileData && !(fileData instanceof Uint8Array)) {
      fileData = new Uint8Array(fileData);
    }
    var ext = (fileName || "").toLowerCase().split(".").pop();
    if (ext === "odt") {
      return await this._detectODT(fileData);
    }
    return await this._detectDOCX(fileData);
  },

  // =============================================
  // DOCX
  // =============================================

  _detectDOCX: async function(fileData) {
    var zip = await JSZip.loadAsync(fileData);
    var result = this._emptyResult("docx");

    // Level 1: field codes + content controls
    var xmlPaths = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"];
    for (var i = 0; i < xmlPaths.length; i++) {
      var file = zip.file(xmlPaths[i]);
      if (!file) continue;
      var text = await file.async("string");
      this._scanFieldCodes(text, result);
      this._scanContentControls(text, result);
    }

    // Level 2: customXml
    var allFiles = Object.keys(zip.files);
    for (var j = 0; j < allFiles.length; j++) {
      var f = allFiles[j];
      if (f.startsWith("customXml/item") && f.endsWith(".xml") && f.indexOf("Props") === -1) {
        var cxText = await zip.file(f).async("string");
        this._scanCustomXml(cxText, result);
      }
    }

    // Level 3: webExtensions
    for (var k = 0; k < allFiles.length; k++) {
      var wf = allFiles[k];
      if (wf.startsWith("word/webextensions/") && wf.endsWith(".xml")) {
        var weText = await zip.file(wf).async("string");
        this._scanWebExtensions(weText, result);
      }
    }

    // Level 4: settings.xml (docVars)
    var settingsFile = zip.file("word/settings.xml");
    if (settingsFile) {
      var settingsText = await settingsFile.async("string");
      this._scanSettings(settingsText, result);
    }

    this._finalizeResults(result);
    return result;
  },

  // =============================================
  // ODT
  // =============================================

  _detectODT: async function(fileData) {
    var zip = await JSZip.loadAsync(fileData);
    var result = this._emptyResult("odt");

    var contentFile = zip.file("content.xml");
    if (!contentFile) {
      this._finalizeResults(result);
      return result;
    }
    var text = await contentFile.async("string");
    this._scanODTReferenceMarks(text, result);

    this._finalizeResults(result);
    return result;
  },

  _emptyResult: function(format) {
    return {
      managers: [],
      fieldCodes: { total: 0, citation: 0, bibliography: 0 },
      contentControls: { total: 0, citation: 0 },
      customXml: [],
      format: format,
      raw: []
    };
  },

  // =============================================
  // LEVEL 1a: FIELD CODES
  // =============================================

  _scanFieldCodes: function(xmlText, result) {
    var instrRegex = /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/g;
    var match;

    while ((match = instrRegex.exec(xmlText)) !== null) {
      var instr = match[1].trim();
      if (!instr) continue;
      result.fieldCodes.total++;

      // -- Zotero --
      if (instr.indexOf("ZOTERO_ITEM") !== -1) {
        this._hit(result, "Zotero", "field_code", "ADDIN ZOTERO_ITEM", "citation");
        result.fieldCodes.citation++;
      }
      else if (instr.indexOf("ZOTERO_BIBL") !== -1) {
        this._hit(result, "Zotero", "field_code", "ADDIN ZOTERO_BIBL", "bibliography");
        result.fieldCodes.bibliography++;
      }
      else if (instr.indexOf("ZOTERO_PREF") !== -1) {
        this._hit(result, "Zotero", "field_code", "ADDIN ZOTERO_PREF", "preferences");
      }

      // -- Mendeley Desktop --
      // Uses "ADDIN CSL_CITATION" like Zotero but JSON contains "mendeley" key
      else if (instr.indexOf("CSL_CITATION") !== -1 && instr.indexOf("ZOTERO") === -1) {
        if (instr.indexOf('"mendeley"') !== -1 || instr.indexOf("&quot;mendeley&quot;") !== -1) {
          this._hit(result, "Mendeley Desktop", "field_code", "ADDIN CSL_CITATION + mendeley key", "citation");
        } else {
          this._hit(result, "Mendeley Desktop (or CSL-based)", "field_code", "ADDIN CSL_CITATION", "citation");
        }
        result.fieldCodes.citation++;
      }

      // -- Paperpile --
      else if (instr.indexOf("paperpile_citation") !== -1 || instr.indexOf("PAPERPILE") !== -1) {
        this._hit(result, "Paperpile", "field_code", "ADDIN paperpile_citation", "citation");
        result.fieldCodes.citation++;
      }
      else if (instr.indexOf("paperpile_bibliography") !== -1) {
        this._hit(result, "Paperpile", "field_code", "ADDIN paperpile_bibliography", "bibliography");
        result.fieldCodes.bibliography++;
      }

      // -- EndNote --
      // Must check EN.CITE.DATA before EN.CITE (substring match)
      else if (instr.indexOf("EN.CITE.DATA") !== -1) {
        this._hit(result, "EndNote", "field_code", "ADDIN EN.CITE.DATA", "traveling_library");
      }
      else if (instr.indexOf("EN.CITE") !== -1) {
        this._hit(result, "EndNote", "field_code", "ADDIN EN.CITE", "citation");
        result.fieldCodes.citation++;
      }
      else if (instr.indexOf("EN.REFLIST") !== -1) {
        this._hit(result, "EndNote", "field_code", "ADDIN EN.REFLIST", "bibliography");
        result.fieldCodes.bibliography++;
      }

      // -- RefWorks --
      else if (instr.indexOf("RW.CITE") !== -1) {
        this._hit(result, "RefWorks", "field_code", "ADDIN RW.CITE", "citation");
        result.fieldCodes.citation++;
      }
      else if (instr.indexOf("RW.BIB") !== -1) {
        this._hit(result, "RefWorks", "field_code", "ADDIN RW.BIB", "bibliography");
        result.fieldCodes.bibliography++;
      }

      // -- F1000 / Papers / ReadCube --
      else if (instr.indexOf("F1000") !== -1) {
        this._hit(result, "F1000 / Papers / ReadCube", "field_code", "ADDIN F1000", "citation");
        result.fieldCodes.citation++;
      }

      // -- Bookends --
      else if (instr.indexOf("BKENDS") !== -1) {
        this._hit(result, "Bookends", "field_code", "ADDIN BKENDS", "citation");
        result.fieldCodes.citation++;
      }

      // -- ProCite --
      else if (instr.indexOf("PROCITE") !== -1) {
        this._hit(result, "ProCite", "field_code", "ADDIN PROCITE", "citation");
        result.fieldCodes.citation++;
      }

      // -- Reference Manager --
      else if (instr.indexOf("REFMGR") !== -1) {
        this._hit(result, "Reference Manager", "field_code", "ADDIN REFMGR", "citation");
        result.fieldCodes.citation++;
      }

      // -- Sente --
      else if (instr.indexOf("SENTE") !== -1) {
        this._hit(result, "Sente", "field_code", "ADDIN SENTE", "citation");
        result.fieldCodes.citation++;
      }

      // -- Colwiz --
      else if (instr.indexOf("COLWIZ") !== -1) {
        this._hit(result, "Colwiz", "field_code", "ADDIN COLWIZ", "citation");
        result.fieldCodes.citation++;
      }

      // -- Sciwheel --
      else if (instr.indexOf("Sciwheel") !== -1) {
        this._hit(result, "Sciwheel", "field_code", "ADDIN Sciwheel", "citation");
        result.fieldCodes.citation++;
      }

      // -- NoteExpress --
      else if (instr.indexOf("NE.Cite") !== -1 || instr.indexOf("NE.CITE") !== -1) {
        this._hit(result, "NoteExpress", "field_code", "ADDIN NE.Cite", "citation");
        result.fieldCodes.citation++;
      }

      // -- Word Built-in Bibliography --
      else if (/^\s*CITATION\s+\S+/.test(instr)) {
        this._hit(result, "Word Built-in", "field_code", "CITATION Tag", "citation");
        result.fieldCodes.citation++;
      }
      else if (/^\s*BIBLIOGRAPHY\b/.test(instr)) {
        this._hit(result, "Word Built-in", "field_code", "BIBLIOGRAPHY", "bibliography");
        result.fieldCodes.bibliography++;
      }
    }
  },

  // =============================================
  // LEVEL 1b: CONTENT CONTROLS
  // =============================================

  _scanContentControls: function(xmlText, result) {
    var sdtRegex = /<w:sdt>([\s\S]*?)<\/w:sdt>/g;
    var match;

    while ((match = sdtRegex.exec(xmlText)) !== null) {
      var inner = match[1];
      result.contentControls.total++;

      var tagMatch = inner.match(/<w:tag\s+w:val="([^"]*)"/);
      var aliasMatch = inner.match(/<w:alias\s+w:val="([^"]*)"/);
      var tag = tagMatch ? tagMatch[1] : "";
      var alias = aliasMatch ? aliasMatch[1] : "";
      var tagLower = tag.toLowerCase();
      var aliasLower = alias.toLowerCase();

      // Citavi
      if (tagLower.indexOf("citaviplaceholder") !== -1) {
        this._hit(result, "Citavi", "content_control", "CitaviPlaceholder#UUID", "citation");
        result.contentControls.citation++;
      }
      // Mendeley Cite
      else if (tagLower.indexOf("mendeley") !== -1 || aliasLower.indexOf("mendeley") !== -1) {
        this._hit(result, "Mendeley Cite", "content_control", "Mendeley SDT", "citation");
        result.contentControls.citation++;
      }
      // Citavi bibliography
      else if (tagLower.indexOf("citavi") !== -1 &&
               (aliasLower.indexOf("literaturverzeichnis") !== -1 || aliasLower.indexOf("bibliography") !== -1)) {
        this._hit(result, "Citavi", "content_control", "Citavi Bibliography SDT", "bibliography");
      }
      // Word Built-in bibliography SDT
      else if (inner.indexOf('w:val="Bibliographies"') !== -1) {
        this._hit(result, "Word Built-in", "content_control", "docPartGallery Bibliographies", "bibliography");
      }
    }
  },

  // =============================================
  // LEVEL 2: CUSTOM XML
  // =============================================

  _scanCustomXml: function(xmlText, result) {
    if (xmlText.indexOf("schemas.openxmlformats.org/officeDocument/2006/bibliography") !== -1) {
      var sourceCount = (xmlText.match(/<b:Source>/g) || []).length;
      if (sourceCount > 0) {
        this._hit(result, "Word Built-in", "custom_xml",
          "MS Bibliography XML (" + sourceCount + " sources)", "metadata");
        result.customXml.push({ type: "ms_bibliography", sourceCount: sourceCount });
      }
    }
    if (xmlText.indexOf("Citavi") !== -1 || xmlText.indexOf("citavi") !== -1) {
      this._hit(result, "Citavi", "custom_xml", "Citavi metadata in customXml", "metadata");
      result.customXml.push({ type: "citavi_metadata" });
    }
  },

  // =============================================
  // LEVEL 3: WEB EXTENSIONS
  // =============================================

  _scanWebExtensions: function(xmlText, result) {
    if (xmlText.indexOf("Mendeley") !== -1 || xmlText.indexOf("mendeley") !== -1) {
      this._hit(result, "Mendeley Cite", "web_extension", "Mendeley Office Add-in", "addon");
    }
    if (xmlText.indexOf("Citavi") !== -1 || xmlText.indexOf("citavi") !== -1) {
      this._hit(result, "Citavi", "web_extension", "Citavi Office Add-in", "addon");
    }
    if (xmlText.indexOf("Paperpile") !== -1 || xmlText.indexOf("paperpile") !== -1) {
      this._hit(result, "Paperpile", "web_extension", "Paperpile Office Add-in", "addon");
    }
  },

  // =============================================
  // LEVEL 4: SETTINGS.XML (docVars)
  // =============================================

  _scanSettings: function(xmlText, result) {
    if (xmlText.indexOf("ZOTERO_PREF") !== -1) {
      this._hit(result, "Zotero", "doc_var", "ZOTERO_PREF in docVars", "preferences");
    }
    if (xmlText.indexOf("ZOTERO_BREF_") !== -1) {
      this._hit(result, "Zotero", "doc_var", "ZOTERO_BREF_ bookmarks", "citation");
    }
  },

  // =============================================
  // LEVEL 5: ODT REFERENCE MARKS
  // =============================================

  _scanODTReferenceMarks: function(xmlText, result) {
    var zoteroMarks = (xmlText.match(/ZOTERO_ITEM\s+CSL_CITATION/g) || []).length;
    if (zoteroMarks > 0) {
      this._hit(result, "Zotero", "reference_mark",
        "ZOTERO_ITEM CSL_CITATION (" + zoteroMarks + "x)", "citation", zoteroMarks);
    }
    if (xmlText.indexOf("ZOTERO_BIBL") !== -1) {
      this._hit(result, "Zotero", "reference_mark", "ZOTERO_BIBL", "bibliography");
    }
    var jabrefMarks = (xmlText.match(/JR_cite\d+_\d+_/g) || []).length;
    if (jabrefMarks > 0) {
      this._hit(result, "JabRef", "reference_mark",
        "JR_cite reference marks (" + jabrefMarks + "x)", "citation", jabrefMarks);
    }
    if (xmlText.indexOf('text:name="JR_bib"') !== -1) {
      this._hit(result, "JabRef", "reference_mark", "JR_bib section", "bibliography");
    }
  },

  // =============================================
  // HELPERS
  // =============================================

  _hit: function(result, manager, mechanism, signature, role, count) {
    result.raw.push({
      manager: manager,
      mechanism: mechanism,
      signature: signature,
      role: role,
      count: count || 1
    });
  },

  _finalizeResults: function(result) {
    var map = {};
    var i, hit, m, name;

    for (i = 0; i < result.raw.length; i++) {
      hit = result.raw[i];
      if (!map[hit.manager]) {
        map[hit.manager] = {
          name: hit.manager,
          mechanisms: [],
          signatures: [],
          citations: 0,
          hasBibliography: false,
          hasMetadata: false,
          hasPreferences: false,
          confidence: 0
        };
      }
      m = map[hit.manager];
      if (m.mechanisms.indexOf(hit.mechanism) === -1) m.mechanisms.push(hit.mechanism);
      m.signatures.push(hit.signature);
      if (hit.role === "citation") m.citations += hit.count;
      if (hit.role === "bibliography") m.hasBibliography = true;
      if (hit.role === "metadata" || hit.role === "traveling_library") m.hasMetadata = true;
      if (hit.role === "preferences") m.hasPreferences = true;
    }

    var managers = [];
    for (name in map) {
      m = map[name];
      var conf = 0;
      if (m.citations > 0) conf += 60;
      if (m.citations > 5) conf += 10;
      if (m.hasBibliography) conf += 15;
      if (m.hasMetadata) conf += 10;
      if (m.hasPreferences) conf += 5;
      m.confidence = Math.min(conf, 100);
      managers.push(m);
    }

    managers.sort(function(a, b) {
      return (b.confidence - a.confidence) || (b.citations - a.citations);
    });

    result.managers = managers;
  }
};
