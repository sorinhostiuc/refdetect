var RefDetect = {
  id: null,
  version: null,
  rootURI: null,
  _initialized: false,
  Detector: null, // populated by detector.js (loaded after this file)

  init: function({ id, version, rootURI }) {
    if (this._initialized) return;
    this.id = id;
    this.version = version;
    this.rootURI = rootURI;
    this._initialized = true;

    // FIX: attach to Zotero global so the dialog window can access it
    Zotero.RefDetect = this;

    this._addMenuWhenReady();
    Zotero.debug("RefDetect: initialized v" + version);
  },

  _addMenuWhenReady: async function() {
    if (Zotero.uiReadyPromise) {
      await Zotero.uiReadyPromise;
    }
    var win = Zotero.getMainWindow();
    if (!win) {
      Zotero.debug("RefDetect: no main window");
      return;
    }
    this._addMenuItem(win.document);
  },

  shutdown: function() {
    this._removeMenuItem();
    this._initialized = false;
    Zotero.debug("RefDetect: shutdown");
  },

  _addMenuItem: function(doc) {
    if (!doc) doc = Zotero.getMainWindow()?.document;
    if (!doc) return;
    var menuTools = doc.getElementById("menu_ToolsPopup");
    if (!menuTools) return;

    var sep = doc.createXULElement("menuseparator");
    sep.id = "refdetect-separator";
    menuTools.appendChild(sep);

    var mi = doc.createXULElement("menuitem");
    mi.id = "refdetect-menu-detect";
    mi.setAttribute("label", "Detect Reference Manager...");
    mi.addEventListener("command", () => this.openDialog());
    menuTools.appendChild(mi);
  },

  _removeMenuItem: function() {
    var doc = Zotero.getMainWindow()?.document;
    if (!doc) return;
    var el1 = doc.getElementById("refdetect-menu-detect");
    var el2 = doc.getElementById("refdetect-separator");
    if (el1) el1.remove();
    if (el2) el2.remove();
  },

  openDialog: function() {
    // FIX: pass Detector reference explicitly so the dialog can use it
    var io = {
      Detector: this.Detector,
      Zotero: Zotero
    };

    Zotero.getMainWindow().openDialog(
      "chrome://refdetect/content/ui/dialog.xhtml",
      "refdetect-dialog",
      "chrome,dialog=no,centerscreen,resizable=yes,width=580,height=520",
      io
    );
  }
};
