var RefDetect;
var chromeHandle;

if (typeof setImmediate === "undefined") {
  var setImmediate = fn => setTimeout(fn, 0);
}

function install(data, reason) {}

function startup({ id, version, resourceURI, rootURI }, reason) {
  var aomStartup = Cc["@mozilla.org/addons/addon-manager-startup;1"]
    .getService(Ci.amIAddonManagerStartup);
  var manifestURI = Services.io.newURI(rootURI + "manifest.json");
  chromeHandle = aomStartup.registerChrome(manifestURI, [
    ["content", "refdetect", rootURI + "content/"]
  ]);

  var chromeBase = "chrome://refdetect/content/";
  Services.scriptloader.loadSubScript(chromeBase + "lib/jszip.min.js");
  Services.scriptloader.loadSubScript(chromeBase + "refdetect.js");
  Services.scriptloader.loadSubScript(chromeBase + "detector.js");

  RefDetect.init({ id, version, rootURI });
}

function shutdown({ id, version, resourceURI, rootURI }, reason) {
  if (reason === APP_SHUTDOWN) return;
  RefDetect.shutdown();
  Zotero.RefDetect = undefined;
  RefDetect = undefined;
  if (chromeHandle) {
    chromeHandle.destruct();
    chromeHandle = null;
  }
}

function uninstall(data, reason) {}
