/* global Office, Word */

Office.onReady(function () {
  Office.actions.associate("onDocOpen", onDocOpen);
});

var autoSaveTimer = null;

function onDocOpen(event) {
  startAutoSave();
  event.completed();
}

function startAutoSave() {
  var intervalMs = getIntervalMs();

  if (autoSaveTimer !== null) {
    clearInterval(autoSaveTimer);
    autoSaveTimer = null;
  }

  autoSaveTimer = setInterval(function () {
    performSave();
  }, intervalMs);
}

function performSave() {
  Word.run(function (context) {
    var doc = context.document;
    context.load(doc, "saved, url");
    return context.sync().then(function () {
      // Skip save if document has never been saved to disk (no URL = unsaved new document)
      if (!doc.url || doc.url === "") {
        return;
      }
      doc.save();
      return context.sync().then(function () {
        var timestamp = new Date().toISOString();
        localStorage.setItem("autoSave_lastSaved", timestamp);
      });
    });
  }).catch(function (err) {
    console.error("Auto-save failed:", err);
  });
}

function restartWithNewInterval(ms) {
  localStorage.setItem("autoSave_intervalMs", String(ms));
  startAutoSave();
}

function getIntervalMs() {
  var stored = localStorage.getItem("autoSave_intervalMs");
  if (stored) {
    var parsed = parseInt(stored, 10);
    if (!isNaN(parsed) && parsed > 0) {
      return parsed;
    }
  }
  return 5 * 60 * 1000; // default: 5 minutes
}

// Expose for taskpane.js access via shared runtime
/* global restartWithNewInterval, performSave */
if (typeof window !== "undefined") {
  window.restartWithNewInterval = restartWithNewInterval;
  window.performSave = performSave;
}
