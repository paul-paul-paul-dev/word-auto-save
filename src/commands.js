/* global Office, Word */

// Must be at top level (not inside Office.onReady) for OnDocumentOpened to fire
Office.actions.associate("onDocOpen", onDocOpen);

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
    context.document.save();
    return context.sync().then(function () {
      localStorage.setItem("autoSave_lastSaved", new Date().toISOString());
    });
  }).catch(function (err) {
    // New unsaved documents will throw here — that's fine, just ignore
    console.warn("Auto-save skipped or failed:", err.message || err);
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

if (typeof window !== "undefined") {
  window.startAutoSave = startAutoSave;
  window.restartWithNewInterval = restartWithNewInterval;
  window.performSave = performSave;
}
