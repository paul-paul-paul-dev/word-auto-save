/* global Office */

Office.onReady(function () {
  // Persist "load on startup" so Word auto-starts this add-in on every document
  // open from now on — user only needs to open the taskpane once ever.
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);

  if (typeof window.startAutoSave === "function") {
    window.startAutoSave();
  }

  populateInterval();
  populateLastSaved();

  document.getElementById("save-btn").addEventListener("click", applyInterval);
  document.getElementById("save-now-btn").addEventListener("click", saveNow);

  updateTimerStatus();
});

function populateInterval() {
  var stored = localStorage.getItem("autoSave_intervalMs");
  if (stored) {
    var minutes = Math.round(parseInt(stored, 10) / 60000);
    if (!isNaN(minutes) && minutes > 0) {
      document.getElementById("interval").value = minutes;
    }
  }
}

function populateLastSaved() {
  var ts = localStorage.getItem("autoSave_lastSaved");
  var el = document.getElementById("last-saved");
  if (ts) {
    el.textContent = new Date(ts).toLocaleTimeString();
  } else {
    el.textContent = "—";
  }
}

function applyInterval() {
  var minutes = parseInt(document.getElementById("interval").value, 10);
  if (isNaN(minutes) || minutes < 1) {
    showStatus("Please enter a valid number of minutes (minimum 1).", true);
    return;
  }

  var ms = minutes * 60000;

  if (typeof window.restartWithNewInterval === "function") {
    window.restartWithNewInterval(ms);
    showStatus("Interval updated.");
    updateTimerStatus();
  } else {
    showStatus("Could not update interval — shared runtime not ready.", true);
  }
}

function saveNow() {
  if (typeof window.performSave === "function") {
    window.performSave();
    showStatus("Saving…");
    setTimeout(function () {
      populateLastSaved();
      showStatus("Saved.");
    }, 1500);
  } else {
    showStatus("Save function not available.", true);
  }
}

function updateTimerStatus() {
  var el = document.getElementById("timer-status");
  if (!el) return;
  var stored = localStorage.getItem("autoSave_intervalMs");
  var ms = stored ? parseInt(stored, 10) : 5 * 60 * 1000;
  var minutes = Math.round(ms / 60000);
  el.textContent = "Timer active — saving every " + minutes + " minute" + (minutes === 1 ? "" : "s") + ".";
}

function showStatus(message, isError) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status" + (isError ? " error" : "");
}
