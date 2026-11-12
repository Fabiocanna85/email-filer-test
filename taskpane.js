/*
 * Email Filer (test) - task pane
 * Compose: choose file name and location, then confirm.
 * Read:    show the filing choice stored on the email (to test whether it
 *          survives sending).
 */

var DATE_FORMAT_YYMMDD = true;       // YYMMDD, as agreed. Set false for DDMMYY.
var MAX_RECENT = 50;
var SETTINGS_KEY = "recentFolders";

var allFolders = [];
var selectedFolder = "";
var customProps = null;

function $(id) { return document.getElementById(id); }

Office.onReady(function (info) {
  if (!info || !info.host) {
    $("lead").textContent = "";
    $("notOutlook").classList.remove("hidden");
    return;
  }

  var item = Office.context.mailbox.item;
  // In read mode the subject is a plain string; in compose it is an object.
  var isCompose = item && typeof item.subject !== "string";

  if (isCompose) {
    startCompose();
  } else {
    startRead();
  }
});

/* ---------------------------------------------------------------- compose */

function startCompose() {
  $("lead").textContent = "Choose a file name and location, then confirm. You can then send.";
  $("composePanel").classList.remove("hidden");

  allFolders = loadFolders();
  renderFolders(allFolders);
  prefillInitials();
  updatePreview();

  $("fromInit").addEventListener("input", updatePreview);
  $("toInit").addEventListener("input", updatePreview);
  $("textPart").addEventListener("input", updatePreview);
  $("refreshInit").addEventListener("click", prefillInitials);

  $("searchBtn").addEventListener("click", runSearch);
  $("searchBox").addEventListener("keydown", function (e) { if (e.key === "Enter") runSearch(); });
  $("folderList").addEventListener("change", function () {
    var v = $("folderList").value;
    if (v) setSelected(v);
  });
  $("addBtn").addEventListener("click", addFolder);
  $("removeBtn").addEventListener("click", removeSelectedFolder);
  $("confirmBtn").addEventListener("click", confirmFiling);

  Office.context.mailbox.item.loadCustomPropertiesAsync(function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      customProps = r.value;
      var existingFolder = customProps.get("filingFolder");
      var existingName = customProps.get("filingFileName");
      if (existingFolder) setSelected(existingFolder);
      if (customProps.get("filingConfirmed") === "yes" && existingName) {
        setStatus("Already confirmed as: " + existingName + ". Press Send.", "ok");
      }
    }
  });
}

function prefillInitials() {
  var item = Office.context.mailbox.item;
  if (item.from && item.from.getAsync) {
    item.from.getAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
        $("fromInit").value = initialsOf(r.value.displayName);
        updatePreview();
      }
    });
  }
  if (item.to && item.to.getAsync) {
    item.to.getAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.length > 0) {
        $("toInit").value = initialsOf(r.value[0].displayName);   // first recipient only
        updatePreview();
      }
    });
  }
}

function initialsOf(name) {
  if (!name) return "";
  // Ignore anything in brackets, e.g. "Jane Bloggs (Arup)"
  name = name.replace(/\(.*?\)/g, " ");
  return name.split(/[\s,]+/).filter(Boolean).map(function (w) {
    return w.charAt(0).toUpperCase();
  }).join("");
}

function dateString() {
  var d = new Date();
  var dd = ("0" + d.getDate()).slice(-2);
  var mm = ("0" + (d.getMonth() + 1)).slice(-2);
  var yy = String(d.getFullYear()).slice(-2);
  return DATE_FORMAT_YYMMDD ? yy + mm + dd : dd + mm + yy;
}

function buildFileName() {
  var from = $("fromInit").value.trim().toUpperCase() || "??";
  var to = $("toInit").value.trim().toUpperCase() || "??";
  var text = $("textPart").value.trim();
  var name = dateString() + " - " + from + " to " + to + " - " + text;
  return name.replace(/[\\\/:*?"<>|]/g, "-");
}

function updatePreview() {
  $("preview").textContent = buildFileName() + ".eml";
}

/* ---------------------------------------------------------------- folders */

function normalise(p) {
  return String(p || "").trim().replace(/\\/g, "/").replace(/\/+$/, "");
}

function loadFolders() {
  try {
    var raw = Office.context.roamingSettings.get(SETTINGS_KEY);
    var list = raw ? JSON.parse(raw) : [];
    return Array.isArray(list) ? list : [];
  } catch (e) {
    return [];
  }
}

function saveFolders(list, done) {
  Office.context.roamingSettings.set(SETTINGS_KEY, JSON.stringify(list));
  Office.context.roamingSettings.saveAsync(function (r) {
    if (done) done(r.status === Office.AsyncResultStatus.Succeeded);
  });
}

function dedupe(list) {
  var seen = {};
  var out = [];
  list.forEach(function (p) {
    var n = normalise(p);
    var key = n.toLowerCase();
    if (n && !seen[key]) { seen[key] = true; out.push(n); }
  });
  return out.slice(0, MAX_RECENT);
}

function renderFolders(list) {
  var sel = $("folderList");
  sel.innerHTML = "";
  if (list.length === 0) {
    var o = document.createElement("option");
    o.textContent = allFolders.length === 0
      ? "(No locations yet - add one below)"
      : "(No matching locations)";
    o.disabled = true;
    sel.appendChild(o);
    return;
  }
  list.forEach(function (p) {
    var o = document.createElement("option");
    o.value = p; o.textContent = p;
    if (normalise(p).toLowerCase() === normalise(selectedFolder).toLowerCase()) o.selected = true;
    sel.appendChild(o);
  });
}

function runSearch() {
  var q = $("searchBox").value.trim().toLowerCase();
  if (!q) { renderFolders(allFolders); return; }
  renderFolders(allFolders.filter(function (p) { return p.toLowerCase().indexOf(q) !== -1; }));
}

function addFolder() {
  var p = normalise($("newFolder").value);
  if (!p) { setStatus("Type a location to add.", "err"); return; }
  allFolders = dedupe([p].concat(allFolders));   // newest first, no duplicates
  saveFolders(allFolders);
  $("newFolder").value = "";
  $("searchBox").value = "";
  setSelected(p);
  renderFolders(allFolders);
  setStatus("", "");
}

function removeSelectedFolder() {
  var v = $("folderList").value;
  if (!v) { setStatus("Select a location to remove.", "err"); return; }
  allFolders = allFolders.filter(function (p) {
    return normalise(p).toLowerCase() !== normalise(v).toLowerCase();
  });
  saveFolders(allFolders);
  if (normalise(selectedFolder).toLowerCase() === normalise(v).toLowerCase()) setSelected("");
  runSearch();
  setStatus("", "");
}

function setSelected(p) {
  selectedFolder = normalise(p);
  var box = $("selectedBox");
  if (selectedFolder) {
    box.textContent = selectedFolder;
    box.classList.remove("empty");
  } else {
    box.textContent = "No location selected";
    box.classList.add("empty");
  }
}

/* ---------------------------------------------------------------- confirm */

function confirmFiling() {
  var text = $("textPart").value.trim();
  if (!text) { setStatus("Enter a description for the file name.", "err"); return; }
  if (!selectedFolder) { setStatus("Select a location.", "err"); return; }
  if (!customProps) { setStatus("Could not store the choice on this email. Close and reopen the panel.", "err"); return; }

  var fileName = buildFileName();
  customProps.set("filingConfirmed", "yes");
  customProps.set("filingFolder", selectedFolder);
  customProps.set("filingFileName", fileName);
  customProps.set("filingConfirmedAt", new Date().toISOString());

  customProps.saveAsync(function (r) {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Could not save the choice: " + (r.error && r.error.message), "err");
      return;
    }
    allFolders = dedupe([selectedFolder].concat(allFolders));
    saveFolders(allFolders);
    setStatus("Confirmed. Now press Send.", "ok");
  });
}

function setStatus(msg, kind) {
  var s = $("status");
  s.textContent = msg;
  s.className = kind || "";
}

/* ------------------------------------------------------------------- read */

function startRead() {
  $("lead").textContent = "This shows whether the filing choice stayed with the email after sending.";
  $("readPanel").classList.remove("hidden");

  Office.context.mailbox.item.loadCustomPropertiesAsync(function (r) {
    var dl = $("storedList");
    dl.innerHTML = "";
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      addRow(dl, "Result", "Could not read stored properties: " + (r.error && r.error.message));
      return;
    }
    var keys = ["filingConfirmed", "filingFileName", "filingFolder", "filingConfirmedAt"];
    var found = false;
    keys.forEach(function (k) {
      var v = r.value.get(k);
      if (v) { found = true; addRow(dl, k, v); }
    });
    if (!found) addRow(dl, "Result", "Nothing stored on this email.");
  });
}

function addRow(dl, term, value) {
  var dt = document.createElement("dt"); dt.textContent = term;
  var dd = document.createElement("dd"); dd.textContent = value;
  dl.appendChild(dt); dl.appendChild(dd);
}
