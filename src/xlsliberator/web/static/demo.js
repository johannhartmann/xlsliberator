/* Live conversion demo embedded in the landing page.
 *
 * Progressive enhancement: the upload form posts to /jobs without JS (the
 * server redirects to the standalone job page). With JS we intercept the
 * submit, drive the real /api/jobs endpoints, and render progress, the event
 * log, the report and download links inline — no page navigation. */
(function () {
  "use strict";

  var root = document.getElementById("demo-root");
  if (!root) return;

  var maxUploadMb = parseInt(root.dataset.maxUploadMb, 10) || 100;

  var els = {
    upload: document.getElementById("demo-upload"),
    job: document.getElementById("demo-job"),
    form: document.getElementById("demo-form"),
    fileInput: document.getElementById("demo-file"),
    drop: document.getElementById("demo-drop"),
    dropText: document.getElementById("demo-drop-text"),
    submit: document.getElementById("demo-submit"),
    uploadError: document.getElementById("demo-upload-error"),
    samples: Array.prototype.slice.call(document.querySelectorAll(".l-sample")),
    fileName: document.getElementById("demo-filename"),
    fileExt: document.getElementById("demo-fileext"),
    jobId: document.getElementById("demo-jobid"),
    fileSize: document.getElementById("demo-filesize"),
    status: document.getElementById("demo-status"),
    chips: document.getElementById("demo-chips"),
    stage: document.getElementById("demo-stage"),
    percent: document.getElementById("demo-percent"),
    fill: document.getElementById("demo-fill"),
    log: document.getElementById("demo-log"),
    logEndpoint: document.getElementById("demo-log-endpoint"),
    report: document.getElementById("demo-report"),
    reportGrid: document.getElementById("demo-report-grid"),
    warnings: document.getElementById("demo-warnings"),
    warningsTitle: document.getElementById("demo-warnings-title"),
    warningsList: document.getElementById("demo-warnings-list"),
    downloads: document.getElementById("demo-downloads"),
    cancel: document.getElementById("demo-cancel"),
    reset: document.getElementById("demo-reset"),
  };

  var GREEN = "oklch(0.56 0.16 145)";
  var GREEN_DARK = "oklch(0.42 0.13 145)";
  var GREEN_DOT = "oklch(0.52 0.13 145)";
  var RUST = "oklch(0.5 0.14 62)";
  var RED = "oklch(0.5 0.18 25)";

  var PHASE_LABELS = ["Analyse", "Konvertierung", "Übersetzung", "Verifizierung"];
  // Backend JobPhase -> chip index, progress floor, log tag.
  var PHASE_MAP = {
    uploaded: { idx: 0, floor: 5, tag: "INFO" },
    queued: { idx: 0, floor: 8, tag: "INFO" },
    analyzing: { idx: 0, floor: 15, tag: "ANALYSE" },
    converting: { idx: 1, floor: 45, tag: "KONVERT." },
    translating: { idx: 2, floor: 70, tag: "ÜBERSETZ." },
    verifying: { idx: 3, floor: 90, tag: "VERIFIZ." },
    completed: { idx: 4, floor: 100, tag: "FERTIG" },
    failed: { idx: -1, floor: 0, tag: "FEHLER" },
    cancelled: { idx: -1, floor: 0, tag: "ABBRUCH" },
  };
  var TERMINAL = { completed: 1, failed: 1, cancelled: 1 };

  var state = {
    file: null,
    jobId: null,
    nextEvent: 0,
    percent: 0,
    phaseIdx: 0,
    finalPhase: null,
    polling: false,
  };

  // The bundled workbooks exercise transport and the basic migration pipeline;
  // they are intentionally not presented as serious migration evidence.
  els.samples.forEach(function (card) {
    var details = card.querySelector("div");
    if (!details) return;
    var label = document.createElement("div");
    label.textContent = "LEVEL 0 · BASIS-PIPELINE";
    label.style.cssText =
      "font-family:'IBM Plex Mono',monospace;font-size:9px;font-weight:600;" +
      "letter-spacing:0.04em;color:#7C828A;margin-top:7px;";
    details.appendChild(label);
  });

  // ---- helpers -------------------------------------------------------------

  function extOf(name) {
    var m = /\.([^.]+)$/.exec(name || "");
    return m ? m[1].toLowerCase() : "xlsx";
  }

  function humanSize(bytes) {
    if (!bytes && bytes !== 0) return "—";
    if (bytes < 1024) return bytes + " B";
    var kb = bytes / 1024;
    if (kb < 1024) return kb.toFixed(kb < 10 ? 1 : 0) + " KB";
    return (kb / 1024).toFixed(1) + " MB";
  }

  function deNumber(value) {
    var n = Number(value);
    if (!isFinite(n)) return String(value);
    return n.toLocaleString("de-DE");
  }

  function timeOf(iso) {
    var d = iso ? new Date(iso) : new Date();
    if (isNaN(d.getTime())) return "";
    function p(n) { return String(n).padStart(2, "0"); }
    return p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
  }

  function show(el) { el.classList.remove("hidden"); }
  function hide(el) { el.classList.add("hidden"); }

  // ---- file selection ------------------------------------------------------

  function setFile(file) {
    state.file = file;
    if (file) {
      els.dropText.textContent = file.name + " · " + humanSize(file.size);
      els.submit.textContent = "Arbeitsmappe konvertieren";
      els.submit.style.background = GREEN;
      els.submit.style.color = "#fff";
      els.submit.style.cursor = "pointer";
      els.submit.disabled = false;
    } else {
      els.dropText.textContent = "Datei hierher ziehen oder Beispiel wählen";
      els.submit.textContent = "Beispiel wählen oder Datei laden";
      els.submit.style.background = "rgba(25,27,30,0.05)";
      els.submit.style.color = "#A9ADB2";
      els.submit.style.cursor = "not-allowed";
      els.submit.disabled = true;
    }
    hide(els.uploadError);
  }

  function selectSampleCard(card) {
    els.samples.forEach(function (other) {
      var checked = other === card;
      other.classList.toggle("selected", checked);
      var mark = other.querySelector(".l-sample-check");
      if (mark) mark.classList.toggle("hidden", !checked);
    });
  }

  function clearSampleCards() {
    els.samples.forEach(function (card) {
      card.classList.remove("selected");
      var mark = card.querySelector(".l-sample-check");
      if (mark) mark.classList.add("hidden");
    });
  }

  els.fileInput.addEventListener("change", function () {
    var file = els.fileInput.files && els.fileInput.files[0];
    clearSampleCards();
    setFile(file || null);
  });

  els.samples.forEach(function (card) {
    card.addEventListener("click", function () {
      var url = card.dataset.sample;
      var name = card.dataset.name;
      els.fileInput.value = "";
      selectSampleCard(card);
      els.dropText.textContent = "Lade Beispiel …";
      fetch(url)
        .then(function (r) {
          if (!r.ok) throw new Error("HTTP " + r.status);
          return r.blob();
        })
        .then(function (blob) {
          setFile(new File([blob], name, { type: blob.type }));
        })
        .catch(function () {
          clearSampleCards();
          setFile(null);
          showUploadError("Beispiel konnte nicht geladen werden.");
        });
    });
  });

  // Drag & drop onto the dropzone.
  ["dragenter", "dragover"].forEach(function (evt) {
    els.drop.addEventListener(evt, function (e) {
      e.preventDefault();
      els.drop.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach(function (evt) {
    els.drop.addEventListener(evt, function (e) {
      e.preventDefault();
      els.drop.classList.remove("dragover");
    });
  });
  els.drop.addEventListener("drop", function (e) {
    var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) {
      clearSampleCards();
      setFile(file);
    }
  });

  function showUploadError(message) {
    els.uploadError.textContent = message;
    show(els.uploadError);
  }

  // ---- submit / job lifecycle ---------------------------------------------

  els.form.addEventListener("submit", function (e) {
    e.preventDefault();
    if (!state.file) {
      showUploadError("Bitte eine Datei oder ein Beispiel wählen.");
      return;
    }
    if (state.file.size > maxUploadMb * 1024 * 1024) {
      showUploadError("Datei überschreitet das Limit von " + maxUploadMb + " MB.");
      return;
    }
    startJob(state.file);
  });

  function startJob(file) {
    var fd = new FormData();
    fd.append("file", file, file.name);

    els.submit.disabled = true;
    els.submit.textContent = "Wird hochgeladen …";

    fetch("/api/jobs", {
      method: "POST",
      headers: { Accept: "application/json" },
      body: fd,
    })
      .then(function (r) {
        return r.json().then(function (data) {
          if (!r.ok) throw new Error(data.detail || "Upload fehlgeschlagen");
          return data;
        });
      })
      .then(function (job) {
        enterJobView(file, job);
      })
      .catch(function (err) {
        setFile(state.file);
        showUploadError(err.message || "Upload fehlgeschlagen.");
      });
  }

  function enterJobView(file, job) {
    state.jobId = job.id;
    state.nextEvent = 0;
    state.percent = 0;
    state.phaseIdx = 0;
    state.finalPhase = null;

    els.fileName.textContent = file.name;
    els.fileExt.textContent = extOf(file.name);
    els.jobId.textContent = job.id;
    els.fileSize.textContent = humanSize(file.size);
    els.logEndpoint.textContent = "/api/jobs/" + job.id + "/events";
    els.log.innerHTML = "";
    els.reportGrid.innerHTML = "";
    els.downloads.innerHTML = "";
    els.warningsList.innerHTML = "";
    hide(els.report);
    hide(els.warnings);
    hide(els.reset);
    show(els.cancel);
    els.cancel.disabled = false;
    els.fill.style.background = GREEN;
    setProgress(0, "running");
    renderChips("running");
    setStatus("running");

    hide(els.upload);
    show(els.job);

    state.polling = true;
    poll();
  }

  function poll() {
    if (!state.polling || !state.jobId) return;
    fetch("/api/jobs/" + state.jobId + "/events?since=" + state.nextEvent)
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (payload) {
        (payload.events || []).forEach(applyEvent);
        state.nextEvent = payload.next;
        if (state.finalPhase) {
          finish(state.finalPhase);
        } else if (state.polling) {
          window.setTimeout(poll, 1000);
        }
      })
      .catch(function () {
        if (state.polling) window.setTimeout(poll, 1500);
      });
  }

  function applyEvent(event) {
    var info = PHASE_MAP[event.phase] || { idx: state.phaseIdx, floor: 0, tag: "INFO" };
    // Keep the phase monotonic: the backend can emit a late "analyzing" event
    // (e.g. metadata extraction) after converting; never walk the chips back.
    if (info.idx >= 0) state.phaseIdx = Math.max(state.phaseIdx, info.idx);

    var floor = typeof event.percent === "number" ? event.percent : info.floor;
    state.percent = Math.max(state.percent, Math.min(100, floor));

    appendLogRow(event, info.tag);

    if (TERMINAL[event.phase]) {
      state.finalPhase = event.phase;
    } else {
      setProgress(state.percent, "running");
      renderChips("running");
    }
  }

  function appendLogRow(event, tag) {
    var li = document.createElement("li");
    li.style.cssText =
      "display:flex;align-items:baseline;gap:10px;padding:6px 7px;border-radius:6px;";

    var tagBase =
      "font-family:'IBM Plex Mono',monospace;font-size:9.5px;font-weight:600;" +
      "padding:2px 7px;border-radius:5px;flex:none;letter-spacing:0.02em;";
    var tagStyle, msgColor, rowBg = "transparent";
    if (event.level === "warning") {
      tagStyle = tagBase + "color:oklch(0.78 0.13 55);background:oklch(0.55 0.14 62 / 0.18);";
      msgColor = "oklch(0.82 0.09 60)";
      rowBg = "oklch(0.55 0.14 62 / 0.09)";
    } else if (event.level === "error") {
      tagStyle = tagBase + "color:oklch(0.8 0.12 25);background:oklch(0.55 0.18 25 / 0.2);";
      msgColor = "oklch(0.82 0.11 25)";
      rowBg = "oklch(0.55 0.18 25 / 0.12)";
    } else {
      tagStyle = tagBase + "color:#8B9097;background:rgba(255,255,255,0.06);";
      msgColor = "#C9CDD3";
    }
    li.style.background = rowBg;

    var time = document.createElement("span");
    time.style.cssText =
      "font-family:'IBM Plex Mono',monospace;font-size:9.5px;color:#5C6066;flex:none;width:50px;";
    time.textContent = timeOf(event.timestamp);

    var tagEl = document.createElement("span");
    tagEl.style.cssText = tagStyle;
    tagEl.textContent = tag;

    var msg = document.createElement("span");
    msg.style.cssText = "font-size:12.5px;line-height:1.4;color:" + msgColor + ";";
    msg.textContent = event.message;

    li.appendChild(time);
    li.appendChild(tagEl);
    li.appendChild(msg);
    els.log.appendChild(li);
    els.log.scrollTop = els.log.scrollHeight;
  }

  function setProgress(percent, mode) {
    els.percent.textContent = percent + " %";
    els.fill.style.width = percent + "%";
    var done = mode === "completed";
    els.fill.style.background = done ? GREEN_DOT : GREEN;
    els.percent.style.color = done ? GREEN_DARK : GREEN;
    if (mode === "completed") els.stage.textContent = "Abgeschlossen";
    else if (mode === "failed") els.stage.textContent = "Fehlgeschlagen";
    else if (mode === "cancelled") els.stage.textContent = "Abgebrochen";
    else els.stage.textContent = "Wird verarbeitet — " + (PHASE_LABELS[state.phaseIdx] || "Analyse");
  }

  function renderChips(mode) {
    els.chips.innerHTML = "";
    var completed = mode === "completed";
    for (var i = 0; i < PHASE_LABELS.length; i++) {
      var st;
      if (completed || state.phaseIdx > i) st = "done";
      else if (i === state.phaseIdx && mode === "running") st = "active";
      else st = "pending";
      els.chips.appendChild(buildChip(PHASE_LABELS[i], st));
    }
  }

  function buildChip(label, st) {
    var chip = document.createElement("div");
    var base =
      "display:inline-flex;align-items:center;gap:7px;padding:6px 11px;" +
      "border-radius:999px;font-size:11.5px;font-weight:500;";
    var dotBase =
      "width:15px;height:15px;border-radius:50%;display:inline-flex;align-items:center;" +
      "justify-content:center;font-size:9px;font-weight:700;color:#fff;flex:none;";
    var icon = "";
    if (st === "done") {
      chip.style.cssText = base + "background:oklch(0.52 0.13 145 / 0.13);color:" + GREEN_DARK + ";";
      var d1 = dot(dotBase + "background:" + GREEN_DOT + ";", "✓");
      chip.appendChild(d1);
    } else if (st === "active") {
      chip.style.cssText = base + "background:oklch(0.56 0.16 145 / 0.12);color:" + GREEN + ";";
      var d2 = dot(dotBase + "background:" + GREEN + ";animation:xlsPulse 1.4s ease-in-out infinite;", "");
      chip.appendChild(d2);
    } else {
      chip.style.cssText = base + "background:transparent;color:#9AA0A6;border:1px solid rgba(25,27,30,0.12);";
      chip.appendChild(dot(dotBase + "background:#C7CBD0;", icon));
    }
    var text = document.createElement("span");
    text.textContent = label;
    chip.appendChild(text);
    return chip;
  }

  function dot(style, icon) {
    var span = document.createElement("span");
    span.style.cssText = style;
    span.textContent = icon;
    return span;
  }

  function setStatus(mode) {
    var label, bg, color;
    if (mode === "completed") { label = "Fertig"; bg = "oklch(0.52 0.13 145 / 0.14)"; color = GREEN_DARK; }
    else if (mode === "failed") { label = "Fehlgeschlagen"; bg = "oklch(0.55 0.18 25 / 0.14)"; color = RED; }
    else if (mode === "cancelled") { label = "Abgebrochen"; bg = "oklch(0.55 0.14 62 / 0.14)"; color = RUST; }
    else { label = "Läuft"; bg = "oklch(0.56 0.16 145 / 0.12)"; color = GREEN; }
    els.status.textContent = label;
    els.status.style.cssText =
      "display:inline-flex;align-items:center;font-family:'IBM Plex Mono',monospace;" +
      "font-size:11px;font-weight:600;padding:5px 11px;border-radius:999px;background:" +
      bg + ";color:" + color + ";";
  }

  function finish(phase) {
    state.polling = false;
    hide(els.cancel);
    var mode = phase === "completed" ? "completed" : phase === "cancelled" ? "cancelled" : "failed";
    setStatus(mode);
    renderChips(mode);
    if (mode === "completed") {
      state.percent = 100;
      setProgress(100, "completed");
      loadReport();
    } else {
      setProgress(state.percent, mode);
    }
    show(els.reset);
  }

  function loadReport() {
    fetch("/api/jobs/" + state.jobId + "/report")
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (report) {
        if (report) renderReport(report);
        show(els.report);
      })
      .catch(function () { /* downloads still available below */ })
      .then(function () { renderDownloads(); });
  }

  function renderReport(report) {
    var rows = [
      ["Dauer", Number(report.duration_seconds || 0).toFixed(2) + " s", ""],
      ["Tabellen", deNumber(report.sheet_count || 0), ""],
      ["Zellen", deNumber(report.total_cells || 0), ""],
      ["Formeln", deNumber(report.total_formulas || 0), ""],
      ["Formel-Übereinst.", Number(report.formula_match_rate || 0).toFixed(2) + " %", GREEN_DARK],
      ["VBA-Module", deNumber(report.vba_modules || 0), ""],
      ["VBA-Prozeduren", deNumber(report.vba_procedures || 0), ""],
      ["Warnungen", deNumber(report.warnings_count || 0), report.warnings_count ? RUST : ""],
      ["Fehler", deNumber(report.errors_count || 0), report.errors_count ? RED : ""],
    ];
    els.reportGrid.innerHTML = "";
    rows.forEach(function (row) {
      var cell = document.createElement("div");
      cell.style.cssText = "background:#fff;padding:13px 14px;";
      var label = document.createElement("div");
      label.style.cssText =
        "font-family:'IBM Plex Mono',monospace;font-size:9.5px;letter-spacing:0.02em;" +
        "text-transform:uppercase;color:#9AA0A6;margin-bottom:5px;";
      label.textContent = row[0];
      var value = document.createElement("div");
      value.style.cssText = "font-size:16.5px;font-weight:700;letter-spacing:-0.01em;" + (row[2] ? "color:" + row[2] + ";" : "");
      value.textContent = row[1];
      cell.appendChild(label);
      cell.appendChild(value);
      els.reportGrid.appendChild(cell);
    });

    var warnings = report.warnings || [];
    if (warnings.length) {
      els.warningsList.innerHTML = "";
      els.warningsTitle.lastChild.textContent =
        warnings.length + " Warnung" + (warnings.length === 1 ? "" : "en") + " — manuelle Prüfung empfohlen";
      warnings.forEach(function (text) {
        var li = document.createElement("li");
        li.style.cssText = "font-size:12.5px;line-height:1.45;color:#4D5258;";
        li.textContent = text;
        els.warningsList.appendChild(li);
      });
      show(els.warnings);
    } else {
      hide(els.warnings);
    }
  }

  function renderDownloads() {
    var id = state.jobId;
    var odsName = (state.file ? state.file.name : "workbook").replace(/\.[^.]+$/, "") + ".ods";
    els.downloads.innerHTML = "";

    var ods = document.createElement("a");
    ods.href = "/jobs/" + id + "/download";
    ods.style.cssText =
      "display:inline-flex;align-items:center;gap:8px;text-decoration:none;font-size:13.5px;" +
      "font-weight:600;color:#fff;background:" + GREEN + ";border-radius:10px;padding:11px 16px;";
    var badge = document.createElement("span");
    badge.style.cssText =
      "font-family:'IBM Plex Mono',monospace;font-size:9.5px;font-weight:600;" +
      "background:rgba(255,255,255,0.2);border-radius:4px;padding:2px 6px;";
    badge.textContent = "ODS";
    ods.appendChild(badge);
    ods.appendChild(document.createTextNode(odsName));
    els.downloads.appendChild(ods);

    [["report.json", "/jobs/" + id + "/report.json"], ["report.md", "/jobs/" + id + "/report.md"]].forEach(function (pair) {
      var a = document.createElement("a");
      a.href = pair[1];
      a.textContent = pair[0];
      a.style.cssText =
        "display:inline-flex;align-items:center;text-decoration:none;font-size:13.5px;font-weight:500;" +
        "color:#16233B;background:#fff;border:1px solid rgba(25,27,30,0.16);border-radius:10px;" +
        "padding:11px 14px;font-family:'IBM Plex Mono',monospace;";
      els.downloads.appendChild(a);
    });
  }

  els.cancel.addEventListener("click", function () {
    if (!state.jobId) return;
    els.cancel.disabled = true;
    fetch("/api/jobs/" + state.jobId + "/cancel", { method: "POST" }).catch(function () {});
    // The cancellation surfaces as a terminal event through the normal poll.
  });

  els.reset.addEventListener("click", function () {
    state.polling = false;
    state.jobId = null;
    clearSampleCards();
    els.fileInput.value = "";
    setFile(null);
    hide(els.job);
    show(els.upload);
  });

  // Progressive enhancement: JS validates selection, so drop the native
  // required attribute and start from an empty selection.
  els.fileInput.removeAttribute("required");
  setFile(null);
})();
