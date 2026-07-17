const state = {
  clients: [],
  selectedClient: null,
};

const els = {
  clientSearch: document.querySelector("#clientSearch"),
  clearClient: document.querySelector("#clearClient"),
  clientOptions: document.querySelector("#clientOptions"),
  clientCombobox: document.querySelector(".client-combobox"),
  validFrom: document.querySelector("#validFrom"),
  validTo: document.querySelector("#validTo"),
  resetHeader: document.querySelector("#resetHeader"),
  outputDir: document.querySelector("#outputDir"),
  browseOutputDir: document.querySelector("#browseOutputDir"),
  importText: document.querySelector("#importText"),
  materialsBody: document.querySelector("#materialsBody"),
  importRows: document.querySelector("#importRows"),
  applyCustomsCode: document.querySelector("#applyCustomsCode"),
  addRow: document.querySelector("#addRow"),
  removeEmpty: document.querySelector("#removeEmpty"),
  clearTable: document.querySelector("#clearTable"),
  generateOutputIt: document.querySelector("#generateOutputIt"),
  generateOutputEn: document.querySelector("#generateOutputEn"),
  generateSideIt: document.querySelector("#generateSideIt"),
  generateSideEn: document.querySelector("#generateSideEn"),
  openPrintLog: document.querySelector("#openPrintLog"),
  printLogModal: document.querySelector("#printLogModal"),
  closePrintLog: document.querySelector("#closePrintLog"),
  printLogBody: document.querySelector("#printLogBody"),
  logStatus: document.querySelector("#logStatus"),
  exportLogExcel: document.querySelector("#exportLogExcel"),
  exportLogTxt: document.querySelector("#exportLogTxt"),
  status: document.querySelector("#status"),
};

let clientDropdownCloseTimer = null;

function cancelClientDropdownClose() {
  if (clientDropdownCloseTimer) {
    window.clearTimeout(clientDropdownCloseTimer);
    clientDropdownCloseTimer = null;
  }
}

function closeClientDropdown() {
  cancelClientDropdownClose();
  els.clientOptions.hidden = true;
}

function scheduleClientDropdownClose(delay = 180) {
  cancelClientDropdownClose();
  clientDropdownCloseTimer = window.setTimeout(() => {
    if (!els.clientCombobox.matches(":hover")) {
      closeClientDropdown();
    }
  }, delay);
}

function clientSearchText(client) {
  return `${client.name || ""} ${client.code || ""}`.toLowerCase();
}

function clientLabel(client) {
  const name = String(client.name || "").toUpperCase();
  return client.code ? `${client.code} - ${name}` : name;
}

function matchingClients(query) {
  const needle = query.trim().toLowerCase();
  if (!needle) return state.clients;
  return state.clients.filter((client) => clientSearchText(client).includes(needle));
}

function renderClients(query = els.clientSearch.value) {
  cancelClientDropdownClose();
  els.clientOptions.innerHTML = "";
  const matches = matchingClients(query);

  if (!matches.length) {
    closeClientDropdown();
    return;
  }

  matches.forEach((client) => {
    const option = document.createElement("button");
    option.type = "button";
    option.className = "client-option";
    option.setAttribute("role", "option");

    const code = document.createElement("span");
    code.className = "client-option-code";
    code.textContent = client.code || "";
    option.append(code);

    const name = document.createElement("strong");
    name.className = "client-option-name";
    name.textContent = String(client.name || "").toUpperCase();
    option.append(name);

    option.addEventListener("click", () => {
      state.selectedClient = client;
      els.clientSearch.value = clientLabel(client);
      closeClientDropdown();
    });

    els.clientOptions.append(option);
  });
  els.clientOptions.hidden = false;
}

async function loadClients() {
  const response = await fetch("clients.json");
  state.clients = (await response.json()).sort((a, b) => {
    return String(a.name || "").localeCompare(String(b.name || ""), "it", { sensitivity: "base" });
  });
}

function resolveClient() {
  const value = els.clientSearch.value.trim().toLowerCase();
  if (state.selectedClient && clientLabel(state.selectedClient).toLowerCase() === value) {
    return state.selectedClient;
  }
  return state.clients.find((client) => {
    return client.name.toLowerCase() === value || clientLabel(client).toLowerCase() === value || String(client.code || "").toLowerCase() === value;
  }) || null;
}

function escapeAttr(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function rowTemplate(row = {}) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input data-field="code" value="${escapeAttr(row.code || "")}" placeholder="Codice articolo"></td>
    <td><input data-field="description" value="${escapeAttr(row.description || "")}" placeholder="Descrizione"></td>
    <td><input data-field="customsCode" value="${escapeAttr(row.customsCode || "")}" placeholder="Codice doganale"></td>
    <td class="asterisk-cell">*</td>
    <td><button class="icon secondary" type="button" title="Elimina riga" aria-label="Elimina riga">
      <svg viewBox="0 0 24 24"><path d="M3 6h18"></path><path d="M8 6V4h8v2"></path><path d="M19 6l-1 14H6L5 6"></path></svg>
    </button></td>
  `;
  tr.querySelector("button").addEventListener("click", () => tr.remove());
  return tr;
}

function addRow(row) {
  els.materialsBody.append(rowTemplate(row));
}

function replaceRows(rows) {
  els.materialsBody.innerHTML = "";
  rows.forEach((row) => addRow(row));
}

function parseImportText(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !/^codice\b/i.test(line) && !/^articolo\b/i.test(line))
    .map((line) => {
      const normalized = line.replace(/\s*[;|]\s*/g, "\t");
      const parts = normalized.split(/\t+/).map((part) => part.trim()).filter(Boolean);
      if (parts.length >= 2) {
        return { code: parts[0], description: parts.slice(1).join(" ") };
      }
      const match = line.match(/^(\S+)\s+(.+)$/);
      return {
        code: match ? match[1] : line,
        description: match ? match[2].trim() : "",
      };
    });
}

function collectMaterials() {
  return [...els.materialsBody.querySelectorAll("tr")].map((tr) => {
    const values = {};
    tr.querySelectorAll("input").forEach((input) => {
      values[input.dataset.field] = input.value.trim();
    });
    return values;
  });
}

function applyCustomsCodeToRows() {
  const rows = [...els.materialsBody.querySelectorAll("tr")];
  if (!rows.length) {
    setStatus("Inserisci prima almeno una riga in tabella.", "err");
    return;
  }
  const value = window.prompt("Codice doganale da applicare a tutte le righe:", "");
  if (value === null) return;
  rows.forEach((tr) => {
    const input = tr.querySelector('input[data-field="customsCode"]');
    if (input) input.value = value.trim();
  });
  setStatus("Codice doganale applicato a tutte le righe.", "ok");
}

function payload(language) {
  return {
    cliente: resolveClient(),
    language,
    validFrom: els.validFrom.value,
    validTo: els.validTo.value,
    outputDir: els.outputDir.value.trim(),
    materials: collectMaterials(),
  };
}

function setStatus(text, type = "") {
  els.status.textContent = text;
  els.status.className = type;
  els.status.hidden = !text;
}

function autoResizeImportText() {
  els.importText.style.height = "auto";
  els.importText.style.height = `${els.importText.scrollHeight}px`;
}

async function generatePdf(language) {
  setStatus("Generazione in corso...");
  try {
    if (!resolveClient()) throw new Error("Seleziona un cliente dall'elenco filtrabile.");
    if (!els.validFrom.value || !els.validTo.value) throw new Error("Inserisci data inizio validita e data fine validita.");
    if (!els.outputDir.value.trim()) throw new Error("Inserisci la cartella destinazione PDF.");

    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload(language)),
    });
    const result = await response.json();
    if (!result.ok) throw new Error(result.error || "Errore durante la generazione.");
    setStatus(`PDF creato: ${result.path}`, "ok");
  } catch (error) {
    setStatus(error.message, "err");
  }
}

async function browseOutputFolder() {
  els.browseOutputDir.disabled = true;
  setStatus("Seleziona la cartella nella finestra di Esplora risorse...");
  try {
    const response = await fetch("/api/select-folder", { method: "POST" });
    const result = await response.json();
    if (!result.ok) throw new Error(result.error || "Impossibile selezionare la cartella.");
    if (result.path) {
      els.outputDir.value = result.path;
      setStatus(`Cartella selezionata: ${result.path}`, "ok");
    } else {
      setStatus("");
    }
  } catch (error) {
    setStatus(error.message, "err");
  } finally {
    els.browseOutputDir.disabled = false;
  }
}

function logCell(row, key) {
  const td = document.createElement("td");
  td.textContent = row[key] || "";
  return td;
}

async function openPrintLog() {
  els.printLogModal.hidden = false;
  els.logStatus.textContent = "Caricamento registro...";
  els.printLogBody.innerHTML = "";
  try {
    const response = await fetch("/api/logs");
    const result = await response.json();
    if (!result.ok) throw new Error(result.error || "Impossibile leggere il registro.");
    result.rows.forEach((row) => {
      const tr = document.createElement("tr");
      tr.append(
        logCell(row, "data_ora"),
        logCell(row, "utente"),
        logCell(row, "lingua"),
        logCell(row, "codice_cliente"),
        logCell(row, "cliente"),
      );
      const validity = document.createElement("td");
      validity.textContent = `${row.validita_da || ""} - ${row.validita_a || ""}`;
      tr.append(validity, logCell(row, "articoli"), logCell(row, "file_pdf"));
      els.printLogBody.append(tr);
    });
    els.logStatus.textContent = result.rows.length ? `${result.rows.length} attestazioni registrate.` : "Nessuna attestazione registrata.";
  } catch (error) {
    els.logStatus.textContent = error.message;
  }
}

function closePrintLog() {
  els.printLogModal.hidden = true;
}

els.addRow.addEventListener("click", () => addRow());
els.clientCombobox.addEventListener("mouseenter", () => {
  cancelClientDropdownClose();
  renderClients();
});
els.clientCombobox.addEventListener("mouseleave", () => scheduleClientDropdownClose());
els.clientOptions.addEventListener("mouseenter", cancelClientDropdownClose);
els.clientSearch.addEventListener("pointerdown", () => renderClients());
els.clientSearch.addEventListener("focus", () => renderClients());
els.clientSearch.addEventListener("input", () => {
  state.selectedClient = null;
  renderClients();
});
els.clientSearch.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    const firstOption = els.clientOptions.querySelector(".client-option");
    if (firstOption) {
      event.preventDefault();
      firstOption.click();
    }
  }
  if (event.key === "Escape") {
    closeClientDropdown();
  }
});
els.clientCombobox.addEventListener("focusout", (event) => {
  if (!els.clientCombobox.contains(event.relatedTarget)) {
    scheduleClientDropdownClose(120);
  }
});
document.addEventListener("pointerdown", (event) => {
  if (!els.clientCombobox.contains(event.target)) {
    closeClientDropdown();
  }
}, true);
els.clearClient.addEventListener("click", () => {
  state.selectedClient = null;
  els.clientSearch.value = "";
  closeClientDropdown();
  els.clientSearch.focus();
});
els.resetHeader.addEventListener("click", () => {
  state.selectedClient = null;
  els.clientSearch.value = "";
  closeClientDropdown();
  els.validFrom.value = "";
  els.validTo.value = "";
  setStatus("Filtri testata resettati.", "ok");
});
els.importRows.addEventListener("click", () => {
  const rows = parseImportText(els.importText.value);
  if (!rows.length) {
    setStatus("Incolla almeno una riga con codice articolo e descrizione.", "err");
    return;
  }
  replaceRows(rows);
  els.importText.value = "";
  autoResizeImportText();
  setStatus(`${rows.length} righe importate in tabella.`, "ok");
});
els.importText.addEventListener("input", autoResizeImportText);
els.removeEmpty.addEventListener("click", () => {
  [...els.materialsBody.querySelectorAll("tr")].forEach((tr) => {
    const isEmpty = [...tr.querySelectorAll("input")].every((input) => !input.value.trim());
    if (isEmpty) tr.remove();
  });
});
els.clearTable.addEventListener("click", () => {
  replaceRows([]);
  setStatus("Tabella cancellata.", "ok");
});
els.applyCustomsCode.addEventListener("click", applyCustomsCodeToRows);
els.generateSideIt.addEventListener("click", () => generatePdf("it"));
els.generateSideEn.addEventListener("click", () => generatePdf("en"));
els.generateOutputIt.addEventListener("click", () => generatePdf("it"));
els.generateOutputEn.addEventListener("click", () => generatePdf("en"));
els.browseOutputDir.addEventListener("click", browseOutputFolder);
els.openPrintLog.addEventListener("click", openPrintLog);
els.closePrintLog.addEventListener("click", closePrintLog);
els.printLogModal.addEventListener("pointerdown", (event) => {
  if (event.target === els.printLogModal) closePrintLog();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !els.printLogModal.hidden) closePrintLog();
});
els.exportLogExcel.addEventListener("click", () => {
  window.location.href = "/api/export-log?format=xlsx";
});
els.exportLogTxt.addEventListener("click", () => {
  window.location.href = "/api/export-log?format=txt";
});

loadClients();
autoResizeImportText();
