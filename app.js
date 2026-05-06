const state = {
  folderName: null,
  xmlFile: null,
  jpgFile: null,
  metadata: {},
  kmzPath: null,
  kmzName: null,
};

const els = {
  selectFolder: document.getElementById("btn-select-folder"),
  createKml: document.getElementById("btn-create-kml"),
  importExcel: document.getElementById("btn-import-excel"),
  folderInfo: document.getElementById("folder-info"),
  kmlInfo: document.getElementById("kml-info"),
  excelInfo: document.getElementById("excel-info"),
  status: document.getElementById("status"),
};

function setStatus(msg, kind = "") {
  els.status.textContent = msg;
  els.status.className = "status" + (kind ? " " + kind : "");
}

function setInfo(el, msg, kind = "") {
  el.textContent = msg;
  el.className = "info" + (kind ? " " + kind : "");
}

function checkApiSupport() {
  if (!window.showDirectoryPicker || !window.showSaveFilePicker || !window.showOpenFilePicker) {
    setStatus("This app requires Chrome or Edge (File System Access API).", "err");
    els.selectFolder.disabled = true;
    return false;
  }
  return true;
}

// Stub: parse XML and return metadata object. Will be filled in once
// the user provides an example XML schema.
function parseMetadata(xmlText) {
  return {};
}

els.selectFolder.addEventListener("click", async () => {
  try {
    const dirHandle = await window.showDirectoryPicker();
    state.folderName = dirHandle.name;

    let xmlHandle = null;
    let jpgHandle = null;

    for await (const [name, handle] of dirHandle.entries()) {
      if (handle.kind !== "file") continue;
      const lower = name.toLowerCase();
      if (lower.endsWith(".xml")) xmlHandle = handle;
      else if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) jpgHandle = handle;
    }

    if (!xmlHandle || !jpgHandle) {
      setInfo(els.folderInfo, "Folder must contain one .xml and one .jpg file.", "err");
      els.createKml.disabled = true;
      els.importExcel.disabled = true;
      return;
    }

    state.xmlFile = await xmlHandle.getFile();
    state.jpgFile = await jpgHandle.getFile();
    const xmlText = await state.xmlFile.text();
    state.metadata = parseMetadata(xmlText);

    setInfo(
      els.folderInfo,
      `Folder: ${state.folderName} — XML: ${state.xmlFile.name}, JPG: ${state.jpgFile.name}`,
      "ok"
    );
    els.createKml.disabled = false;
    els.importExcel.disabled = false;
    setStatus("Folder loaded.", "ok");
  } catch (err) {
    if (err.name === "AbortError") return;
    setStatus("Failed to read folder: " + err.message, "err");
  }
});

function buildKml(jpgFileName, metadata) {
  const lat = metadata.latitude ?? 0;
  const lon = metadata.longitude ?? 0;
  const name = metadata.name ?? state.folderName ?? "Placemark";
  const description = `<![CDATA[<img src="${jpgFileName}" width="400" />]]>`;

  return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>${escapeXml(name)}</name>
    <Placemark>
      <name>${escapeXml(name)}</name>
      <description>${description}</description>
      <Point>
        <coordinates>${lon},${lat},0</coordinates>
      </Point>
    </Placemark>
  </Document>
</kml>`;
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

els.createKml.addEventListener("click", async () => {
  try {
    if (!state.xmlFile || !state.jpgFile) {
      setStatus("Select a folder first.", "err");
      return;
    }

    const jpgName = state.jpgFile.name;
    const kmlText = buildKml(jpgName, state.metadata);

    const zip = new JSZip();
    zip.file("doc.kml", kmlText);
    zip.file(jpgName, await state.jpgFile.arrayBuffer());
    const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.google-earth.kmz" });

    const suggestedName = (state.folderName || "archive") + ".kmz";
    const handle = await window.showSaveFilePicker({
      suggestedName,
      types: [{ description: "KMZ file", accept: { "application/vnd.google-earth.kmz": [".kmz"] } }],
    });

    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();

    state.kmzName = handle.name;
    state.kmzPath = handle.name; // browsers do not expose absolute path; we keep the file name
    setInfo(els.kmlInfo, `Saved as ${handle.name}`, "ok");
    setStatus("KMZ created.", "ok");
  } catch (err) {
    if (err.name === "AbortError") return;
    setStatus("Failed to create KMZ: " + err.message, "err");
  }
});

const EMPTY_COLUMNS = ["folder", "xml_file", "jpg_file", "kml_path", "created_at"];

els.importExcel.addEventListener("click", async () => {
  try {
    const [fileHandle] = await window.showOpenFilePicker({
      types: [{ description: "Excel file", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
    });

    const file = await fileHandle.getFile();
    const buf = await file.arrayBuffer();

    let workbook;
    let sheetName;
    let rows;

    if (file.size === 0) {
      workbook = XLSX.utils.book_new();
      sheetName = "Sheet1";
      rows = [];
    } else {
      workbook = XLSX.read(buf, { type: "array" });
      sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    }

    const newRow = {
      folder: state.folderName || "",
      xml_file: state.xmlFile ? state.xmlFile.name : "",
      jpg_file: state.jpgFile ? state.jpgFile.name : "",
      kml_path: state.kmzName || "",
      created_at: new Date().toISOString(),
      ...state.metadata,
    };

    rows.push(newRow);

    const headers = rows.length > 0
      ? Array.from(new Set([...EMPTY_COLUMNS, ...Object.keys(newRow), ...rows.flatMap((r) => Object.keys(r))]))
      : EMPTY_COLUMNS;

    const newSheet = XLSX.utils.json_to_sheet(rows, { header: headers });

    if (state.kmzName) {
      const kmlPathCol = headers.indexOf("kml_path");
      if (kmlPathCol >= 0) {
        const rowIdx = rows.length; // 1-based header + rows.length = last row index
        const cellRef = XLSX.utils.encode_cell({ c: kmlPathCol, r: rowIdx });
        if (newSheet[cellRef]) {
          newSheet[cellRef].l = { Target: state.kmzName, Tooltip: "Open KMZ in default app (Google Earth)" };
        }
      }
    }

    workbook.Sheets[sheetName] = newSheet;
    if (!workbook.SheetNames.includes(sheetName)) workbook.SheetNames.push(sheetName);

    const out = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const writable = await fileHandle.createWritable();
    await writable.write(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    await writable.close();

    setInfo(els.excelInfo, `Updated ${fileHandle.name} — ${rows.length} row(s).`, "ok");
    setStatus("Excel updated.", "ok");
  } catch (err) {
    if (err.name === "AbortError") return;
    setStatus("Failed to update Excel: " + err.message, "err");
  }
});

checkApiSupport();
