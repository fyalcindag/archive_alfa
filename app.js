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

function parseMetadata(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  const parseError = doc.querySelector("parsererror");
  if (parseError) throw new Error("Invalid XML: " + parseError.textContent.split("\n")[0]);

  const out = {};
  const counters = new Map();

  function visit(node, path) {
    for (const attr of node.attributes || []) {
      const key = path ? `${path}.@${attr.name}` : `@${attr.name}`;
      assign(out, key, attr.value);
    }

    const elementChildren = Array.from(node.children);
    if (elementChildren.length === 0) {
      const text = (node.textContent || "").trim();
      if (text) assign(out, path, text);
      return;
    }

    const siblingTags = elementChildren.map((c) => c.tagName);
    for (const child of elementChildren) {
      const tag = child.tagName;
      const repeated = siblingTags.filter((t) => t === tag).length > 1;
      let segment = tag;
      if (repeated) {
        const counterKey = path + "/" + tag;
        const i = counters.get(counterKey) ?? 0;
        counters.set(counterKey, i + 1);
        segment = `${tag}[${i}]`;
      }
      visit(child, path ? `${path}.${segment}` : segment);
    }
  }

  function assign(target, key, value) {
    if (target[key] === undefined) target[key] = value;
    else target[key] += " | " + value;
  }

  if (doc.documentElement) visit(doc.documentElement, doc.documentElement.tagName);
  return out;
}

function findCoordinates(metadata) {
  let lat, lon;
  for (const [key, value] of Object.entries(metadata)) {
    const k = key.toLowerCase();
    const num = parseFloat(value);
    if (Number.isNaN(num)) continue;
    if (lat === undefined && /(^|[^a-z])lat(itude)?($|[^a-z])/.test(k)) lat = num;
    else if (lon === undefined && /(^|[^a-z])(lon|lng|longitude)($|[^a-z])/.test(k)) lon = num;
  }
  return { lat, lon };
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
  const { lat = 0, lon = 0 } = findCoordinates(metadata);
  const name = state.folderName || "Placemark";
  const fields = Object.entries(metadata)
    .map(([k, v]) => `<b>${escapeXml(k)}:</b> ${escapeXml(v)}`)
    .join("<br/>");
  const description = `<![CDATA[<img src="${jpgFileName}" width="400" /><br/>${fields}]]>`;

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
