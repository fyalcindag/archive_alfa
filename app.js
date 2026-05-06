const state = {
  missionId: null,
  missionDirHandle: null,
  frames: [], // [{ frameId, xmlFile, jpgFile, metadata, kmzName }]
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
  if (!window.showDirectoryPicker || !window.showOpenFilePicker) {
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

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function buildKml(name, jpgFileName, metadata) {
  const { lat = 0, lon = 0 } = findCoordinates(metadata);
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

async function readFrame(subDirHandle) {
  let xmlHandle = null;
  let jpgHandle = null;

  for await (const [name, handle] of subDirHandle.entries()) {
    if (handle.kind !== "file") continue;
    const lower = name.toLowerCase();
    if (lower.endsWith("meta.xml")) xmlHandle = handle;
    else if ((lower.endsWith(".jpg") || lower.endsWith(".jpeg")) && !jpgHandle) jpgHandle = handle;
  }

  if (!xmlHandle || !jpgHandle) return null;

  const xmlFile = await xmlHandle.getFile();
  const jpgFile = await jpgHandle.getFile();
  const xmlText = await xmlFile.text();
  const metadata = parseMetadata(xmlText);

  return {
    frameId: subDirHandle.name,
    xmlFile,
    jpgFile,
    metadata,
    kmzName: null,
  };
}

els.selectFolder.addEventListener("click", async () => {
  try {
    const dirHandle = await window.showDirectoryPicker();
    state.missionDirHandle = dirHandle;
    state.missionId = dirHandle.name;
    state.frames = [];

    const skipped = [];
    for await (const [name, handle] of dirHandle.entries()) {
      if (handle.kind !== "directory") continue;
      const frame = await readFrame(handle);
      if (frame) state.frames.push(frame);
      else skipped.push(name);
    }

    if (state.frames.length === 0) {
      setInfo(els.folderInfo, `No valid frames found in "${state.missionId}". Each subfolder must contain a JPG and a *meta.xml file.`, "err");
      els.createKml.disabled = true;
      els.importExcel.disabled = true;
      return;
    }

    let msg = `Mission: ${state.missionId} — ${state.frames.length} frame(s) detected.`;
    if (skipped.length) msg += ` Skipped: ${skipped.join(", ")}`;
    setInfo(els.folderInfo, msg, "ok");
    els.createKml.disabled = false;
    els.importExcel.disabled = false;
    setStatus("Mission folder loaded.", "ok");
  } catch (err) {
    if (err.name === "AbortError") return;
    setStatus("Failed to read folder: " + err.message, "err");
  }
});

els.createKml.addEventListener("click", async () => {
  try {
    if (state.frames.length === 0) {
      setStatus("Select a mission folder first.", "err");
      return;
    }

    const outDirHandle = await window.showDirectoryPicker({ mode: "readwrite" });

    let count = 0;
    for (const frame of state.frames) {
      const kml = buildKml(`${state.missionId}/${frame.frameId}`, frame.jpgFile.name, frame.metadata);
      const zip = new JSZip();
      zip.file("doc.kml", kml);
      zip.file(frame.jpgFile.name, await frame.jpgFile.arrayBuffer());
      const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.google-earth.kmz" });

      const fileName = `${state.missionId}_${frame.frameId}.kmz`;
      const fileHandle = await outDirHandle.getFileHandle(fileName, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();

      frame.kmzName = fileName;
      count++;
      setStatus(`Created ${count}/${state.frames.length} KMZ files…`, "ok");
    }

    setInfo(els.kmlInfo, `Saved ${count} KMZ file(s) into "${outDirHandle.name}".`, "ok");
    setStatus("All KMZs created.", "ok");
  } catch (err) {
    if (err.name === "AbortError") return;
    setStatus("Failed to create KMZs: " + err.message, "err");
  }
});

const BASE_COLUMNS = ["MissionID", "FrameID", "kml_path", "created_at"];

els.importExcel.addEventListener("click", async () => {
  try {
    if (state.frames.length === 0) {
      setStatus("Select a mission folder first.", "err");
      return;
    }

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

    const newRows = state.frames.map((frame) => ({
      MissionID: state.missionId,
      FrameID: frame.frameId,
      kml_path: frame.kmzName || "",
      created_at: new Date().toISOString(),
      ...frame.metadata,
    }));

    const allRows = [...rows, ...newRows];
    const headers = Array.from(
      new Set([...BASE_COLUMNS, ...allRows.flatMap((r) => Object.keys(r))])
    );

    const newSheet = XLSX.utils.json_to_sheet(allRows, { header: headers });

    const kmlPathCol = headers.indexOf("kml_path");
    if (kmlPathCol >= 0) {
      for (let i = 0; i < newRows.length; i++) {
        const rowIdx = rows.length + i + 1; // +1 for header row
        const cellRef = XLSX.utils.encode_cell({ c: kmlPathCol, r: rowIdx });
        const linkTarget = newRows[i].kml_path;
        if (newSheet[cellRef] && linkTarget) {
          newSheet[cellRef].l = { Target: linkTarget, Tooltip: "Open KMZ in default app (Google Earth)" };
        }
      }
    }

    workbook.Sheets[sheetName] = newSheet;
    if (!workbook.SheetNames.includes(sheetName)) workbook.SheetNames.push(sheetName);

    const out = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const writable = await fileHandle.createWritable();
    await writable.write(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    await writable.close();

    setInfo(els.excelInfo, `Updated ${fileHandle.name} — added ${newRows.length} row(s), total ${allRows.length}.`, "ok");
    setStatus("Excel updated.", "ok");
  } catch (err) {
    if (err.name === "AbortError") return;
    setStatus("Failed to update Excel: " + err.message, "err");
  }
});

checkApiSupport();
