const state = {
  missionId: null,
  missionDirHandle: null,
  frames: [], // [{ frameId, frameNumber, xmlFile, jpgFile, metadata, kmzName }]
  centerOfMission: null, // "lat,lon" string or null
  combinedKmzName: null,
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

function frameNumberFromName(name) {
  const m = name.match(/^(\d+)/);
  return m ? parseInt(m[1], 10) : null;
}

function getMetadataField(metadata, fieldName) {
  const target = fieldName.toLowerCase();
  for (const [key, value] of Object.entries(metadata)) {
    const lastSegment = key.split(".").pop().replace(/^@/, "").toLowerCase();
    if (lastSegment === target) return value;
  }
  return null;
}

function computeCenterOfMission(frames) {
  const numbered = frames.filter((f) => f.frameNumber !== null);
  if (numbered.length < 1) return null;

  const sorted = [...numbered].sort((a, b) => a.frameNumber - b.frameNumber);
  const minFrame = sorted[0];
  const maxFrame = sorted[sorted.length - 1];

  const corners = [
    [getMetadataField(minFrame.metadata, "UpperLeftLatitude"), getMetadataField(minFrame.metadata, "UpperLeftLongitude")],
    [getMetadataField(minFrame.metadata, "UpperRightLatitude"), getMetadataField(minFrame.metadata, "UpperRightLongitude")],
    [getMetadataField(maxFrame.metadata, "LowerLeftLatitude"), getMetadataField(maxFrame.metadata, "LowerLeftLongitude")],
    [getMetadataField(maxFrame.metadata, "LowerRightLatitude"), getMetadataField(maxFrame.metadata, "LowerRightLongitude")],
  ];

  const lats = corners.map(([la]) => parseFloat(la));
  const lons = corners.map(([, lo]) => parseFloat(lo));
  if (lats.some(isNaN) || lons.some(isNaN)) return null;

  const lat = lats.reduce((a, b) => a + b, 0) / 4;
  const lon = lons.reduce((a, b) => a + b, 0) / 4;
  return { lat, lon };
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

function buildPlacemark(name, imgDataUri, metadata) {
  const { lat = 0, lon = 0 } = findCoordinates(metadata);
  const fields = Object.entries(metadata)
    .map(([k, v]) => `<b>${escapeXml(k)}:</b> ${escapeXml(v)}`)
    .join("<br/>");
  const img = imgDataUri ? `<img src="${imgDataUri}" width="400" /><br/>` : "";
  const description = `<![CDATA[${img}${fields}]]>`;
  return `    <Placemark>
      <name>${escapeXml(name)}</name>
      <description>${description}</description>
      <Point>
        <coordinates>${lon},${lat},0</coordinates>
      </Point>
    </Placemark>`;
}

function wrapKml(docName, placemarksXml) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>${escapeXml(docName)}</name>
${placemarksXml}
  </Document>
</kml>`;
}

function fileToDataUri(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
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
    frameNumber: frameNumberFromName(subDirHandle.name),
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
    state.centerOfMission = null;
    state.combinedKmzName = null;

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

    state.frames.sort((a, b) => (a.frameNumber ?? Infinity) - (b.frameNumber ?? Infinity));
    const center = computeCenterOfMission(state.frames);
    state.centerOfMission = center ? `${center.lat.toFixed(8)},${center.lon.toFixed(8)}` : null;

    let msg = `Mission: ${state.missionId} — ${state.frames.length} frame(s) detected.`;
    if (state.centerOfMission) msg += ` Center: ${state.centerOfMission}.`;
    else msg += ` Center: not computed (missing corner fields).`;
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

    const placemarks = [];
    let count = 0;
    for (const frame of state.frames) {
      const dataUri = await fileToDataUri(frame.jpgFile);
      const placemark = buildPlacemark(frame.frameId, dataUri, frame.metadata);
      placemarks.push(placemark);

      const kml = wrapKml(`${state.missionId}/${frame.frameId}`, placemark);
      const fileName = `${state.missionId}_${frame.frameId}.kml`;
      const fileHandle = await outDirHandle.getFileHandle(fileName, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(new Blob([kml], { type: "application/vnd.google-earth.kml+xml" }));
      await writable.close();

      frame.kmzName = fileName;
      count++;
      setStatus(`Created ${count}/${state.frames.length} KML files…`, "ok");
    }

    let centerPlacemark = "";
    if (state.centerOfMission) {
      const [clat, clon] = state.centerOfMission.split(",");
      centerPlacemark = `    <Placemark>
      <name>Center of Mission</name>
      <Point><coordinates>${clon},${clat},0</coordinates></Point>
    </Placemark>\n`;
    }

    const combinedName = `${state.missionId}_combined.kml`;
    const combinedKml = wrapKml(state.missionId, centerPlacemark + placemarks.join("\n"));
    const combinedHandle = await outDirHandle.getFileHandle(combinedName, { create: true });
    const cw = await combinedHandle.createWritable();
    await cw.write(new Blob([combinedKml], { type: "application/vnd.google-earth.kml+xml" }));
    await cw.close();
    state.combinedKmzName = combinedName;

    setInfo(els.kmlInfo, `Saved ${count} per-frame KML(s) and combined "${combinedName}" into "${outDirHandle.name}".`, "ok");
    setStatus("All KMLs created.", "ok");
  } catch (err) {
    if (err.name === "AbortError") return;
    setStatus("Failed to create KMZs: " + err.message, "err");
  }
});

const BASE_COLUMNS = ["MissionID", "FrameID", "centerOfMission", "kml_path", "created_at"];

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
      centerOfMission: state.centerOfMission || "",
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
