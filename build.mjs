import * as XLSX from "xlsx";
import { readFileSync, writeFileSync } from "fs";

const SRC = "team_task_tracker.xlsx";
const OUT = "data.json";

function excelSerialToISO(serial) {
  if (!serial && serial !== 0) return "";
  if (typeof serial === "string") {
    const d = new Date(serial);
    return isNaN(d.getTime()) ? "" : d.toISOString().split("T")[0];
  }
  if (typeof serial === "number" && serial > 1) {
    const epoch = Date.UTC(1900, 0, 1);
    const ms = (serial - 2) * 86400000 + epoch;
    return new Date(ms).toISOString().split("T")[0];
  }
  return "";
}

const buf = readFileSync(SRC);
const wb = XLSX.read(buf, { type: "buffer" });
const ws = wb.Sheets["Task Tracker"];
if (!ws) throw new Error('Sheet "Task Tracker" not found');

const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const headerIdx = raw.findIndex(
  (r) => r && r.some((c) => String(c).trim().toLowerCase() === "task id")
);
if (headerIdx === -1) throw new Error("Header row with 'Task ID' not found");

const header = raw[headerIdx].map((h) => String(h ?? "").trim().toLowerCase());
const rows = raw.slice(headerIdx + 1);

const idx = (key) => header.indexOf(key);

const data = rows
  .filter((r) => {
    const taskId = String(r[idx("task id")] ?? "").trim();
    return taskId.length > 0 && taskId.toLowerCase().startsWith("tsk");
  })
  .map((r) => ({
    id: String(r[idx("task id")] ?? "").trim(),
    name: String(r[idx("task name")] ?? "").trim(),
    assignee: String(r[idx("assigned to")] ?? "").trim(),
    brand: String(r[idx("brand")] ?? "").trim(),
    type: String(r[idx("task type")] ?? "").trim(),
    priority: String(r[idx("priority")] ?? "").trim(),
    status: String(r[idx("status")] ?? "").trim(),
    start: excelSerialToISO(r[idx("start date")]),
    deadline: excelSerialToISO(r[idx("deadline")]),
    completed: excelSerialToISO(r[idx("completion date")]),
    notes: String(r[idx("notes")] ?? "").trim(),
  }));

writeFileSync(OUT, JSON.stringify(data, null, 2));
console.log(`Built ${OUT} with ${data.length} tasks`);
