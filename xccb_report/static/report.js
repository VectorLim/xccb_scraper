"use strict";

const element = (id) => document.getElementById(id);
const form = element("report-form");
const names = element("names");
const search = element("search");
const dateMode = element("date-mode");
const date = element("date");
const resultFilter = element("result-filter");
const jobStatus = element("job-status");
const previous = sessionStorage.getItem("job");
let active = Boolean(previous);
let available = false;
let result;

for (const field of [names, search, dateMode, resultFilter]) {
  const saved = sessionStorage.getItem(field.id);
  if (saved !== null) field.value = saved;
  field.addEventListener("input", () => sessionStorage.setItem(field.id, field.value));
}

async function api(url, options) {
  const response = await fetch(url, options);
  if (response.ok) return response.json();
  const payload = await response.json().catch(() => ({}));
  const error = new Error(payload.error || "The report service is unavailable.");
  error.status = response.status;
  throw error;
}

function updateButton() {
  element("process").disabled = active || !available;
}

async function refreshAvailability() {
  try {
    const status = await api("/api/status");
    available = status.enabled && status.worker_online;
    element("availability").textContent = available
      ? "Windows worker online"
      : status.enabled ? "Windows worker offline" : "Live reporting is disabled";
  } catch {
    available = false;
    element("availability").textContent = "Report service unavailable";
  }
  updateButton();
}

function table(id, headers, rows, linkColumn = -1) {
  const target = element(id);
  target.replaceChildren();
  const head = target.createTHead().insertRow();
  for (const heading of headers) {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = heading;
    head.append(cell);
  }
  const body = target.createTBody();
  for (const row of rows) {
    const line = body.insertRow();
    row.forEach((value, index) => {
      const cell = line.insertCell();
      if (index === linkColumn && value) {
        const link = document.createElement("a");
        link.href = value;
        link.textContent = "Open workspace";
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        cell.append(link);
      } else {
        cell.textContent = value ?? "";
      }
    });
  }
}

function dateKey(row) {
  const match = /^\d{4}-\d{2}-\d{2}/.exec(row[6]);
  return match ? match[0].slice(0, dateMode.value === "month" ? 7 : 10) : "";
}

function populateDates(selected) {
  const values = [...new Set(result.documents.map(dateKey).filter(Boolean))].sort().reverse();
  date.replaceChildren(new Option("All dates", ""));
  for (const value of values) date.add(new Option(value, value));
  date.value = values.includes(selected) ? selected : "";
}

function filterDocuments() {
  if (!result) return;
  const query = search.value.trim().toLowerCase();
  const rows = result.documents.filter((row) =>
    (!date.value || dateKey(row) === date.value)
    && (resultFilter.value === "all" || row[0] === resultFilter.value)
    && (!query || row.some((value) => String(value).toLowerCase().includes(query)))
  );
  table("documents", result.document_headers, rows, 9);
  element("visible-count").textContent = rows.length + " of " + result.documents.length + " submissions";
}

search.addEventListener("input", filterDocuments);
resultFilter.addEventListener("change", filterDocuments);
date.addEventListener("change", () => {
  sessionStorage.setItem("submitted-date", date.value);
  filterDocuments();
});
dateMode.addEventListener("change", () => {
  sessionStorage.removeItem("submitted-date");
  if (result) populateDates("");
  filterDocuments();
});

async function follow(id) {
  const url = "/api/jobs/" + id;
  try {
    const job = await api(url);
    if (job.state === "queued" || job.state === "running") {
      jobStatus.textContent = job.state === "queued"
        ? "Queued" : "Processing · " + job.processed + " submissions checked";
      setTimeout(() => follow(id), 2000);
      return;
    }
    jobStatus.textContent = job.message;
    if (job.state === "complete") {
      result = job.result;
      element("quality").textContent = result.status;
      element("excel").href = url + "/excel";
      populateDates(sessionStorage.getItem("submitted-date"));
      filterDocuments();
      table("managers", result.manager_headers, result.managers);
      element("results").hidden = false;
    } else {
      sessionStorage.removeItem("job");
    }
  } catch (error) {
    jobStatus.textContent = error.message;
    if (error.status !== 404) {
      setTimeout(() => follow(id), 5000);
      return;
    }
    sessionStorage.removeItem("job");
  }
  active = false;
  updateButton();
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  active = true;
  updateButton();
  jobStatus.textContent = "Uploading workbook…";
  try {
    const job = await api("/api/jobs", {
      method: "POST", headers: {"X-Report-Request": "1"}, body: new FormData(form),
    });
    element("results").hidden = true;
    result = undefined;
    sessionStorage.setItem("job", job.id);
    follow(job.id);
  } catch (error) {
    jobStatus.textContent = error.message;
    active = false;
    updateButton();
  }
});

refreshAvailability();
setInterval(refreshAvailability, 10000);
if (previous) follow(previous);
