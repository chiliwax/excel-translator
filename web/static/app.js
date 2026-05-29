const jobId = document.body.dataset.jobId;
const inspectForm = document.getElementById("inspect-form");

if (inspectForm) {
  inspectForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const message = document.getElementById("inspect-message");
    const button = document.getElementById("inspect-button");
    const jobForm = document.getElementById("job-form");
    const data = new FormData(inspectForm);

    message.textContent = "Inspecting workbook...";
    button.disabled = true;
    jobForm.hidden = true;

    try {
      const response = await fetch("/inspect", { method: "POST", body: data });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.detail || "Could not inspect workbook.");

      document.getElementById("upload-id").value = payload.upload_id;
      document.getElementById("selected-filename").textContent = payload.filename;
      renderSheetList(payload.sheets || []);
      message.textContent = `Found ${payload.sheets.length} sheet${payload.sheets.length === 1 ? "" : "s"}. Select exclusions below.`;
      jobForm.hidden = false;
      jobForm.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      message.textContent = error.message;
    } finally {
      button.disabled = false;
    }
  });
}

function renderSheetList(sheets) {
  const list = document.getElementById("sheet-list");
  list.textContent = "";

  if (!sheets.length) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = "No sheets found.";
    list.append(empty);
    return;
  }

  for (const sheet of sheets) {
    const label = document.createElement("label");
    label.className = "sheet-option";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.name = "exclude_sheet_names";
    checkbox.value = sheet;

    const text = document.createElement("span");
    text.textContent = sheet;

    label.append(checkbox, text);
    list.append(label);
  }
}

function updateJob(job) {
  document.title = `${job.status} - Translation Job`;
  document.getElementById("job-title").textContent = job.status.charAt(0).toUpperCase() + job.status.slice(1);
  document.getElementById("job-message").textContent = job.message || job.status;
  document.getElementById("progress-bar").style.width = `${job.progress_percent || 0}%`;
  document.getElementById("progress-text").textContent = `${job.progress_current}/${job.progress_total} batches`;

  const error = document.getElementById("error");
  if (job.error) {
    error.hidden = false;
    error.textContent = job.error;
  }

  const stats = document.getElementById("stats");
  if (job.stats) {
    stats.hidden = false;
    stats.textContent = JSON.stringify(job.stats, null, 2);
  }

  const download = document.getElementById("download");
  if (job.download_url) {
    download.hidden = false;
    download.href = job.download_url;
  } else {
    download.hidden = true;
  }

  return job.status === "done" || job.status === "error";
}

async function poll() {
  const response = await fetch(`/jobs/${jobId}/status`, { cache: "no-store" });
  if (!response.ok) return false;
  return updateJob(await response.json());
}

if (jobId) {
  const timer = setInterval(async () => {
    try {
      if (await poll()) clearInterval(timer);
    } catch (_) {
      // Keep polling; transient network errors should not fail the page.
    }
  }, 1000);
  poll().then((done) => {
    if (done) clearInterval(timer);
  });
}
