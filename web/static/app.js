const jobId = document.body.dataset.jobId;

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
