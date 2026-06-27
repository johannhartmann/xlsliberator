(function () {
  const container = document.getElementById("job");
  if (!container) return;

  const jobId = container.dataset.jobId;
  const status = document.getElementById("status");
  const events = document.getElementById("events");
  const downloads = document.getElementById("downloads");
  let next = events ? events.children.length : 0;

  function appendEvent(event) {
    const item = document.createElement("li");
    item.className = event.level;
    const phase = document.createElement("span");
    phase.textContent = event.phase;
    item.appendChild(phase);
    item.appendChild(document.createTextNode(event.message));
    events.appendChild(item);
    status.textContent = event.phase;
    if (event.phase === "completed") downloads.classList.remove("hidden");
  }

  async function poll() {
    const response = await fetch(`/api/jobs/${jobId}/events?since=${next}`);
    if (!response.ok) return;
    const payload = await response.json();
    payload.events.forEach(appendEvent);
    next = payload.next;
    if (!["completed", "failed", "cancelled"].includes(status.textContent)) {
      window.setTimeout(poll, 1000);
    }
  }

  window.setTimeout(poll, 1000);
})();
