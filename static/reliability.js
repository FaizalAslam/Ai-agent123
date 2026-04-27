(function () {
  function isSuccess(data) {
    return data && (data.success === true || data.status === "success");
  }

  function hasKnownShape(data) {
    return data && (Object.prototype.hasOwnProperty.call(data, "success") ||
      Object.prototype.hasOwnProperty.call(data, "status"));
  }

  function responseMessage(data) {
    if (!data) return "Unexpected backend response";
    const filePath = data.file_path || data.output_file || "";
    const base = data.message || data.error || "Unexpected backend response";
    return filePath ? `${base} Output: ${filePath}` : base;
  }

  window.post = async function post(url, body = {}) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    const text = await res.text();
    let data;
    try {
      data = text ? JSON.parse(text) : {};
    } catch (err) {
      throw new Error(`Malformed JSON from backend (${res.status}).`);
    }

    if (!res.ok) {
      const msg = responseMessage(data) || `Backend request failed (${res.status}).`;
      const error = new Error(msg);
      error.data = data;
      throw error;
    }

    return data;
  };

  window.sendCommand = async function sendCommand() {
    const input = document.getElementById("cmdInput");
    const cmd = input.value.trim();
    if (!cmd) return;
    const button = input.parentElement ? input.parentElement.querySelector("button") : null;

    input.disabled = true;
    if (button) button.disabled = true;
    setStatus("cmdStatus", "Processing...", "info");

    try {
      const data = await window.post("/execute", { command: cmd });
      if (!hasKnownShape(data)) throw new Error("Unexpected backend response.");
      const ok = isSuccess(data);
      setStatus("cmdStatus", responseMessage(data), ok ? "success" : "error");
      addToLog("cmdLog", cmd, ok);
      if (ok) input.value = "";
    } catch (err) {
      setStatus("cmdStatus", err.message || "Command failed.", "error");
      addToLog("cmdLog", cmd, false);
    } finally {
      input.disabled = false;
      if (button) button.disabled = false;
      input.focus();
    }
  };

  window.sendOfficeCommand = async function sendOfficeCommand() {
    const input = document.getElementById("officeInput");
    const cmd = input.value.trim();
    if (!cmd) return;
    const full = `agent: ${activeApp}: ${cmd}`;
    const button = input.parentElement ? input.parentElement.querySelector("button") : null;

    input.disabled = true;
    if (button) button.disabled = true;
    setStatus("officeStatus", "Processing command...", "info");

    try {
      const data = await window.post("/office/execute", {
        command: full,
        app: activeApp,
        raw: cmd,
      });
      if (!hasKnownShape(data)) throw new Error("Unexpected backend response.");
      const ok = isSuccess(data);
      setStatus("officeStatus", responseMessage(data), ok ? "success" : "error");
      addToLog("officeLog", full, ok);
      if (ok) input.value = "";
    } catch (err) {
      setStatus("officeStatus", err.message || "Office command failed.", "error");
      addToLog("officeLog", full, false);
    } finally {
      input.disabled = false;
      if (button) button.disabled = false;
      updatePreview();
      input.focus();
    }
  };
}());
