\
let email = {
  subject: "",
  senderName: "",
  senderEmail: "",
  received: ""
};

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus("Open this add-in inside Outlook.");
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item) {
    setStatus("Select or open an email first.");
    return;
  }

  email.subject = item.subject || "(No subject)";
  email.senderName = item.from?.displayName || "";
  email.senderEmail = item.from?.emailAddress || "";
  email.received = item.dateTimeCreated
    ? new Date(item.dateTimeCreated).toLocaleString()
    : "";

  document.getElementById("subject").textContent = email.subject;
  document.getElementById("from").textContent =
    email.senderName
      ? `${email.senderName}${email.senderEmail ? ` <${email.senderEmail}>` : ""}`
      : email.senderEmail;
  document.getElementById("received").textContent = email.received;

  const person = email.senderName || email.senderEmail || "sender";
  document.getElementById("action").value = `Reply to ${person}: ${email.subject}`;
  document.getElementById("date").value = localDateString(new Date());

  document.getElementById("addButton").addEventListener("click", openObsidian);
  document.getElementById("copyButton").addEventListener("click", copyUri);
});

function localDateString(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function buildTask() {
  const action = document.getElementById("action").value.trim();
  const date = document.getElementById("date").value;

  if (!action) throw new Error("Enter a task.");
  if (!date) throw new Error("Choose a scheduled date.");

  let task = `- [ ] ${action} ⏳ ${date}`;

  if (document.getElementById("includeContext").checked) {
    const from = email.senderName || email.senderEmail || "Unknown sender";
    task += `\n  - Email context: From ${from} | Subject: ${email.subject}`;
  }

  return task;
}

function buildUri() {
  const cfg = window.ADDIN_CONFIG;
  const params = new URLSearchParams({
    vault: cfg.vaultName,
    filepath: cfg.taskFile,
    heading: cfg.taskHeading,
    mode: "append",
    data: buildTask(),
    openmode: "silent"
  });
  return `obsidian://adv-uri?${params.toString()}`;
}

function openObsidian() {
  try {
    const uri = buildUri();
    setStatus("Opening Obsidian…");
    window.location.href = uri;
    setTimeout(() => setStatus("Task sent. Check Task Inbox → Work."), 1200);
  } catch (error) {
    setStatus(error.message);
  }
}

async function copyUri() {
  try {
    const uri = buildUri();
    await navigator.clipboard.writeText(uri);
    setStatus("Obsidian link copied. Paste it into Safari or Chrome if Outlook blocks opening it.");
  } catch (error) {
    setStatus(`Could not copy automatically: ${error.message}`);
  }
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
}
