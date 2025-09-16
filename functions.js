function sendDataToApi(event) {
  fetch("https://your-api.com/endpoint", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ docUrl: Office.context.document.url })
  })
  .then(res => res.json())
  .then(data => console.log("API result:", data))
  .catch(err => console.error(err))
  .finally(() => event.completed());
}
