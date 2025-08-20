function sendAnyway() {
  Office.context.ui.messageParent("send");
}
function cancelSend() {
  Office.context.ui.messageParent("cancel");
}
window.onload = function () {
  const urlParams = new URLSearchParams(window.location.search);
  const count = urlParams.get("count");
  document.getElementById("msg").innerText =
    `This email has ${count} recipients. Do you want to send it to all?`;
};
