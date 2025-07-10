Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Add-in is ready in Outlook");
  }
});
