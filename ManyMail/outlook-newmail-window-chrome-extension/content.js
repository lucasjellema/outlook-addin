console.log("Outlook New Mail Extension loaded");

const inspectIfNewMail = () => {
  // this is a bit brittle: relying on the aria-label value
  const bccDiv = document.querySelector('[aria-label="Bcc"]');

  // alternatively - find the span with text 'autosend@contoso.com' - if it exists ; then find the <i> element that is its next sibling and click it
  // const spans = document.querySelectorAll('span');
  // const targetSpan = Array.from(spans).find(span => span.textContent.trim() === 'autosend@contoso.com');


  if (bccDiv && bccDiv.textContent.includes("autosend@contoso.com")) {
    try {
      const removeIconContainer = bccDiv.children[0].children[0].children[0].children[1];
       if (removeIconContainer) {
        removeIconContainer.click(); // click on the remove icon for the bcc entry
      }
    } catch (error) {
      console.log("error while resetting bcc field ", error);
    }
    const sendButton = document.querySelector('[aria-label="Send"]');
    if (sendButton) {
      setTimeout(() => { sendButton.click(); }, 1500); // allow some time for the new mail to be fully constructed from the Addin
    }
  }
}

setTimeout(() => {
  inspectIfNewMail();
}, 1000); // it takes some time for the document to be ready to be inspected (the bcc field to be set), hence this timeout
