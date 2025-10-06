const scriptURL = "https://script.google.com/macros/s/YOUR_DEPLOYED_SCRIPT_ID/exec";

function handleFormSubmit(formId, formType) {
  const form = document.getElementById(formId);
  if (!form) return; // Skip if form not present in DOM

  form.addEventListener("submit", function(e) {
    e.preventDefault();

    const formData = new FormData(form);
    const data = Object.fromEntries(formData.entries());
    data.formType = formType; // Extra field to identify form type

    fetch(scriptURL, {
      method: "POST",
      body: JSON.stringify(data),
      headers: { "Content-Type": "application/json" }
    })
    .then((res) => res.json())
    .then((res) => {
      alert(`✅ ${formType} form submitted successfully!`);
      form.reset();
    })
    .catch((err) => {
      console.error(`❌ Error submitting ${formType}:`, err);
      alert("Error! Check console.");
    });
  });
}

// Attach only if form exists
handleFormSubmit("formExitTicket", "ExitTicket");
handleFormSubmit("formBusinessTravel", "BusinessTravel");
handleFormSubmit("formAnnualTicket", "AnnualTicket");
handleFormSubmit("formEntryRequest", "EntryRequest");
