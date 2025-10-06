document.getElementById("formBusinessTravel").addEventListener("submit", function(e) {
    e.preventDefault();

    const form = e.target;
    const formData = new FormData(form);
    const data = Object.fromEntries(formData.entries());
    data.formType = "BusinessTravel";
    const scriptURL = "https://script.google.com/macros/s/AKfycbz9hXoqeMs0iAkaweBaBwQ4Cn60SVmjrwWl4cPz9S-5C1nV2JQYtqOGmHgFf11864nX/exec";
  
    fetch(scriptURL, {
      method: "POST",
      body: JSON.stringify(data),
      headers: { "Content-Type": "application/json" }
    })
  .then((res) => res.json())
  .then((res) => {
    alert("✅ Business Travel form submitted successfully!");
    form.reset();
  })
  .catch((err) => {
    console.error("❌ Error submitting EntryRequest:", err);
    alert("Error! Check console.");
  });
});

document.getElementById("formAnnualTicket").addEventListener("submit", function(e) {
  e.preventDefault();

  const form = e.target;
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());
  data.formType = "AnnualTicket";

  const scriptURL = "https://script.google.com/macros/s/AKfycbz9hXoqeMs0iAkaweBaBwQ4Cn60SVmjrwWl4cPz9S-5C1nV2JQYtqOGmHgFf11864nX/exec";

  fetch(scriptURL, {
    method: "POST",
    body: JSON.stringify(data),
    headers: {
      "Content-Type": "application/json"
    }
  })
  .then((res) => res.json())
  .then((res) => {
    alert("✅ Annual Ticket form submitted successfully!");
    form.reset();
  })
  .catch((err) => {
    console.error("❌ Error submitting Annual Ticket:", err);
    alert("Error! Check console.");
  });
});

document.getElementById("formEntryRequest").addEventListener("submit", function(e) {
  e.preventDefault();

  const form = e.target;
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());
  data.formType = "EntryRequest";
  
  const scriptURL = "https://script.google.com/macros/s/AKfycbz9hXoqeMs0iAkaweBaBwQ4Cn60SVmjrwWl4cPz9S-5C1nV2JQYtqOGmHgFf11864nX/exec";

  fetch(scriptURL, {
    method: "POST",
    body: JSON.stringify(data),
    headers: { "Content-Type": "application/json" }
  })
  .then((res) => res.json())
  .then((res) => {
    alert("✅ EntryRequest form submitted successfully!");
    form.reset();
  })
  .catch((err) => {
    console.error("❌ Error submitting EntryRequest:", err);
    alert("Error! Check console.");
  });
});

document.getElementById("formExitTicket").addEventListener("submit", function(e) {
  e.preventDefault();

  const form = e.target;
  const formData = new FormData(form)
  const data = Object.fromEntries(formData.entries());
  data.formType = "ExitTicket";

  const scriptURL = "https://script.google.com/macros/s/AKfycbz9hXoqeMs0iAkaweBaBwQ4Cn60SVmjrwWl4cPz9S-5C1nV2JQYtqOGmHgFf11864nX/exec";

  fetch(scriptURL, {
    method: "POST",
    body: JSON.stringify(data),
    headers: { "Content-Type": "application/json" }
  })
  .then((res) => res.json())
  .then((res) => {
    alert("✅ Exit Ticket form submitted successfully!");
    form.reset();
  })
  .catch((err) => {
    console.error("❌ Error submitting Exit Ticket:", err);
    alert("Error! Check console.");
  });
});
