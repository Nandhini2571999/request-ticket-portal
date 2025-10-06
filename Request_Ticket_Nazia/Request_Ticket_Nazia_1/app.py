from flask import Flask, render_template, request
import requests
import logging

# Your deployed Apps Script Web App URL
SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz9hXoqeMs0iAkaweBaBwQ4Cn60SVmjrwWl4cPz9S-5C1nV2JQYtqOGmHgFf11864nX/exec"

app = Flask(__name__)

# Enable Flask logging to console
logging.basicConfig(level=logging.DEBUG)

# ✅ Home route
@app.route("/")
def home():
    app.logger.debug("Rendering Home.html")
    return render_template("Home.html")


# ✅ Business Travel form
@app.route("/business_travel", methods=["GET", "POST"])
def business_travel():
    if request.method == "POST":
        data = request.form.to_dict()
        data["formType"] = "BusinessTravel"
        app.logger.debug(f"BusinessTravel data received: {data}")
        try:
            res = requests.post(SCRIPT_URL, json=data)
            app.logger.debug(f"Google Script Response: {res.text}")
            res.raise_for_status()
            return "✅ Business Travel Request Submitted!"
        except Exception as e:
            app.logger.error(f"Error submitting BusinessTravel: {e}")
            return f"❌ Error submitting Business Travel: {e}"
    return render_template("Business_Travel.html")


# ✅ Annual Ticket form
@app.route("/annual_ticket", methods=["GET", "POST"])
def annual_ticket():
    if request.method == "POST":
        data = request.form.to_dict()
        print(data)
        data["formType"] = "AnnualTicket"
        app.logger.debug(f"AnnualTicket data received: {data}")
        try:
            res = requests.post(SCRIPT_URL, json=data)
            app.logger.debug(f"Google Script Response: {res.text}")
            res.raise_for_status()
            return "✅ Annual Ticket Request Submitted!"
        except Exception as e:
            app.logger.error(f"Error submitting AnnualTicket: {e}")
            return f"❌ Error submitting Annual Ticket: {e}"
    return render_template("Annual_Ticket.html")


# ✅ Entry Request form
@app.route("/entry_request", methods=["GET", "POST"])
def entry_request():
    if request.method == "POST":
        data = request.form.to_dict()
        data["formType"] = "EntryRequest"
        app.logger.debug(f"EntryRequest data received: {data}")
        try:
            res = requests.post(SCRIPT_URL, json=data)
            app.logger.debug(f"Google Script Response: {res.text}")
            res.raise_for_status()
            return "✅ Entry Request Submitted!"
        except Exception as e:
            app.logger.error(f"Error submitting EntryRequest: {e}")
            return f"❌ Error submitting Entry Request: {e}"
    return render_template("Entry_Request.html")


# ✅ Exit Ticket form
@app.route("/exit_ticket", methods=["GET", "POST"])
def exit_ticket():
    if request.method == "POST":
        data = request.form.to_dict()
        data["formType"] = "ExitTicket"
        app.logger.debug(f"ExitTicket data received: {data}")
        try:
            res = requests.post(SCRIPT_URL, json=data)
            app.logger.debug(f"Google Script Response: {res.text}")
            res.raise_for_status()
            return "✅ Exit Ticket Request Submitted!"
        except Exception as e:
            app.logger.error(f"Error submitting ExitTicket: {e}")
            return f"❌ Error submitting Exit Ticket: {e}"
    return render_template("Exit_Ticket.html")


if __name__ == "__main__":
    # Debug mode = shows detailed error page + auto reload
    app.run(debug=True, port=5000)
