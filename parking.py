from flask import Flask, render_template, request
import random
from openpyxl import Workbook
import os

app = Flask(__name__)

residents = []
allocations = {}

@app.route("/", methods=["GET", "POST"])
def home():
    message = ""

    global residents, allocations

    if request.method == "POST":
        action = request.form.get("action")

        if action == "add":
            flat = request.form.get("flat")
            name = request.form.get("name")

            for resident in residents:
                if resident["flat"] == flat:
                    message = "Flat already exists!"
                    return render_template("index.html", message=message, allocations=allocations)

                if resident["name"].lower() == name.lower():
                    message = "Resident already exists!"
                    return render_template("index.html", message=message, allocations=allocations)

            residents.append({
                "flat": flat,
                "name": name
            })

            message = "Resident added successfully!"

        elif action == "allocate":
            slots = int(request.form.get("slots"))

            if slots < len(residents):
                message = "Not enough parking slots!"
            else:
                parking_slots = [f"Slot-{i}" for i in range(1, slots + 1)]
                random.shuffle(parking_slots)

                allocations.clear()

                for resident in residents:
                    allocations[resident["flat"]] = {
                        "name": resident["name"],
                        "slot": parking_slots.pop()
                    }

                wb = Workbook()
                ws = wb.active
                ws.title = "Parking Allocation"

                ws.append(["Flat Number", "Resident Name", "Parking Slot"])

                for flat, data in allocations.items():
                    ws.append([flat, data["name"], data["slot"]])

                wb.save("parking_allocation.xlsx")

                message = "Parking allocated successfully!"

    return render_template("index.html", message=message, allocations=allocations)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))