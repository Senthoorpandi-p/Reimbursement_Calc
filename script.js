const totalAmount = 10000;

const now = new Date();
const year = now.getFullYear();
const month = now.getMonth();

const daysInMonth = new Date(year, month + 1, 0).getDate();

let saturdays = 0;
let sundays = 0;

let exportData = [];

/* Count Saturdays & Sundays */

for (let i = 1; i <= daysInMonth; i++) {

    const day = new Date(year, month, i).getDay();

    if (day === 6) {
        saturdays++;
    }

    if (day === 0) {
        sundays++;
    }

}

/* Calculate reimbursement per day */

const perDay = totalAmount / daysInMonth;

document.getElementById("perDayAmount").innerText = perDay.toFixed(2);


/* File Upload */

document.getElementById("fileInput").addEventListener("change", function (e) {

    const file = e.target.files[0];

    const reader = new FileReader();

    reader.onload = function (event) {

        exportData = [];

        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data, { type: "array" });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const rows = XLSX.utils.sheet_to_json(sheet);

        const internCount = {};

        /* Count Present Days */

        rows.forEach(function (row) {

            const id = row.InternID;
            const name = row.Name;

            if (!internCount[id]) {

                internCount[id] = {
                    name: name,
                    days: 0
                };

            }

            internCount[id].days++;

        });


        const table = document.getElementById("tableBody");

        table.innerHTML = "";

        /* Create Table */

        for (let id in internCount) {

            const presentDays = internCount[id].days;

            const payableDays = presentDays + saturdays + sundays;

            const amount = (perDay * payableDays).toFixed(2);

            const tr = document.createElement("tr");

            tr.innerHTML =
                "<td>" + id + "</td>" +
                "<td>" + internCount[id].name + "</td>" +
                "<td>" + presentDays + "</td>" +
                "<td>" + amount + "</td>";

            table.appendChild(tr);

            /* Store Export Data */

            exportData.push({
                InternID: id,
                Name: internCount[id].name,
                PresentDays: presentDays,
                Saturdays: saturdays,
                Sundays: sundays,
                TotalAmount: amount
            });

        }

        /* Show Summary */

        document.getElementById("satCount").innerText = saturdays;
        document.getElementById("sunCount").innerText = sundays;

        document.getElementById("summaryBox").style.display = "block";

        /* Show Export Button */

        document.getElementById("exportBtn").style.display = "inline-block";

    };

    reader.readAsArrayBuffer(file);

});


/* Export Excel */

document.getElementById("exportBtn").addEventListener("click", function () {

    const worksheet = XLSX.utils.json_to_sheet(exportData);

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Reimbursement");

    /* Get Month Name */

    const now = new Date();

    const monthName = now.toLocaleString('default', { month: 'long' });

    const year = now.getFullYear();

    const fileName = "Reimbursement_" + monthName + "_" + year + ".xlsx";

    XLSX.writeFile(workbook, fileName);

});
