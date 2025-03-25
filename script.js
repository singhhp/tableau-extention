document.addEventListener("DOMContentLoaded", function () {
    logMessage("Initializing Tableau Extension...");

    // ✅ Ensure Tableau API is available
    if (typeof tableau === "undefined") {
        logMessage("❌ Tableau API is NOT available. Make sure this is running inside Tableau.");
        return;
    }

    // ✅ Initialize Tableau Extensions API
    tableau.extensions.initializeAsync().then(() => {
        logMessage("✅ Tableau Extensions API loaded successfully!");
    }).catch(err => {
        logMessage("❌ Error initializing Tableau API: " + err.message);
        console.error(err);
    });

    // ✅ Add event listener for export button
    document.getElementById("exportBtn").addEventListener("click", exportToExcel);
});

// ✅ Function to export Tableau data to Excel
async function exportToExcel() {
    try {
        logMessage("🔄 Fetching data from Tableau...");

        // 🔹 Get active dashboard
        let dashboard = tableau.extensions.dashboardContent.dashboard;

        // 🔹 Get first worksheet (Modify this if needed)
        let worksheet = dashboard.worksheets[0]; // Change index if needed

        // 🔹 Get summary data (without duplicates)
        let dataTable = await worksheet.getSummaryDataAsync();
        logMessage("✅ Data fetched successfully!");

        // 🔹 Convert Tableau data into an array
        let data = [];
        let columns = dataTable.columns.map(col => col.fieldName);
        data.push(columns); // Add headers

        // 🔹 Store unique rows in a Set to remove duplicates
        let uniqueRows = new Set();

        dataTable.data.forEach(row => {
            let rowData = row.map(cell => cell.formattedValue).join("||"); // Convert row to string
            uniqueRows.add(rowData);
        });

        // 🔹 Convert Set back to an array
        uniqueRows.forEach(rowString => {
            data.push(rowString.split("||"));
        });

        // 🔹 Create XLSX workbook
        let wb = XLSX.utils.book_new();
        let ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Tableau Data");

        // 🔹 Save as Excel file
        XLSX.writeFile(wb, "tableau_data.xlsx");
        logMessage("✅ Exported to Excel successfully, without duplicates!");

    } catch (err) {
        logMessage("❌ Error exporting data: " + err.message);
        console.error(err);
    }
}



// ✅ Function to log messages to debug log
function logMessage(message) {
    let logDiv = document.getElementById("debugLog");
    let p = document.createElement("p");
    p.textContent = message;
    logDiv.appendChild(p);
}
