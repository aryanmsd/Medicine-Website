let workbook;  // Declare the workbook globally so it persists across calls
const excelFilePath = "userdata.xlsx";  // Path to the Excel file

// Event listener for the medication form submission
document.getElementById("medicationForm").addEventListener("submit", function (e) {
    e.preventDefault();
    
    let medicationName = document.getElementById("medicationName").value;
    let dosage = document.getElementById("dosage").value;
    let time = document.getElementById("time").value;

    addReminder(medicationName, dosage, time);
    updateExcelFile(medicationName, dosage, time);  // Update the existing Excel file
    
    document.getElementById("medicationForm").reset();  // Reset the form after submission
});

document.getElementById("contactForm").addEventListener("submit", function(event) {
    event.preventDefault();

    let name = document.getElementById("name").value;
    let email = document.getElementById("email").value;
    let message = document.getElementById("message").value;

    if (name && email && message) {
        alert('Message sent successfully!');
        document.getElementById("contactForm").reset();
    } else {
        alert('Please fill out all fields.');
    }
});

// Function to add a reminder to the list and handle the delete option
function addReminder(medicationName, dosage, time) {
    let reminderList = document.getElementById("reminderList");
    let listItem = document.createElement("li");

    // Create the delete button
    let deleteButton = document.createElement("button");
    deleteButton.innerHTML = "Delete";
    deleteButton.style.marginLeft = "10px";
    
    // Event handler to remove the list item and delete from Excel
    deleteButton.onclick = function () {
        reminderList.removeChild(listItem);  // Remove the item from the list
        deleteFromExcel(medicationName);     // Remove the medication from the Excel file
    };

    // Add medication details and the delete button
    listItem.innerHTML = `${medicationName} - ${dosage} at ${time}`;
    listItem.appendChild(deleteButton);
    reminderList.appendChild(listItem);

    // Set the notification for the medication
    setNotification(time, medicationName);
}

// Function to set a notification for the medication at the specified time
function setNotification(time, medicationName) {
    let now = new Date();
    let reminderTime = new Date();
    let [hours, minutes] = time.split(":");
    reminderTime.setHours(hours, minutes, 0);

    let timeDifference = reminderTime - now;

    if (timeDifference > 0) {
        setTimeout(function () {
            alert(`Time to take your medication: ${medicationName}`);
        }, timeDifference);
    }
}

// Function to load the existing Excel workbook or create a new one
function loadWorkbook() {
    try {
        workbook = XLSX.readFile(excelFilePath);  // Try to load the existing workbook
    } catch (error) {
        // If the workbook doesn't exist, create a new one
        workbook = XLSX.utils.book_new();
        let ws_data = [["Medication", "Dosage", "Time"]];  // Initial headings
        let worksheet = XLSX.utils.aoa_to_sheet(ws_data);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Medications");
    }
}

// Function to update the existing Excel file with new medication
function updateExcelFile(medicationName, dosage, time) {
    if (!workbook) {
        loadWorkbook();  // Load the workbook if not already loaded
    }

    // Get the worksheet and existing data
    let worksheet = workbook.Sheets["Medications"];
    let ws_data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Append the new row with medication data
    ws_data.push([medicationName, dosage, time]);

    // Update the worksheet with the new data
    let updatedWorksheet = XLSX.utils.aoa_to_sheet(ws_data);
    workbook.Sheets["Medications"] = updatedWorksheet;

    // Overwrite the existing file with the updated data
    XLSX.writeFile(workbook, excelFilePath);  // This will write to the file but not trigger a download
}

// Function to delete a medication from the Excel workbook
function deleteFromExcel(medicationName) {
    if (!workbook) {
        loadWorkbook();  // Load the workbook if not already loaded
    }

    let worksheet = workbook.Sheets["Medications"];
    let ws_data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Find and remove the row that matches the medication name
    ws_data = ws_data.filter(row => row[0] !== medicationName);

    // Update the worksheet with the remaining data
    let updatedWorksheet = XLSX.utils.aoa_to_sheet(ws_data);
    workbook.Sheets["Medications"] = updatedWorksheet;

    // Overwrite the existing file with the updated data
    XLSX.writeFile(workbook, excelFilePath);
}
