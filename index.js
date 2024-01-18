const XLSX = require('xlsx');
const path = require('path');

// Function to convert time strings to minutes
function timeStringToMinutes(timeString) {
    if (!timeString || typeof timeString !== 'string') {
        return 0; // Return 0 if timeString is undefined or not a string
    }

    const [hours, minutes] = timeString.split(':').map(Number);
    return hours * 60 + minutes;
}

// Function to analyze the Excel file
function analyzeExcelFile() {
    try {
        const fileName = 'Assignment_Timecard.xlsx'; // Replace with the actual file name
        const filePath = path.join(__dirname, fileName);
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Access data in the worksheet
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Task a) Identify employees who have worked for 7 consecutive days
        data.forEach((row, rowIndex) => {
            if (rowIndex > 0) {
                const employeeName = row[7]; // Assuming employee name is in column 7
                const currentDate = new Date(row[5]); // Assuming the date is in column 5

                // Check for 7 consecutive days
                const consecutiveDays = data.slice(rowIndex, rowIndex + 7).every((nextRow, i) => {
                    const nextDate = new Date(nextRow[5]);
                    const expectedDate = new Date(currentDate);
                    expectedDate.setDate(expectedDate.getDate() + i);
                    return nextDate.toDateString() === expectedDate.toDateString();
                });

                if (consecutiveDays) {
                    console.log(`${employeeName} has worked for 7 consecutive days.`);
                }
            }
        });

        // Task b) Identify employees with less than 10 hours between shifts but greater than 1 hour
        data.forEach((row, rowIndex) => {
            if (rowIndex > 0) {
                const employeeName = row[7]; // Assuming employee name is in column 7
                const startTime = timeStringToMinutes(row[2]); // Assuming time is in column 2
                const previousEndTime = timeStringToMinutes(data[rowIndex - 1][3]); // Assuming time out is in column 3

                // Check for less than 10 hours between shifts but greater than 1 hour
                const timeBetweenShifts = startTime - previousEndTime;
                if (timeBetweenShifts >= 60 && timeBetweenShifts <= 600) {
                    console.log(`${employeeName} has less than 10 hours between shifts but greater than 1 hour.`);
                }
            }
        });

        // Task c) Identify employees who have worked for more than 14 hours in a single shift
        data.forEach((row, rowIndex) => {
            if (rowIndex > 0) {
                const employeeName = row[7]; // Assuming employee name is in column 7
                const startTime = timeStringToMinutes(row[2]); // Assuming time is in column 2
                const endTime = timeStringToMinutes(row[3]); // Assuming time out is in column 3

                // Check for more than 14 hours in a single shift
                const shiftDuration = endTime - startTime;
                if (shiftDuration > 840) {
                    console.log(`${employeeName} has worked for more than 14 hours in a single shift.`);
                }
            }
        });
    } catch (error) {
        console.error('Error reading or analyzing the Excel file:', error);
    }
}

// Example usage
analyzeExcelFile();
