import fs from "fs";
import XLSX from "xlsx";

// Load the Excel file
const workbook = XLSX.readFile("./Assignment_Timecard.xlsx");
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert Excel sheet to JSON
const jsonData = XLSX.utils.sheet_to_json(sheet);

// Function to convert Excel serial number to Date and format it
const convertSerialToDate = (serialNumber, isTime = false) => {
  const baseDate = isTime ? new Date(1900, 0, 0) : new Date(1900, 0, 1);
  const valueInMs = serialNumber * 24 * 60 * 60 * 1000;
  const date = new Date(baseDate.getTime() + valueInMs);

  return date;
};

// Iterate through the data and convert date and time serial numbers to Date objects
jsonData.forEach((entry) => {
  entry["Time"] = convertSerialToDate(entry["Time"]);
  entry["Time Out"] = convertSerialToDate(entry["Time Out"]);
  entry["Pay Cycle Start Date"] = convertSerialToDate(
    entry["Pay Cycle Start Date"]
  );
  entry["Pay Cycle End Date"] = convertSerialToDate(
    entry["Pay Cycle End Date"]
  );
});


let result = [];

function calculateConsecutiveDays(employeeData) {
  // Convert string dates to Date objects
  const startDate = new Date(employeeData[0]["Pay Cycle Start Date"]);
  const endDate = new Date(
    employeeData[employeeData.length - 1]["Pay Cycle End Date"]
    );
    
    // Calculate the number of consecutive days
    const daysDifference = Math.ceil(
      (endDate - startDate) / (1000 * 60 * 60 * 24)
      );
      return daysDifference;
    }
    
// Function to check if an employee has worked for 7 consecutive days
const printEmployeesForConsecutiveDays = (
  data,
  consecutiveDaysThreshold = 7
) => {
  const employeeDataByPosition = {};

  for (const entry of data) {
    const positionId = entry["Position ID"];
    if (!employeeDataByPosition[positionId]) {
      employeeDataByPosition[positionId] = [];
    }

    employeeDataByPosition[positionId].push(entry);
  }

  for (const [positionId, employeeData] of Object.entries(
    employeeDataByPosition
  )) {
    const consecutiveDays = calculateConsecutiveDays(employeeData);
    if (consecutiveDays >= consecutiveDaysThreshold) {
      const employeeName = employeeData[0]["Employee Name"];
      const positionStatus = employeeData[0]["Position Status"];
      result.push(
        `Employee Name: ${employeeName}, Position ID: ${positionId}, Position Status: ${positionStatus}  Has Worked For 7 Consecutive Days`
      );
    }
  }
  
};

function calculateTimeDifference(time1, time2) {
  const dateTime1 = new Date(time1);
  const dateTime2 = new Date(time2);
  const timeDifference = Math.abs(dateTime2 - dateTime1) / (60 * 60 * 1000); // in hours
  return timeDifference;
}


// Function to check if an employee has taken short break between shift
const printEmployeesForTimeBetweenShifts = (data, minTime, maxTime) => {
  const employeeDataByPosition = {};

  for (const entry of data) {
    const positionId = entry["Position ID"];
    if (!employeeDataByPosition[positionId]) {
      employeeDataByPosition[positionId] = [];
    }

    employeeDataByPosition[positionId].push(entry);
  }

  for (const [positionId, employeeData] of Object.entries(
    employeeDataByPosition
  )) {
    for (let i = 1; i < employeeData.length; i++) {
      const timeBetweenShifts = calculateTimeDifference(
        employeeData[i - 1]["Time Out"],
        employeeData[i]["Time"]
      );

      if (timeBetweenShifts > minTime && timeBetweenShifts < maxTime) {
        const employeeName = employeeData[0]["Employee Name"];
        const positionStatus = employeeData[0]["Position Status"];
        result.push(
          `Employee Name: ${employeeName}, Position ID: ${positionId}, Position Status: ${positionStatus}  Has Taken Short Break Between The Shifts`
        );
        break; // Break the inner loop as we found one match for this position
      }
    }
  }
  
};

function calculateShiftDuration(timeIn, timeOut) {
  const dateTimeIn = new Date(timeIn);
  const dateTimeOut = new Date(timeOut); 
  const shiftDuration = ( dateTimeOut - dateTimeIn ) / ( 60 * 60 * 1000 ); // in hours
  return shiftDuration;
}


// Function to check if an employee has worked for more than 14 hours in a single shift
const printEmployeesForLongShifts = (data, thresholdHours) => {
  for (const entry of data) {
    const shiftDuration = calculateShiftDuration(
      entry["Time"],
      entry["Time Out"]
    ).toFixed(0)

    if ( shiftDuration > thresholdHours ) {
      const employeeName = entry["Employee Name"];
      const positionId = entry["Position ID"];
      const positionStatus = entry["Position Status"];
      result.push(
        `Employee Name: ${employeeName}Position ID: ${positionId}, Position Status: ${positionStatus},Shift Duration: Has worked for ${shiftDuration} hours`
      );
    } 
  }
  
};


const [longshift, timeshift, consecutiveday] = await Promise.all([
  printEmployeesForLongShifts(jsonData, 14),
  printEmployeesForTimeBetweenShifts(jsonData, 1, 10),
  printEmployeesForConsecutiveDays(jsonData),
]);

console.log(result)

const arr = [longshift, timeshift, consecutiveday, result].flat();

fs.writeFileSync("./output.txt", Buffer.from(arr.join("\n\n")));