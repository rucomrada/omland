// Replace this with the name of your sheet that contains rider records
const EMPLOYEE_SHEET_NAME = 'employees';
// Replace this with the sheet name where you want to store form submissions
const FORM_RESPONSE_SHEET = 'Evaluations';
function updateAnalysisSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const evaluationSheet = sheet.getSheetByName('Evaluations');
  let analysisSheet = sheet.getSheetByName('Analysis');
  if (!analysisSheet) {
    analysisSheet = sheet.insertSheet('Analysis');
  } else {
    analysisSheet.clearContents();
    analysisSheet.clearFormats();
    analysisSheet.clearConditionalFormatRules();
    analysisSheet.clearNotes();
  }
  const data = evaluationSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row
  // Indexes for important columns
  const nameIndex = headers.indexOf('Name');
  const payrollIndex = headers.indexOf('Payroll Number');
  const branchIndex = headers.indexOf('Branch');
  const achievedTripsIndex = headers.indexOf('Achieved Trips');
  const fuelAvgIndex = headers.indexOf('Fuel Average');
  const tripPercentIndex = headers.indexOf('%TRIPS ARCHIEVED');
  const tripPointsIndex = headers.indexOf('ARCHIEVED POINTS ON TRIPS');
  const fuelPointsIndex = headers.indexOf('FUEL POINTS ARCHIEVED');
  const averageScoreIndex = headers.indexOf('AVERAGE SCORE');
  // Set headers
  const outputHeaders = [
    'Name', 'Payroll Number', 'Branch', 'Achieved Trips', 'Fuel Average',
    '% Trips Achieved', 'Trip Points', 'Fuel Points', 'Average Score', '% Score Achieved'
  ];
  analysisSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  const analysisData = {};
  data.forEach(row => {
    const name = row[nameIndex];
    if (!name) return;
    if (!analysisData[name]) {
      analysisData[name] = {
        count: 0,
        totalTrips: 0,
        totalFuel: 0,
        totalTripPercent: 0,
        totalTripPoints: 0,
        totalFuelPoints: 0,
        totalAvgScore: 0,
        payroll: row[payrollIndex],
        branch: row[branchIndex],
        validFuelCount: 0
      };
    }
    const fuel = Number(row[fuelAvgIndex]);
    if (!isNaN(fuel) && fuel > 0) {
      analysisData[name].validFuelCount++;
      analysisData[name].totalFuel += fuel;
    }
    analysisData[name].count++;
    analysisData[name].totalTrips += Number(row[achievedTripsIndex]) || 0;
    analysisData[name].totalTripPercent += parseFloat(row[tripPercentIndex]) || 0;
    analysisData[name].totalTripPoints += Number(row[tripPointsIndex]) || 0;
    analysisData[name].totalFuelPoints += Number(row[fuelPointsIndex]) || 0;
    analysisData[name].totalAvgScore += Number(row[averageScoreIndex]) || 0;
  });
  const outputRows = [];
  for (let name in analysisData) {
    const d = analysisData[name];
    const count = d.count;
    const avgFuel = d.validFuelCount ? d.totalFuel / d.validFuelCount : 0;
    const avgTripPercent = (d.totalTripPercent / count);
    const avgScore = d.totalAvgScore / count;
    const percentScore = ((avgScore / 65) * 100);
    const row = [
      name,
      d.payroll,
      d.branch,
      d.totalTrips,
      avgFuel.toFixed(2),
      avgTripPercent.toFixed(2) + '%',
      d.totalTripPoints,
      d.totalFuelPoints,
      avgScore.toFixed(2),
      percentScore.toFixed(2) + '%'
    ];
    outputRows.push({ row, score: percentScore });
  }
  // Sort descending by score
  outputRows.sort((a, b) => b.score - a.score);
  // Insert sorted data and apply formatting
  let rowIndex = 2;
  for (const item of outputRows) {
    analysisSheet.getRange(rowIndex, 1, 1, item.row.length).setValues([item.row]);
    const range = analysisSheet.getRange(rowIndex, 1, 1, item.row.length);
    const score = item.score;
    if (score >= 85) {
      range.setBackground('#C6EFCE'); // Green
      range.setFontWeight('bold');
    } else if (score >= 70) {
      range.setBackground('#FFFF99'); // Yellow
    } else {
      range.setBackground('#FFCCCC'); // Red
    }
    rowIndex++;
  }
  analysisSheet.autoResizeColumns(1, 10);
}
// Function to load employee data into the frontend dropdown
function getEmployeeData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EMPLOYEE_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const nameIndex = headers.indexOf('Name');
  const payrollIndex = headers.indexOf('PayrollNumber');
  const branchIndex = headers.indexOf('Branch');
  const bikeIndex = headers.indexOf('Bike');
  const designationIndex = headers.indexOf('Designation');
  const routeIndex = headers.indexOf('Route');
  let employeeMap = {};
  data.forEach(row => {
    const name = row[nameIndex];
    employeeMap[name] = {
      payrollNumber: row[payrollIndex],
      branch: row[branchIndex],
      bike: row[bikeIndex],
      designation: row[designationIndex],
      route: row[routeIndex]
    };
  });
  return employeeMap;
}
function submitForm(formData) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(FORM_RESPONSE_SHEET);
  if (!sheet) throw new Error("Sheet 'Evaluations' not found!");
  const headers = [
    'Timestamp', 'Name', 'Payroll Number', 'Branch', 'Bike Registration', 'Designation', 'Route',
    'Trip 1', 'Trip 2', 'Trip 3', 'Trip 4', 'Trip 5', 'Trip 6',
    'Achieved Trips', 'Fuel Average', 'No. of CODs',
    'Fraud', 'Breakages', 'Accidents', 'Customer Service',
    'Riding Gear', 'Returns', 'Section'
  ];
  // ✅ Duplicate check logic
  const data = sheet.getDataRange().getValues();
  const nameIndex = headers.indexOf("Name");
  const tsIndex = headers.indexOf("Timestamp");
  const today = new Date().toDateString();
  for (let i = 1; i < data.length; i++) {
    const rowName = data[i][nameIndex];
    const rowTimestamp = new Date(data[i][tsIndex]);
    if (rowName === formData.name && rowTimestamp.toDateString() === today) {
      throw new Error(`⚠️ Employee "${formData.name}" has already been submitted today.`);
    }
  }
  // ✅ Prepare row values
  const values = {
    'Timestamp': new Date(),
    'Name': formData.name,
    'Payroll Number': formData.payrollNumber,
    'Branch': formData.branch,
    'Bike Registration': formData.bike,
    'Designation': formData.designation,
    'Route': formData.route,
    'Trip 1': formData.trips[0] || "",
    'Trip 2': formData.trips[1] || "",
    'Trip 3': formData.trips[2] || "",
    'Trip 4': formData.trips[3] || "",
    'Trip 5': formData.trips[4] || "",
    'Trip 6': formData.trips[5] || "",
    'Achieved Trips': formData.achievedTrips,
    'Fuel Average': formData.fuelAvg,
    'No. of CODs': formData.cods,
    'Fraud': formData.frauds,
    'Breakages': formData.breakages,
    'Accidents': formData.accidents,
    'Customer Service': formData.customerService,
    'Riding Gear': formData.ridingGear,
    'Returns': formData.returns,
    'Section': formData.section
  };
  const row = headers.map(h => values[h] ?? "");
  sheet.appendRow(row);
}
    
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'form'; // Safe fallback
  const file = page === 'dashboard' ? 'dashboard' : 'form';
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle("Rider Evaluation System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getAnalysisData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archive_Analysis");
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove headers
  return data.map(row => ({
    photo: "https://www.w3schools.com/howto/img_avatar.png",
    name: row[0],
    payroll: row[1],
    branch: row[2],
    trips: row[3],
    fuelAvg: row[4],
    avgTripPercent: parseFloat(row[5]) *100,
    tripPoints: row[6],
    fuelPoints: row[7],
    avgScore: row[8],
percentScore: parseFloat(row[9]) * 100
  }));
}
function archiveOldEvaluations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName('Evaluations');
  const archiveSheet = ss.getSheetByName('Evaluations_Archive');
  const today = new Date().toDateString();
  const data = evalSheet.getDataRange().getValues();
  const headers = data[0];
  const remaining = [headers];
  const toArchive = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const timestamp = new Date(row[0]);
    if (timestamp.toDateString() === today) {
      remaining.push(row);
    } else {
      toArchive.push(row);
    }
  }
  // Replace 'Evaluations' with only today's data
  evalSheet.clearContents();
  evalSheet.getRange(1, 1, remaining.length, headers.length).setValues(remaining);
  // Safely append old data to archive
  if (toArchive.length > 0) {
    const lastRow = archiveSheet.getLastRow();
    const startRow = lastRow === 0 ? 1 : lastRow + 1;
    // If archive sheet is empty, add headers first
    if (lastRow === 0) {
      archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    archiveSheet.getRange(startRow, 1, toArchive.length, headers.length).setValues(toArchive);
  }
}
function generateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName('Evaluations');
  const dashboard = ss.getSheetByName('Dashboard');
  const data = evalSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];
  const rows = data.slice(1);
  const achievedTripsIndex = headers.indexOf("Achieved Trips");
  const branchIndex = headers.indexOf("Branch");
  let totalTrips = 0;
  let totalRiders = 0;
  const branchCounts = {};
  for (let i = 0; i < rows.length; i++) {
    const trips = Number(rows[i][achievedTripsIndex]);
    const branch = rows[i][branchIndex];
    if (!isNaN(trips)) {
      totalTrips += trips;
      totalRiders++;
    }
    if (branch) {
      branchCounts[branch] = (branchCounts[branch] || 0) + 1;
    }
  }
  // Clear Dashboard
  dashboard.clearContents();
  // Write Summary
  dashboard.getRange("A1").setValue("📊 Rider Dashboard - Today");
  dashboard.getRange("A3").setValue("Total Riders Evaluated:");
  dashboard.getRange("B3").setValue(totalRiders);
  dashboard.getRange("A4").setValue("Total Trips Achieved:");
  dashboard.getRange("B4").setValue(totalTrips);
  dashboard.getRange("A6").setValue("Branch Distribution:");
  let row = 7;
  for (let branch in branchCounts) {
    dashboard.getRange(`A${row}`).setValue(branch);
    dashboard.getRange(`B${row}`).setValue(branchCounts[branch]);
    row++;
  }
}
function updateArchiveAnalysisSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = sheet.getSheetByName('Evaluations_Archive');
  let analysisSheet = sheet.getSheetByName('Archive_Analysis');
  if (!analysisSheet) {
    analysisSheet = sheet.insertSheet('Archive_Analysis');
  } else {
    analysisSheet.clearContents();
    analysisSheet.clearFormats();
    analysisSheet.clearConditionalFormatRules();
    analysisSheet.clearNotes();
  }
  const data = archiveSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row
  const nameIndex = headers.indexOf('Name');
  const payrollIndex = headers.indexOf('Payroll Number');
  const branchIndex = headers.indexOf('Branch');
  const achievedTripsIndex = headers.indexOf('Achieved Trips');
  const fuelAvgIndex = headers.indexOf('Fuel Average');
  const tripPercentIndex = headers.indexOf('%TRIPS ARCHIEVED');
  const tripPointsIndex = headers.indexOf('ARCHIEVED POINTS ON TRIPS');
  const fuelPointsIndex = headers.indexOf('FUEL POINTS ARCHIEVED');
  const averageScoreIndex = headers.indexOf('AVERAGE SCORE');
  const outputHeaders = [
    'Name', 'Payroll Number', 'Branch', 'Achieved Trips', 'Fuel Average',
    '% Trips Achieved', 'Trip Points', 'Fuel Points', 'Average Score', '% Score Achieved'
  ];
  analysisSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  const analysisData = {};
  data.forEach(row => {
    const name = row[nameIndex];
    if (!name) return;
    if (!analysisData[name]) {
      analysisData[name] = {
        count: 0,
        totalTrips: 0,
        totalFuel: 0,
        totalTripPercent: 0,
        totalTripPoints: 0,
        totalFuelPoints: 0,
        totalAvgScore: 0,
        payroll: row[payrollIndex],
        branch: row[branchIndex],
        validFuelCount: 0
      };
    }
    const fuel = Number(row[fuelAvgIndex]);
    if (!isNaN(fuel) && fuel > 0) {
      analysisData[name].validFuelCount++;
      analysisData[name].totalFuel += fuel;
    }
    analysisData[name].count++;
    analysisData[name].totalTrips += Number(row[achievedTripsIndex]) || 0;
    analysisData[name].totalTripPercent += parseFloat(row[tripPercentIndex]) || 0;
    analysisData[name].totalTripPoints += Number(row[tripPointsIndex]) || 0;
    analysisData[name].totalFuelPoints += Number(row[fuelPointsIndex]) || 0;
    analysisData[name].totalAvgScore += Number(row[averageScoreIndex]) || 0;
  });
  const outputRows = [];
  for (let name in analysisData) {
    const d = analysisData[name];
    const count = d.count;
    const avgFuel = d.validFuelCount ? d.totalFuel / d.validFuelCount : 0;
    const avgTripPercent = (d.totalTripPercent / count);
    const avgScore = d.totalAvgScore / count;
    const percentScore = ((avgScore / 65) * 100);
    const row = [
      name,
      d.payroll,
      d.branch,
      d.totalTrips,
      avgFuel.toFixed(2),
      avgTripPercent.toFixed(2) + '%',
      d.totalTripPoints,
      d.totalFuelPoints,
      avgScore.toFixed(2),
      percentScore.toFixed(2) + '%'
    ];
    outputRows.push({ row, score: percentScore });
  }
  // Sort descending by score
  outputRows.sort((a, b) => b.score - a.score);
  let rowIndex = 2;
  for (const item of outputRows) {
    analysisSheet.getRange(rowIndex, 1, 1, item.row.length).setValues([item.row]);
    const range = analysisSheet.getRange(rowIndex, 1, 1, item.row.length);
    const score = item.score;
    if (score >= 85) {
      range.setBackground('#C6EFCE'); // Green
      range.setFontWeight('bold');
    } else if (score >= 70) {
      range.setBackground('#FFFF99'); // Yellow
    } else {
      range.setBackground('#FFCCCC'); // Red
    }
    rowIndex++;
  }
  analysisSheet.autoResizeColumns(1, 10);
}
function getBranchPerformance() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Evaluations_Archive');
  if (!sheet) return ["Waiting for data..."];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ["No data available"];
  
  const headers = data[0];
  const branchIndex = headers.indexOf('Branch');
  // Try to find 'AVERAGE SCORE' first (from analysis script), fall back to 'Customer Service'
  let scoreIndex = headers.indexOf('AVERAGE SCORE');
  if (scoreIndex === -1) scoreIndex = headers.indexOf('Customer Service');
  
  if (branchIndex === -1 || scoreIndex === -1) return ["Metrics not configured"];
  
  const branchStats = {};
  
  for (let i = 1; i < data.length; i++) {
    const branch = data[i][branchIndex];
    const score = Number(data[i][scoreIndex]);
    
    if (branch && !isNaN(score)) {
      if (!branchStats[branch]) {
        branchStats[branch] = { total: 0, count: 0 };
      }
      branchStats[branch].total += score;
      branchStats[branch].count++;
    }
  }
  
  const results = [];
  for (const branch in branchStats) {
    const avg = branchStats[branch].total / branchStats[branch].count;
    // Format to 1 decimal place. Assuming score is 1-10 or 1-100, we'll just show the number.
    // If it's a percentage calc, we might want to append %.
    results.push({ branch, avg });
  }
  
  // Sort by highest average
  results.sort((a, b) => b.avg - a.avg);
  
  // Format output strings with emojis for top 3
  return results.map((item, index) => {
    let icon = '🏢';
    if (index === 0) icon = '🥇';
    if (index === 1) icon = '🥈';
    if (index === 2) icon = '🥉';
    return `${icon} ${item.branch}: ${item.avg.toFixed(1)}`;
  });
}
