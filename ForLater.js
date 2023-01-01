const electron = require('electron');
const fs = require('fs');

// We'll use the `xlsx` library to parse and create Excel files
const xlsx = require('xlsx');

// Get the path to the user's Documents folder
const documentsPath = electron.app.getPath('documents');

// Set the path to the Excel file
const filePath = `${documentsPath}/people.xlsx`;

// Read the Excel file
const workbook = xlsx.readFile(filePath);

// Get the first sheet in the workbook
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Add a new person to the sheet
const newPerson = {
  id: '12345',
  name: 'John Doe',
  age: 30
};

// Get the next available row in the sheet
const nextRow = sheet['!ref'].split(':')[1].match(/\d+/)[0];

// Set the values for the cells in the new row
sheet[`A${nextRow}`] = { t: 's', v: newPerson.id };
sheet[`B${nextRow}`] = { t: 's', v: newPerson.name };
sheet[`C${nextRow}`] = { t: 'n', v: newPerson.age };

// Update the sheet data
workbook.Sheets[workbook.SheetNames[0]] = sheet;

// Write the updated workbook to the file
xlsx.writeFile(workbook, filePath);

// Remove a person from the sheet
const personToRemove = {
  id: '12345',
};

// Loop through the rows in the sheet
for (let i = 1; i <= nextRow; i++) {
  // Get the ID of the current row
  const currentId = sheet[`A${i}`].v;

  // If the ID of the current row matches the ID of the person we want to remove
  if (currentId === personToRemove.id) {
    // Delete the cells in the current row
    for (let j = 0; j < 3; j++) {
      delete sheet[`${xlsx.utils.encode_col(j)}${i}`];
    }
  }
}

// Update the sheet data
workbook.Sheets[workbook.SheetNames[0]] = sheet;

// Write the updated workbook to the file
xlsx.writeFile(workbook, filePath);