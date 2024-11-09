let filesArray = [];

// Function to add selected file to the list
function addFileToList() {
    const fileInput = document.getElementById("pdfFiles");
    const files = fileInput.files;
    const fileList = document.getElementById("form-group");

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        filesArray.push(file);  // Add file to the array

        let fileContainer = document.createElement("div");
        fileContainer.className = "file-container";

        let fileName = document.createElement("span");
        fileName.textContent = file.name;
        fileContainer.appendChild(fileName);

        let deleteButton = document.createElement("button");
        deleteButton.textContent = "-";
        deleteButton.classList.add("square-button", "delete");
        deleteButton.onclick = () => {
            // Remove file from the array
            filesArray = filesArray.filter(f => f.name !== file.name);
            fileContainer.remove();
        };
        fileContainer.appendChild(deleteButton);

        fileList.appendChild(fileContainer);
    }

    // Clear the file input
    fileInput.value = "";
}

// Function to handle form submission
async function submitForm(event) {
    event.preventDefault();
    
    if (filesArray.length === 0) {
        alert("Please upload at least one PDF file.");
        return;
    }

    const textarea = document.getElementById("regNumbers");
    let regNumbers = textarea.value;
    const regNumbersArray = regNumbers.split(/[\n,]/).map(number => number.trim()).filter(number => number);

    let formdata = new FormData();
    formdata.append("regNumbers", JSON.stringify(regNumbersArray));

    // Append files from filesArray to FormData
    for (let i = 0; i < filesArray.length; i++) {
        formdata.append("pdfFiles", filesArray[i]);
    }

    try {
        let response = await fetch("/upload", {
            method: 'POST',
            body: formdata
        });

        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json();
        if (!data.success) {
            console.error('Error:', data.message);
            alert('Failed to upload files: ' + data.message);
        } else {
            console.log('Files:', data.files);
            alert('Files uploaded successfully.');

            // Proceed to next step or action
        }
    } catch (error) {
        console.error('Fetch error:', error);
        alert('Error uploading files.');
    }
}

// Function to update the sums of each column
function updateColumnSums() {
    const sumRow = document.getElementById("sumRow");
    const camTableBody = document.getElementById("camTable").querySelector("tbody");
    const numRows = camTableBody.rows.length;

    for (let i = 1; i <= 7; i++) { // Columns CO1 to Total (indices 1 to 7)
        let columnSum = 0;
        for (let j = 0; j < numRows; j++) {
            let cellValue = parseFloat(camTableBody.rows[j].cells[i].querySelector('input')?.value || camTableBody.rows[j].cells[i].textContent) || 0;
            columnSum += cellValue;
        }
        sumRow.cells[i].textContent = columnSum;
    }
}

// Function to update the total of a row
function updateRowTotal(row) {
    let total = 0;
    const cells = row.querySelectorAll('input.editable-cell');
    cells.forEach(cell => {
        let value = parseFloat(cell.value) || 0;
        total += value;
    });
    const totalCell = row.querySelector('.total-cell');
    if (totalCell) {
        totalCell.textContent = total;
    }
}

// Function to handle CAM table generation
async function generateCAMTable() {
    try {
        let response = await fetch("/generate_cam", {
            method: 'POST',
            body: JSON.stringify({ fileNames: filesArray.map(file => file.name) }),
            headers: {
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error('Network response was not ok');
        }

        const data = await response.json();
        const camTableContainer = document.getElementById("camTableContainer");
        const camTableBody = document.getElementById("camTable").querySelector("tbody");

        // Clear any existing rows
        camTableBody.innerHTML = "";

        // Populate the table with data
        data.cam_table.forEach((rowData, rowIndex) => {
            let row = document.createElement("tr");
            rowData.forEach((cellData, index) => {
                let cell = document.createElement("td");

                if (index === 0) {
                    // First column should contain the PDF name as text
                    cell.textContent = cellData;
                } else if (index < 7) {
                    // Other columns should be editable
                    let input = document.createElement("input");
                    input.type = "text"; // Use type="text" instead of type="number"
                    input.pattern = "[0-9]*[.]?[0-9]*";
                    input.value = cellData;
                    input.step = "0.01";
                    input.className = "editable-cell";
                    input.oninput = function () {
                        // Ensure only positive integers are allowed
                        this.value = this.value.replace(/[^0-9.]/g, '');
                        updateRowTotal(row);
                        updateColumnSums();
                    };
                    cell.appendChild(input);
                } else {
                    // Total column should not be editable
                    cell.textContent = cellData;
                    cell.className = "total-cell";
                }

                row.appendChild(cell);
            });

            camTableBody.appendChild(row);

            // Update total for the initial data
            updateRowTotal(row);
            
        });

        // Update column sums for the initial data
        updateColumnSums();

        // Show the table
        camTableContainer.style.display = "block";
    } catch (error) {
        console.error('Fetch error:', error);
        alert('Error generating CAM table.');
    }
}

// Function to calculate attainment
async function calculateAttainment(event) {
    event.preventDefault();

    const camTableBody = document.getElementById("camTable").querySelector("tbody");
    const numRows = camTableBody.rows.length;
    let rowDataArray = [];

    for (let i = 0; i < numRows; i++) {
        let rowData = [];
        const cells = camTableBody.rows[i].cells;

        for (let j = 1; j < cells.length - 1; j++) {
            let value = parseFloat(cells[j].querySelector('input')?.value) || 0;
            rowData.push(value);
        }

        rowDataArray.push(rowData);
    }

    const textarea = document.getElementById("regNumbers");
    let regNumbers = textarea.value;
    let regNumbersArray = regNumbers.split(/[\n,]/).map(number => number.trim()).filter(number => number);

    let formdata = new FormData();
    formdata.append("regNumbers", JSON.stringify(regNumbersArray));
    formdata.append("fileOrder", JSON.stringify(filesArray.map(file => file.name)));
    
    // Safe retrieval of targetPercentage value
    const targetPercentageInput = document.getElementById("targetPercentage");
    let box_forlevel = document.getElementById("box_forlevel");
    let targetPercentage = targetPercentageInput ? targetPercentageInput.value.trim() : 0;
    let values=[];
    
    for( let row of box_forlevel.rows){
        for(cell of row.cells){
            if(cell.querySelector('input')){
                console.log(cell.querySelector('input').value)
                values.push(parseFloat(cell.querySelector('input').value));
            }
            else{
                console.log(cell.textContent)


            }
        }
    }
    keys=[3,2,1,0];
    function createDict(keys, values) {
        let dict = {};
        let valueIndex = 0;
    
        keys.forEach((key, index) => {
            // Ensure we don't go out of bounds
            if (valueIndex < values.length) {
                if (index === keys.length - 1) {
                    // Last key gets the remaining values as a pair
                    dict[key] = `${values[valueIndex]}-${values[valueIndex + 1]}`;
                } else {
                    // For non-last keys, assign range
                    dict[key] = `${values[valueIndex]}-${values[valueIndex + 1]}`;
                    valueIndex += 2; // Move to the next pair
                }
            }
        });
    
        return dict;
    }
    
    let Dicttarget = createDict(keys, values);
    console.log(Dicttarget);
    console.log("COPOMapperTable");
   
// Get the table element by ID
let COPOMapperTable = document.getElementById("COPOMapperTable");

// Extract column headers
let headers = [];
for (let cell of COPOMapperTable.rows[0].cells) {
    headers.push(cell.textContent.trim());
}

// Initialize an empty dictionary to store column values
let dictionary = {};

// Iterate over each column index (skipping the first column)
for (let colIndex = 1; colIndex < headers.length; colIndex++) {
    let colValues = [];

    // Collect values for the current column from each row
    for (let rowIndex = 1; rowIndex < COPOMapperTable.rows.length; rowIndex++) {
        let cell = COPOMapperTable.rows[rowIndex].cells[colIndex];
        let inputElement = cell.querySelector('input');

        if (inputElement) {
            let value = inputElement.value.trim(); // Trim to handle any accidental spaces
            if (value === "") {
                colValues.push(NaN); // Push NaN for empty values
            } else {
                colValues.push(parseFloat(value));
            }
        } else {
            let cellText = cell.textContent.trim();
            if (cellText === "") {
                colValues.push(NaN); // Push NaN for empty text content
            } else {
                colValues.push(parseFloat(cellText));
            }
        }
    }

    // Assign the collected column values to the dictionary
    dictionary[headers[colIndex]] = colValues;
}

// Log the dictionary
console.log(dictionary);







    formdata.append("Target_range", JSON.stringify(Dicttarget));
    formdata.append("COPOMapperTablevalues", JSON.stringify(dictionary));
    formdata.append("targetPercentage", JSON.stringify(targetPercentage));
    console.log("Target Percentage:", JSON.stringify(targetPercentage));

    // Append files from filesArray to FormData
    for (let i = 0; i < filesArray.length; i++) {
        formdata.append("pdfFiles", filesArray[i]);
    }

    // Append ComponentsArray to FormData
    formdata.append("ComponentsArray", JSON.stringify(rowDataArray));

    try {
        let response = await fetch("/calculateAttainment", {
            method: 'POST',
            body: formdata
        });

        if (!response.ok) {
            const errorText = await response.text();
            alert(`Error calculating attainment: ${errorText}`);
            throw new Error(`Network response was not ok: ${errorText}`);
        }

        const data = await response.json();
        alert('Attainment calculation successful');
        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = data.download_url;
        downloadLink.style.display = 'block';

    } catch (error) {
        console.error('Fetch error:', error);
        alert('Error calculating attainment');
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const generateCAMButton = document.getElementById('generateCAMButton');
    const calculateAttainmentButton = document.getElementById('calculateAttainmentButton');
    const camTableContainer = document.getElementById('camTableContainer');
    calculateAttainmentButton.addEventListener('click', calculateAttainment);
    document.getElementById("addFileButton").addEventListener('click', addFileToList);
    document.getElementById("uploadForm").addEventListener('submit', submitForm);

    generateCAMButton.addEventListener('click', async () => {
        await generateCAMTable();
        camTableContainer.style.display = 'block';
        calculateAttainmentButton.style.display = 'block';
    });
});
