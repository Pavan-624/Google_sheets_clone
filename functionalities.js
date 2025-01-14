
function parseRange(range) {
    const [start, end] = range.split(':'); // Separate start and end of range
    const startRow = Number(start.match(/\d+/)[0]); // Extract row number
    const startCol = start.match(/[A-Z]+/)[0]; // Extract column letter
    const endRow = Number(end.match(/\d+/)[0]);
    const endCol = end.match(/[A-Z]+/)[0];

    const cells = [];
    for (let row = startRow; row <= endRow; row++) { // Loop through rows
        for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) { // Loop through columns
            cells.push(String.fromCharCode(col) + row); // Convert column back to letter
        }
    }
    return cells;
}

    boldBtn.addEventListener("click", function () {
        if (lsc) { // Check if a cell is selected
            let cellObj = getCellobj(lsc);
            cellObj.cellFormatting.bold = !cellObj.cellFormatting.bold; // Toggle bold
            lsc.style.fontWeight = cellObj.cellFormatting.bold ? "bold" : "normal"; // Apply style
        }
    });

    italicBtn.addEventListener("click", function () {
        if (lsc) {
            let cellObj = getCellobj(lsc);
            cellObj.cellFormatting.italic = !cellObj.cellFormatting.italic; // Toggle italic
            lsc.style.fontStyle = cellObj.cellFormatting.italic ? "italic" : "normal"; // Apply style
        }
    });

    underlineBtn.addEventListener("click", function () {
        if (lsc) {
            let cellObj = getCellobj(lsc);
            cellObj.cellFormatting.underline = !cellObj.cellFormatting.underline; // Toggle underline
            lsc.style.textDecoration = cellObj.cellFormatting.underline ? "underline" : "none"; // Apply style
        }
    });
    function setFont(target){
        if(activeCell){
            let fontInput = target.value;
            console.log(fontInput);
            activeSheetObject[activeCell.id].fontFamily_data = fontInput;
            activeCell.style.fontFamily = fontInput;
            activeCell.focus();
        }
    }
    function setSize(target){
        if(activeCell){
            let sizeInput = target.value;
            activeSheetObject[activeCell.id].fontSize_data = sizeInput;
            activeCell.style.fontSize = sizeInput+'px';
            activeCell.focus();
        }
    }

    // bug fix
    boldBtn.addEventListener('click', (event) => {
        event.stopPropagation();
        toggleBold();
    })
    function toggleBold(){
        if(activeCell){
            if(!activeSheetObject[activeCell.id].isBold) {
                activeCell.style.fontWeight = '600';
            }
            else{
                activeCell.style.fontWeight = '400';
            }
            activeSheetObject[activeCell.id].isBold = !activeSheetObject[activeCell.id].isBold;
            activeCell.focus();
        }
    }

    // bug fix
    italicBtn.addEventListener('click', (event) => {
        event.stopPropagation();
        toggleItalic();
    })
    function toggleItalic(){
        if(activeCell){
            if(!activeSheetObject[activeCell.id].isItalic) {
                activeCell.style.fontStyle = 'italic';
            }
            else{
                activeCell.style.fontStyle = 'normal';
            }
            activeSheetObject[activeCell.id].isItalic = !activeSheetObject[activeCell.id].isItalic;
            activeCell.focus();
        }
    }

    // bug fix
    underlineBtn.addEventListener('click', (event) => {
        event.stopPropagation();
        toggleUnderline();
    })
    function toggleUnderline(){
        if(activeCell){
            if(!activeSheetObject[activeCell.id].isUnderlined) {
                activeCell.style.textDecoration = 'underline';
            }
            else{
                activeCell.style.textDecoration = 'none';
            }
            activeSheetObject[activeCell.id].isUnderlined = !activeSheetObject[activeCell.id].isUnderlined;
            activeCell.focus();
        }
    }


    // bug ix
    document.querySelectorAll('.color-prop').forEach(e => {
        e.addEventListener('click', (event) => event.stopPropagation());
    })
    function textColor(target){
        if(activeCell){
            let colorInput = target.value;
            activeSheetObject[activeCell.id].color = colorInput;
            activeCell.style.color = colorInput;
            activeCell.focus();
        }
    }
    function backgroundColor(target){
        if(activeCell){
            let colorInput = target.value;
            activeSheetObject[activeCell.id].backgroundColor = colorInput;
            activeCell.style.backgroundColor = colorInput;
            activeCell.focus();
        }
    }

    // bug fix
    document.querySelectorAll('.alignment').forEach(e => {
        e.addEventListener('click', (event) => {
            event.stopPropagation();
            let align = e.className.split(' ')[0];
            alignment(align);
        });
    })
    function alignment(align){
        if(activeCell){
            activeCell.style.textAlign = align;
            activeSheetObject[activeCell.id].textAlign = align;
            activeCell.focus();
        }
    }



    document.querySelector('.copy').addEventListener('click', (event) => {
        event.stopPropagation();
        if (activeCell) {
            navigator.clipboard.writeText(activeCell.innerText);
            activeCell.focus();
        }
    })

    document.querySelector('.cut').addEventListener('click', (event) => {
        event.stopPropagation();
        if (activeCell) {
            navigator.clipboard.writeText(activeCell.innerText);
            activeCell.innerText = '';
            activeCell.focus();
        }
    })

    document.querySelector('.paste').addEventListener('click', (event) => {
        event.stopPropagation();
        if (activeCell) {
            navigator.clipboard.readText().then((text) => {
                formula.value = text;
                activeCell.innerText = text;
            })
            activeCell.focus();
        }
    })

    downloadBtn.addEventListener("click",(e)=>{
        let jsonData = JSON.stringify(sheetsArray);
        let file = new Blob([jsonData],{type: "application/json"});
        let a = document.createElement("a");
        a.href = URL.createObjectURL(file);
        a.download = "SheetData.json";
        a.click();
    });

    openBtn.addEventListener("click",(e)=>{
        let input = document.createElement("input");
        input.setAttribute("type","file");
        input.click();

        input.addEventListener("change",(e)=>{
            let fr = new FileReader();
            let files = input.files;
            let fileObj = files[0];

            fr.readAsText(fileObj);
            fr.addEventListener("load",(e)=>{
                let readSheetData = JSON.parse(fr.result);
                readSheetData.forEach(e => {
                    document.querySelector('.new-sheet').click();
                    sheetsArray[activeSheetIndex] = e;
                    activeSheetObject = e;
                    changeSheet();
                })
            });
        });
    });

    formula.addEventListener('input', (event) => {
        const inputFormula = event.target.value;
        activeSheetObject[activeCell.id].formula = inputFormula; // Store formula
        const result = evaluateFormula(inputFormula); // Calculate result
        activeCell.innerText = result; // Update visible value
        activeSheetObject[activeCell.id].content = result; // Update content
    });
    function recalculateSheet() {
        for (let cellId in activeSheetObject) {
            const cell = activeSheetObject[cellId];
            if (cell.formula) {
                const result = evaluateFormula(cell.formula); // Recalculate formula
                document.getElementById(cellId).innerText = result; // Update visible value
                cell.content = result; // Update stored content
            }
        }
    }
        
// Sheet Navigation (Switching between sheets)
document.querySelectorAll('.sheet-menu').forEach(sheet => {
    sheet.addEventListener('click', function() {
        document.querySelectorAll('.sheet-menu').forEach(sheet => {
            sheet.classList.remove('active-sheet');
        });
        this.classList.add('active-sheet');

        activeSheetIndex = Number(this.id.slice(1)) - 1;  // Extract index
        activeSheetObject = sheetsArray[activeSheetIndex];
        changeSheet(); // Reflect changes on the sheet
    });
});


// Utility function to get values from the grid
function getCellValues() {
    let cells = document.querySelectorAll('td');
    let values = [];
    cells.forEach(cell => {
        let value = parseFloat(cell.textContent);
        if (!isNaN(value)) {
            values.push(value);
        }
    });
    return values;
}

// Formula evaluation function that handles cell references and operations like SUM, AVERAGE, etc.
function evaluateFormula(formula) {
    if (!formula.startsWith('=')) return formula; // Skip if not a formula

    const match = formula.match(/=(\w+)\(([\w\d:]+)\)/); // Extract function and range
    if (!match) return "Invalid Formula";

    const [_, func, range] = match; // `func` is the function, `range` is the cell range
    const cells = parseRange(range); // Convert the range into a list of cell IDs

    // Get values of the cells within the range
    const values = cells.map(cell => {
        const cellValue = activeSheetObject[cell]?.content; // Get the cell value
        return isNaN(Number(cellValue)) ? null : Number(cellValue); // Convert to number if valid
    }).filter(value => value !== null); // Remove null values

    return performOperation(func.toUpperCase(), values);
}

// Function to parse cell range (e.g., A1:B3) into an array of cell IDs (this would depend on your grid setup)
function parseRange(range) {
    // A simple mock for range parsing, depending on grid layout
    let [start, end] = range.split(':');
    // You can modify this function to support more complex range parsing
    return [start, end]; // Placeholder, adapt as necessary
}

// Perform operations like SUM, AVERAGE, MAX, MIN, and COUNT on the provided values
function performOperation(operation, values) {
    let result;

    switch (operation) {
        case 'SUM':
            result = calculateSum(values);
            break;
        case 'AVERAGE':
            result = calculateAverage(values);
            break;
        case 'MAX':
            result = calculateMax(values);
            break;
        case 'MIN':
            result = calculateMin(values);
            break;
        case 'COUNT':
            result = calculateCount(values);
            break;
        default:
            result = "Invalid Function";
    }

    return result;
}

// Helper functions for specific calculations
function calculateSum(values) {
    return values.reduce((acc, val) => acc + val, 0);
}

function calculateAverage(values) {
    return values.length > 0 ? values.reduce((acc, val) => acc + val, 0) / values.length : 0;
}

function calculateMax(values) {
    return values.length > 0 ? Math.max(...values) : 0;
}

function calculateMin(values) {
    return values.length > 0 ? Math.min(...values) : 0;
}

function calculateCount(values) {
    return values.length;
}

// Drag to Select Cells
let isMouseDown = false;
let startCell = null;
let endCell = null;

document.querySelectorAll('.cell').forEach(cell => {
    cell.addEventListener('mousedown', (event) => {
        isMouseDown = true;
        startCell = event.target;
        endCell = startCell;
        highlightSelection();
    });

    cell.addEventListener('mouseover', (event) => {
        if (isMouseDown) {
            endCell = event.target;
            highlightSelection();
        }
    });

    cell.addEventListener('mouseup', () => {
        isMouseDown = false;
        resetSelection();
    });
});

function highlightSelection() {
    const range = getRange(startCell, endCell);
    document.querySelectorAll('.cell').forEach(cell => {
        if (range.includes(cell)) {
            cell.classList.add('selected');
        } else {
            cell.classList.remove('selected');
        }
    });
}

function resetSelection() {
    document.querySelectorAll('.cell').forEach(cell => {
        cell.classList.remove('selected');
    });
}

function getRange(start, end) {
    const startRow = parseInt(start.id.match(/\d+/)[0]);
    const startCol = start.id.match(/[A-Z]+/)[0];
    const endRow = parseInt(end.id.match(/\d+/)[0]);
    const endCol = end.id.match(/[A-Z]+/)[0];

    const cells = [];
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
            cells.push(document.getElementById(String.fromCharCode(col) + row));
        }
    }
    return cells;
}


    // try for bug fix
    document.querySelector('body').addEventListener('click', () => {
        resetFunctionality();
    })

    // bug fix
    formula.addEventListener('click', (event) => event.stopPropagation());

    document.querySelectorAll('.select,.color-prop>*').forEach(e => {
        e.addEventListener('click', event => {
            event.stopPropagation();
        });
    })