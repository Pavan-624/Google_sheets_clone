const initialCellState = {
    fontFamily_data: 'monospace',
    fontSize_data: '14',
    isBold: false,
    isItalic: false,
    textAlign: 'start',
    isUnderlined: false,
    color: '#000000',
    backgroundColor: '#ffffff',
    content: '',
    formula: '', // New property for formula
    dependencies: [] // Tracks dependent cells
};

let sheetsArray = [];
let activeSheetIndex = -1;
let activeSheetObject = false;
let activeCell = false;

// Functionality elements
let fontFamilyBtn = document.querySelector('.font-family');
let fontSizeBtn = document.querySelector('.font-size');
let boldBtn = document.querySelector('.bold');
let italicBtn = document.querySelector('.italic');
let underlineBtn = document.querySelector('.underline');
let leftBtn = document.querySelector('.start');
let centerBtn = document.querySelector('.center');
let rightBtn = document.querySelector('.end');
let colorBtn = document.querySelector('#color');
let bgColorBtn = document.querySelector('#bgcolor');
let addressBar = document.querySelector('.address-bar');
let formula = document.querySelector('.formula-bar');
let downloadBtn = document.querySelector(".download");
let openBtn = document.querySelector(".open");

// Grid header row
let gridHeader = document.querySelector('.grid-header');

// Add header columns
let bold = document.createElement('div');
bold.className = 'grid-header-col';
bold.innerText = 'SL. NO.';
gridHeader.append(bold);

for (let i = 65; i <= 90; i++) {
    let bold = document.createElement('div');
    bold.className = 'grid-header-col';
    bold.innerText = String.fromCharCode(i);
    bold.id = String.fromCharCode(i);
    gridHeader.append(bold);
}

// Create grid
for (let i = 1; i <= 100; i++) {
    let newRow = document.createElement('div');
    newRow.className = 'row';
    document.querySelector('.grid').append(newRow);

    let bold = document.createElement('div');
    bold.className = 'grid-cell';
    bold.innerText = i;
    bold.id = i;
    newRow.append(bold);

    for (let j = 65; j <= 90; j++) {
        let cell = document.createElement('div');
        cell.className = 'grid-cell cell-focus';
        cell.id = String.fromCharCode(j) + i;
        cell.contentEditable = true;

        cell.addEventListener('click', (event) => {
            event.stopPropagation();
        });
        cell.addEventListener('focus', cellFocus);
        cell.addEventListener('focusout', cellFocusOut);
        cell.addEventListener('input', cellInput);

        newRow.append(cell);
    }
}

// Cell focus
function cellFocus(event) {
    let key = event.target.id;
    addressBar.innerHTML = event.target.id;
    activeCell = event.target;

    let activeBg = '#c9c8c8';
    let inactiveBg = '#ecf0f1';

    if (activeSheetObject[key]) {
        fontFamilyBtn.value = activeSheetObject[key].fontFamily_data;
        fontSizeBtn.value = activeSheetObject[key].fontSize_data;
        boldBtn.style.backgroundColor = activeSheetObject[key].isBold ? activeBg : inactiveBg;
        italicBtn.style.backgroundColor = activeSheetObject[key].isItalic ? activeBg : inactiveBg;
        underlineBtn.style.backgroundColor = activeSheetObject[key].isUnderlined ? activeBg : inactiveBg;
        setAlignmentBg(key, activeBg, inactiveBg);
        colorBtn.value = activeSheetObject[key].color;
        bgColorBtn.value = activeSheetObject[key].backgroundColor;

        formula.value = activeSheetObject[key].formula || activeCell.innerText;
    }

    document.getElementById(event.target.id.slice(0, 1)).classList.add('row-col-focus');
    document.getElementById(event.target.id.slice(1)).classList.add('row-col-focus');
}

// Handle cell input
function cellInput() {
    let key = activeCell.id;
    let value = activeCell.innerText;

    if (value.startsWith('=')) {
        activeSheetObject[key].formula = value;
        evaluateFormula(key, value);
    } else {
        activeSheetObject[key].content = value;
        activeSheetObject[key].formula = '';
    }
}

function handleFormulaInput(cell, formula) {
    cell.formula = formula; // Store the formula
    cell.dependencies = parseRange(formula.match(/\((.*?)\)/)[1]); // Extract dependencies
    cell.content = evaluateFormula(formula); // Evaluate and store the result
}
// Evaluate formula
function evaluateFormula(key, formula) {
    try {
        let dependencies = formula.match(/[A-Z][0-9]+/g) || [];
        activeSheetObject[key].dependencies = dependencies;

        let evaluatedValue = formula.replace(/[A-Z][0-9]+/g, (ref) => {
            return activeSheetObject[ref]?.content || 0;
        });

        activeSheetObject[key].content = eval(evaluatedValue);
        activeCell.innerText = activeSheetObject[key].content;

        updateDependentCells(key);
    } catch (error) {
        activeCell.innerText = 'ERROR';
    }
}

// Update dependent cells
function updateDependentCells(changedCellId) {
    for (let key in activeSheetObject) {
        if (activeSheetObject[key].dependencies.includes(changedCellId)) {
            evaluateFormula(key, activeSheetObject[key].formula);
        }
    }
}

// Set alignment background
function setAlignmentBg(key, activeBg, inactiveBg) {
    leftBtn.style.backgroundColor = inactiveBg;
    centerBtn.style.backgroundColor = inactiveBg;
    rightBtn.style.backgroundColor = inactiveBg;
    if (key) {
        document.querySelector('.' + activeSheetObject[key].textAlign).style.backgroundColor = activeBg;
    } else {
        leftBtn.style.backgroundColor = activeBg;
    }
}

// Cell focus out
function cellFocusOut(event) {
    document.getElementById(event.target.id.slice(0, 1)).classList.remove('row-col-focus');
    document.getElementById(event.target.id.slice(1)).classList.remove('row-col-focus');
}
