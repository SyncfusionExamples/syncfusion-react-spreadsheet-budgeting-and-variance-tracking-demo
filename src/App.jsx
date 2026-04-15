import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { CellDirective, CellsDirective, RowDirective, RowsDirective, SheetDirective, SheetsDirective, SpreadsheetComponent, getSheet } from '@syncfusion/ej2-react-spreadsheet';
import './App.css';

export default function App() {
    let spreadsheet;
    let budgetData = [
        { "Region": "South", "Branch": "Chennai", "Card Type": "Platinum", "Cards Issued": 2000, "Activation %": 0.80, "Average Spend": 450, "Revolving %": 0.40, "Interest %": 0.36 },
        { "Region": "South", "Branch": "Chennai", "Card Type": "Gold", "Cards Issued": 3000, "Activation %": 0.75, "Average Spend": 410, "Revolving %": 0.35, "Interest %": 0.32 },
        { "Region": "South", "Branch": "Chennai", "Card Type": "Classic", "Cards Issued": 2500, "Activation %": 0.70, "Average Spend": 395, "Revolving %": 0.30, "Interest %": 0.30 },
        { "Region": "South", "Branch": "Coimbatore", "Card Type": "Platinum", "Cards Issued": 1400, "Activation %": 0.78, "Average Spend": 445, "Revolving %": 0.38, "Interest %": 0.35 },
        { "Region": "South", "Branch": "Coimbatore", "Card Type": "Gold", "Cards Issued": 2000, "Activation %": 0.72, "Average Spend": 355, "Revolving %": 0.34, "Interest %": 0.32 },
        { "Region": "South", "Branch": "Coimbatore", "Card Type": "Classic", "Cards Issued": 2500, "Activation %": 0.70, "Average Spend": 390, "Revolving %": 0.30, "Interest %": 0.30 },
        { "Region": "South", "Branch": "Madurai", "Card Type": "Platinum", "Cards Issued": 1200, "Activation %": 0.76, "Average Spend": 340, "Revolving %": 0.36, "Interest %": 0.34 },
        { "Region": "South", "Branch": "Madurai", "Card Type": "Gold", "Cards Issued": 1500, "Activation %": 0.65, "Average Spend": 360, "Revolving %": 0.38, "Interest %": 0.34 },
        { "Region": "South", "Branch": "Madurai", "Card Type": "Classic", "Cards Issued": 1100, "Activation %": 0.60, "Average Spend": 385, "Revolving %": 0.32, "Interest %": 0.31 },
        { "Region": "South", "Branch": "Trichy", "Card Type": "Platinum", "Cards Issued": 1800, "Activation %": 0.78, "Average Spend": 450, "Revolving %": 0.42, "Interest %": 0.37 },
        { "Region": "South", "Branch": "Trichy", "Card Type": "Gold", "Cards Issued": 900, "Activation %": 0.70, "Average Spend": 410, "Revolving %": 0.35, "Interest %": 0.33 },
        { "Region": "South", "Branch": "Trichy", "Card Type": "Classic", "Cards Issued": 700, "Activation %": 0.62, "Average Spend": 390, "Revolving %": 0.28, "Interest %": 0.29 },
        { "Region": "West", "Branch": "Pune", "Card Type": "Platinum", "Cards Issued": 1600, "Activation %": 0.74, "Average Spend": 345, "Revolving %": 0.36, "Interest %": 0.33 },
        { "Region": "West", "Branch": "Pune", "Card Type": "Gold", "Cards Issued": 2200, "Activation %": 0.72, "Average Spend": 410, "Revolving %": 0.33, "Interest %": 0.31 },
        { "Region": "West", "Branch": "Pune", "Card Type": "Classic", "Cards Issued": 1400, "Activation %": 0.68, "Average Spend": 395, "Revolving %": 0.30, "Interest %": 0.29 },
        { "Region": "North", "Branch": "Delhi", "Card Type": "Platinum", "Cards Issued": 2100, "Activation %": 0.77, "Average Spend": 455, "Revolving %": 0.43, "Interest %": 0.36 },
        { "Region": "North", "Branch": "Delhi", "Card Type": "Gold", "Cards Issued": 1900, "Activation %": 0.70, "Average Spend": 315, "Revolving %": 0.34, "Interest %": 0.32 },
        { "Region": "North", "Branch": "Delhi", "Card Type": "Classic", "Cards Issued": 2700, "Activation %": 0.68, "Average Spend": 400, "Revolving %": 0.29, "Interest %": 0.28 },
        { "Region": "South", "Branch": "Kochi", "Card Type": "Platinum", "Cards Issued": 1000, "Activation %": 0.70, "Average Spend": 440, "Revolving %": 0.34, "Interest %": 0.32 },
        { "Region": "South", "Branch": "Kochi", "Card Type": "Gold", "Cards Issued": 1300, "Activation %": 0.67, "Average Spend": 405, "Revolving %": 0.32, "Interest %": 0.30 },
        { "Region": "South", "Branch": "Kochi", "Card Type": "Classic", "Cards Issued": 1600, "Activation %": 0.66, "Average Spend": 395, "Revolving %": 0.31, "Interest %": 0.29 },
        { "Region": "South", "Branch": "Bengaluru", "Card Type": "Platinum", "Cards Issued": 2400, "Activation %": 0.82, "Average Spend": 460, "Revolving %": 0.44, "Interest %": 0.38 },
        { "Region": "South", "Branch": "Bengaluru", "Card Type": "Gold", "Cards Issued": 2100, "Activation %": 0.78, "Average Spend": 330, "Revolving %": 0.36, "Interest %": 0.34 },
        { "Region": "South", "Branch": "Bengaluru", "Card Type": "Classic", "Cards Issued": 1800, "Activation %": 0.70, "Average Spend": 305, "Revolving %": 0.33, "Interest %": 0.31 },
        { "Region": "East", "Branch": "Kolkata", "Card Type": "Platinum", "Cards Issued": 1500, "Activation %": 0.75, "Average Spend": 445, "Revolving %": 0.38, "Interest %": 0.35 },
        { "Region": "East", "Branch": "Kolkata", "Card Type": "Gold", "Cards Issued": 2000, "Activation %": 0.70, "Average Spend": 410, "Revolving %": 0.34, "Interest %": 0.33 },
        { "Region": "East", "Branch": "Kolkata", "Card Type": "Classic", "Cards Issued": 1700, "Activation %": 0.69, "Average Spend": 400, "Revolving %": 0.30, "Interest %": 0.29 },
        { "Region": "West", "Branch": "Mumbai", "Card Type": "Platinum", "Cards Issued": 2300, "Activation %": 0.79, "Average Spend": 360, "Revolving %": 0.41, "Interest %": 0.37 },
        { "Region": "West", "Branch": "Mumbai", "Card Type": "Gold", "Cards Issued": 2500, "Activation %": 0.76, "Average Spend": 420, "Revolving %": 0.35, "Interest %": 0.34 },
        { "Region": "West", "Branch": "Mumbai", "Card Type": "Classic", "Cards Issued": 1900, "Activation %": 0.70, "Average Spend": 405, "Revolving %": 0.32, "Interest %": 0.30 }
    ];

    const actionBegin = (args) =>{
        console.log(args);
    }

    const actionComplete = (args) =>{
        console.log(args);
    }

    const columnLetterToIndex = (letters) => {
        let index = 0;
        for (let i = 0; i < letters.length; i++) {
            index = index * 26 + (letters.charCodeAt(i) - 64);
        }
        return index;
    };

    const onCellSave = (args) => {
        const address = args.address || args.range || '';
        if (!address) return;
        const [sheetPart, cellPart] = address.includes('!') ? address.split('!') : ['BudgetSheet', address];
        if (sheetPart.toLowerCase() === 'budgetsheet') {
            const m = (cellPart || '').match(/([A-Z]+)(\d+)/i);
            if (!m) return;
            const colLetters = m[1].toUpperCase();
            const row = parseInt(m[2], 10);
            const colIndex = columnLetterToIndex(colLetters);
            if (colIndex < 1 || colIndex > 9) {
                return;
            }
            spreadsheet.setBorder({ border: '1px solid #000000' }, `BudgetSheet!A${row}:L${row}`, 'Vertical');
            spreadsheet.addDataValidation({ type: 'List', value1: 'North,South,East,West', ignoreBlank: false }, `budgetSheet!B${row}`);
            const jFormula = `=E${row}*F${row}*G${row}`;
            const kFormula = `=J${row}*H${row}`;
            const lFormula = `=K${row}*I${row}`;
            spreadsheet.updateCell({ formula: jFormula }, `budgetSheet!J${row}`);
            spreadsheet.updateCell({ formula: kFormula }, `budgetSheet!K${row}`);
            spreadsheet.updateCell({ formula: lFormula }, `budgetSheet!L${row}`);
            spreadsheet.cellFormat({ backgroundColor: '#CDE5E7' }, `budgetSheet!C${row} E${row} G${row} I${row} K${row}`);
            if(colIndex == 9){
                for (let i = 4; i < row; i++) {
                        spreadsheet.updateCell({ formula: `=SUMIFS(budgetSheet!J4:J33, budgetSheet!C4:C33, Variance!B${i}, budgetSheet!D4:D33, Variance!C${i})` }, `Variance!D${i}`);
                        spreadsheet.updateCell({ formula: `=SUMIFS(actuals!G4:G1000, actuals!C4:C1000, Variance!B1000, actuals!D4:D1000, Variance!C${i})` }, `Variance!E${i}`);
                        spreadsheet.updateCell({ formula: `=Variance!D${i} - Variance!E${i}` }, `Variance!F${i}`);
                        spreadsheet.updateCell({ formula: `=Variance!F${i} / Variance!D${i}` }, `Variance!G${i}`);
                        spreadsheet.updateCell({ formula: `=SUMIFS(budgetSheet!L4:L33, budgetSheet!C4:C33, Variance!B${i}, budgetSheet!D4:D33, Variance!C${i})` }, `Variance!H${i}`);
                        spreadsheet.updateCell({ formula: `=SUMIFS(actuals!K4:K1000, actuals!C4:C1000, Variance!B${i}, actuals!D4:1000, Variance!C${i})` }, `Variance!I${i}`);
                        spreadsheet.updateCell({ formula: `=SUMIFS(actuals!H4:H1000,actuals!C4:C1000,Variance!B${i}, actuals!D4:D1000, Variance!C${i},actuals!L4:L1000,"Yes")/SUMIFS(actuals!H4:H1000,actuals!C4:C1000,Variance!B${i}, actuals!D4:D1000 , Variance!C${i})` }, `Variance!J${i}`);
                        //spreadsheet.updateCell({ formula: `=(Variance!E${i}/10)*30` }, `Variance!K${i}`);
                        spreadsheet.updateCell({ formula: `=(budgetSheet!H${i}*0.4)+(Variance!J${i}*0.6)` }, `Variance!K${i}`);
                        spreadsheet.updateCell({ formula: `=IF(Variance!K${i}>0.30,"High Risk",IF(Variance!K${i}>0.20,"Moderate Risk","Low Risk"))` }, `Variance!L${i}`);
                    }
            }
            return;
        }
        if (sheetPart.toLowerCase() === 'actuals') {
            // supports single cell or range like B4 or B4:G10
            const r = (cellPart || '').match(/^([A-Z]+)(\d+)(:([A-Z]+)(\d+))?$/i);
            if (!r) return;
            const startRow = parseInt(r[2], 10);
            const endRow = r[5] ? parseInt(r[5], 10) : startRow;
            for (let rr = startRow; rr <= endRow; rr++) {
                // read Branch (col C) and Card_Type (col D) from the sheet model if available
                let branchVal, cardTypeVal , outstandingBalance , transactionAmt;
                try {
                    const sheetModel = getSheet(spreadsheet, 2);
                    const rowIndex = rr - 1; // sheetModel rows are 0-based
                    // In your actuals layout startCell 'B3' => columns: B=Date, C=Branch, D=Card_Type, E=Card_ID, F=Transaction_Amount, G=Revolving_Flag, H=Interest_Charged
                    branchVal = sheetModel.rows?.[rowIndex]?.cells?.[2]?.value; // C => index 2
                    cardTypeVal = sheetModel.rows?.[rowIndex]?.cells?.[3]?.value; // D => index 3
                    outstandingBalance = sheetModel.rows?.[rowIndex]?.cells?.[4]?.value;
                    transactionAmt = sheetModel.rows?.[rowIndex]?.cells?.[5]?.value;
                } catch (e) {
                    // fallback: try reading visible DOM cell text
                    const cellC = spreadsheet.getCell?.({ rowIndex: rr - 1, colIndex: 2 }); // C
                    const cellD = spreadsheet.getCell?.({ rowIndex: rr - 1, colIndex: 3 }); // D
                    branchVal = branchVal || (cellC && cellC.innerText);
                    cardTypeVal = cardTypeVal || (cellD && cellD.innerText);
                }
                // compute rate and update H cell (Interest_Charged). Use a formula so it recalculates if F/G change.
                const rate = getInterestRate(String(branchVal || '').trim(), String(cardTypeVal || '').trim()) || 0;
                spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: `actuals!G${startRow}`, cFColor: 'RedT' });
                spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: `actuals!G${startRow}`, format: { style: { color: 'Green' } } });
                spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: `actuals!I${startRow}`, cFColor: 'RedFT' });
                spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: `actuals!I${startRow}`, cFColor: 'GreenFT' });
                spreadsheet.setBorder({ border: '1px solid #000000' }, `actuals!B${startRow}:I${startRow}`, 'Vertical');              
                //minimum due 
                spreadsheet.updateCell({ formula: `=(actuals!F${startRow} * 0.40` },`actuals!G${startRow}`);
                //paid minimum
                spreadsheet.updateCell({ formula: `=IF(AND(actuals!H${startRow} >=actuals!G${startRow},actuals!H${startRow} > 0),"Yes","No")` }, `actuals!I${startRow}`);
                //isCarryOver
                spreadsheet.updateCell({ formula: `=IF(actuals!F${startRow}-H${startRow} > 0,"Yes","No")` }, `actuals!J${startRow}`);
                //Interest amount
                spreadsheet.updateCell({ formula: `=(actuals!F${startRow}-H${startRow})*${rate}` }, `actuals!K${startRow}`);
                //isLate
                spreadsheet.updateCell({formula:`=IF(AND(actuals!F${startRow} > 0,actuals!I${startRow}="No"),"Yes","No")`},`actuals!L${startRow}`);
                //calculate variance sheet
                const col = r[1].toUpperCase();
                if (col == 'H') {
                    for (let i = 4; i < 33; i++) {
                        spreadsheet.updateCell({ formula: `=SUMIFS(actuals!G4:G${startRow}, actuals!C4:C${startRow}, Variance!B${i}, actuals!D4:D${startRow}, Variance!C${i})` }, `Variance!E${i}`);
                        spreadsheet.updateCell({ formula: `=Variance!D${i} - Variance!E${i}` }, `Variance!F${i}`);
                        spreadsheet.updateCell({ formula: `=Variance!F${i} / Variance!D${i}` }, `Variance!G${i}`);
                        spreadsheet.updateCell({ formula: `=SUMIFS(budgetSheet!L4:L33, budgetSheet!C4:C33, Variance!B${i}, budgetSheet!D4:D33, Variance!C${i})` }, `Variance!H${i}`);
                        spreadsheet.updateCell({ formula: `=SUMIFS(actuals!K4:K${startRow}, actuals!C4:C${startRow}, Variance!B${i}, actuals!D4:D${startRow}, Variance!C${i})` }, `Variance!I${i}`);
                        spreadsheet.updateCell({ formula: `=SUMIFS(actuals!H4:H${startRow},actuals!C4:C${startRow},Variance!B${i}, actuals!D4:D${startRow} , Variance!C${i},actuals!L4:L${startRow},"Yes")/SUMIFS(actuals!H4:H${startRow},actuals!C4:C${startRow},Variance!B${i}, actuals!D4:D${startRow} , Variance!C${i})` }, `Variance!J${i}`);
                        spreadsheet.updateCell({ formula: `=(budgetSheet!H${i}*0.4)+(Variance!J${i}*0.6)` }, `Variance!K${i}`);
                        spreadsheet.updateCell({ formula: `=IF(Variance!K${i}>0.30,"High Risk",IF(Variance!K${i}>0.20,"Moderate Risk","Low Risk"))` }, `Variance!L${i}`);
                    }
                }
            }
        }
    };

    const onCreated = () => {
        spreadsheet.updateRange({ dataSource: budgetData, startCell: 'budgetSheet!B3' }, 1);
        spreadsheet.setColumnsWidth(135, ['budgetSheet!B:L']);
        spreadsheet.setColumnsWidth(130, ['budgetSheet!B:L', 'actuals!B:L', 'Variance!B:M']);
        spreadsheet.setColumnsWidth(120, ['Dashboard!A:E']);
        spreadsheet.setRowsHeight(40, ['budgetSheet!3', 'actuals!3', 'Variance!3', 'actuals!B3:I3', 'Dashboard!9', 'Dashboard!4', 'Dashboard!21', 'Dashboard!27']);
        spreadsheet.setRowsHeight(25, ['budgetSheet!4:33', 'actuals!4:10000', 'actuals!B4:I33', 'Variance!4:33', 'Dashboard!5:8', 'Dashboard!10:20', 'Dashboard!22:24', 'Dashboard!28:37']);
        setTimeout(() => {
            budgetSheetCalculation(spreadsheet);
        }, 100);
        actualsSheetCalculation(spreadsheet);
        varianceSheetCalculation(spreadsheet);
        spreadsheet.cellFormat({ textAlign: 'center' }, 'Variance!C3:M100');
        spreadsheet.cellFormat({ textAlign: 'center' }, 'actuals!C3:L1000');
        spreadsheet.cellFormat({ textAlign: 'center' }, 'budgetSheet!A3:L33');
        spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'budgetSheet!B3:L3');
        spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'actuals!B3:L3');
        spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'Variance!B3:L3');
        setTimeout(() => {
            dashboardCalculation(spreadsheet);
            spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'Dashboard!A4:B4 D4:E4 A9:E9 A21:E21 A27:B27 D27:E27');
            spreadsheet.cellFormat({ backgroundColor: '#CDE5E7' }, 'Dashboard!B5:B6 B28:B37 D28:D30 E5:E7 B10:B19 B22:B24 D10:D19 D22:D24');
            spreadsheet.cellFormat({ backgroundColor: '#5FA4A9 ' }, 'Dashboard!A5:A6 A28:A37 D5:D7 E28:E30 A10:A19 A22:A24 C10:C19 C22:C24 E10:E19 E22:E24');
            spreadsheet.setBorder({ border: '1px solid #000000' }, 'Dashboard!B5:B6 B28:B37 D28:D30 E5:E7 B10:B19 B22:B24 D10:D19 D22:D24', 'Vertical');
            spreadsheet.setBorder({ border: '1px solid #000000' }, 'Dashboard!A5:A6 A28:A37 D5:D7 E28:E30 A10:A19 A22:A24 C10:C19 C22:C24 E10:E19 E22:E24', 'Vertical');
        }, 100);
    }

    const budgetSheetCalculation = (spreadsheet) => {
        spreadsheet.updateCell({ value:'Credit Card Expense Budget'},'budgetSheet!B1');
        spreadsheet.updateCell({ value: 'Expected Spend' }, `budgetSheet!J3`);
        spreadsheet.updateCell({ value: 'Revolving Balance' }, `budgetSheet!K3`);
        spreadsheet.updateCell({ value: 'Expected Interest' }, `budgetSheet!L3`);
        spreadsheet.numberFormat('$#,##0.00', 'budgetSheet!G1:G33');
        spreadsheet.numberFormat('$#,##0.00', 'budgetSheet!J1:L33');
        spreadsheet.merge('budgetSheet!B1:L2');
        spreadsheet.cellFormat({ backgroundColor: '#CDE5E7' }, 'budgetSheet!C4:C33 E4:E33 G4:G33 I4:I33 K4:K33');
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign:'center' }, 'budgetSheet!B1:L2 B3:L3');
        const sheetModel = getSheet(spreadsheet, 1);
        for (let i = 4; i < 34; i++) {
            const regionCell = sheetModel.rows[i - 1].cells[1].value;
            const branchesForRegion = regionBranches[regionCell];
            spreadsheet.addDataValidation(
                { type: 'List', value1: branchesForRegion.join(','), ignoreBlank: false },
                `budgetSheet!C${i}`
            );
            spreadsheet.updateCell({ formula: `=E${i}* F${i} * G${i}` }, `budgetSheet!J${i}`);
            spreadsheet.updateCell({ formula: `=J${i}* H${i}` }, `budgetSheet!K${i}`);
            spreadsheet.updateCell({ formula: `=K${i}* I${i}` }, `budgetSheet!L${i}`);
        }
        spreadsheet.addDataValidation({ type: 'List', value1: 'North,South,East,West', ignoreBlank: false }, 'budgetSheet!B4:B33');
        spreadsheet.protectSheet(1, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    }
    const interestRateMatrix = {
            Chennai:   { Platinum: 0.34, Gold: 0.32, Classic: 0.30 },
            Coimbatore:{ Platinum: 0.35, Gold: 0.32, Classic: 0.30 },
            Madurai:   { Platinum: 0.34, Gold: 0.34, Classic: 0.31 },
            Trichy:    { Platinum: 0.37, Gold: 0.33, Classic: 0.29 },
            Pune:      { Platinum: 0.33, Gold: 0.31, Classic: 0.29 },
            Delhi:     { Platinum: 0.36, Gold: 0.32, Classic: 0.28 },
            Kochi:     { Platinum: 0.32, Gold: 0.30, Classic: 0.29 },
            Bengaluru: { Platinum: 0.38, Gold: 0.34, Classic: 0.31 },
            Kolkata:   { Platinum: 0.35, Gold: 0.33, Classic: 0.29 },
            Mumbai:    { Platinum: 0.37, Gold: 0.34, Classic: 0.30 }
        };

    const regionBranches = {
        South: ['Chennai','Coimbatore','Madurai','Trichy','Kochi','Bengaluru'],
        West:  ['Pune','Mumbai'],
        North: ['Delhi'],
        East:  ['Kolkata']
    };

    const actualDataSource = () => {
        const branches = ['Chennai', 'Coimbatore', 'Madurai', 'Trichy', 'Pune',
            'Delhi', 'Kochi', 'Bengaluru', 'Kolkata', 'Mumbai'];
        const cardTypes = ['Platinum', 'Gold', 'Classic'];
        const yesNo = ['Yes', 'No'];
        const startDate = new Date(2025, 0, 1).getTime();
        const endDate = new Date(2025, 12, 31).getTime();
        const dateRange = endDate - startDate;
        const actualsData = [];
        for (let i = 0; i < 100; i++) {
            const randomDate = new Date(startDate + Math.random() * dateRange);
            const branch = branches[Math.floor(Math.random() * branches.length)];
            const cardType = cardTypes[Math.floor(Math.random() * cardTypes.length)];
            const cardId = `CC${String(Math.floor(Math.random() * 9000000) + 1000000)}`;
            const maxOutstanding = 50000;
            const outstandingBalance = Math.floor(Math.random() * (maxOutstanding + 1)); // can be 0
            const minimumPayment = +(outstandingBalance * 0.40).toFixed(2);
            let transactionAmt = 0;
            let revolvingFlag = 'No';
            let interestCharged = 0;
            let minimumPaymentMet = 'No';
            let delinquencyFlag = 'No';
            if (outstandingBalance === 0) {
                // no balance -> no transaction, no interest, no delinquency
                transactionAmt = 0;
                minimumPaymentMet = 'No';
                revolvingFlag = 'No';
                interestCharged = 0;
                delinquencyFlag = 'No';
            } else {
                const interestRate = getInterestRate(branch, cardType) || 0;
                // 20% chance to pay full outstanding (remaining === 0)
                if (Math.random() < 0.20) {
                    transactionAmt = outstandingBalance; // remaining = 0
                } else {
                    // ensure remaining > 0 by picking 1..outstandingBalance-1 when possible
                    if (outstandingBalance > 1) {
                        transactionAmt = Math.floor(Math.random() * (outstandingBalance - 1)) + 1; // 1 .. outstandingBalance-1
                    } else {
                        transactionAmt = 1; // outstandingBalance == 1 -> pay 1 (remaining 0) but that case is rare
                    }
                }
                const remaining = outstandingBalance - transactionAmt;
                minimumPaymentMet = (transactionAmt >= minimumPayment && transactionAmt > 0) ? 'Yes' : 'No';
                if (remaining === 0) {
                    revolvingFlag = 'No';
                    interestCharged = 0;
                } else {
                    revolvingFlag = 'Yes';
                    interestCharged = parseFloat((remaining * interestRate).toFixed(2));
                }
                delinquencyFlag = (outstandingBalance > 0 && minimumPaymentMet === 'No') ? 'Yes' : 'No';
            }
            actualsData.push({
                'Date': randomDate.toLocaleDateString('en-IN'),
                'Branch': branch,
                'Card Type': cardType,
                'Card ID': cardId,
                'Balance Due': outstandingBalance,
                'Minimum Due': minimumPayment,
                'Credit Amount': transactionAmt,
                'Paid Minimum': minimumPaymentMet,
                'Balance Carried': revolvingFlag,
                'Interest Amount': interestCharged,
                'Payment Due': delinquencyFlag
            });
        }
        return actualsData;
    }
    const getInterestRate = (branch, cardType) => {
        const base = interestRateMatrix[branch]?.[cardType];
        return base;
    };
    const actualsSheetCalculation = (spreadsheet) => {
        spreadsheet.updateCell({ value:'Credit Card Expense Summary'},'actuals!B1');
        spreadsheet.merge('actuals!B1:L2');
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign: 'center' }, 'actuals!B1:L2 B3:L3');
        spreadsheet.cellFormat({ textAlign: 'left' }, 'actuals!B4:B1000');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!F1:H1000');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!K1:K1000');
        spreadsheet.updateRange({ dataSource: actualDataSource(), startCell: 'actuals!B3' }, 2);
        spreadsheet.addDataValidation({ type: 'List', value1: 'Yes,No', ignoreBlank: false }, 'actuals!I4:J1000');
        spreadsheet.addDataValidation({ type: 'List', value1: 'Yes,No', ignoreBlank: false }, 'actuals!L4:L1000');
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!H4:H1000', cFColor: 'RedT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!H4:H1000', format: { style: { color: 'Green' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!J4:J1000', cFColor: 'RedFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!J4:J1000', cFColor: 'GreenFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!I4:I1000', cFColor: 'RedFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!I4:I1000', cFColor: 'GreenFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!I4:I1000', cFColor: 'RedFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!I4:I1000', cFColor: 'GreenFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!L4:L1000', cFColor: 'RedFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!L4:L1000', cFColor: 'GreenFT' });
        spreadsheet.protectSheet(2, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    };

    const varianceSheetCalculation = (spreadsheet) => {
        spreadsheet.updateCell({ value:'Credit Card Expense Variance Analysis'},'Variance!B1');
        spreadsheet.merge('Variance!B1:L2');
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign:'center' }, 'Variance!B1:L2 B3:M3');
        const headers = ['Branch', 'Card Type', 'Expected Spend', 'Actual Spend', 'Variance', 'Variance %', 'Expected Interest', 'Actual Interest', 'Payment Risk %', 'Risk Score', 'Risk Level']; //'Run_Rate',
        const locations = ['Bengaluru', 'Chennai', 'Coimbatore', 'Delhi', 'Kochi', 'Kolkata', 'Madurai', 'Mumbai', 'Pune', 'Trichy'];
        headers.forEach((val, i) => {
            const col = String.fromCharCode(66 + i);
            spreadsheet.updateCell({ value: val }, `variance!${col}3`);
        });
        const cardTypes = ['Gold', 'Classic', 'Platinum'];
        let row = 4;
        locations.forEach((loc) => {
            cardTypes.forEach((cardVal) => {
                spreadsheet.updateCell({ value: loc }, `variance!B${row}`);
                spreadsheet.updateCell({ value: cardVal }, `variance!C${row}`);
                row++;
            });
        });
        let i = 4;
        for (let i = 4; i < row; i++) {
            spreadsheet.updateCell({ formula: `=SUMIFS(budgetSheet!J4:J33, budgetSheet!C4:C33, Variance!B${i}, budgetSheet!D4:D33, Variance!C${i})` }, `Variance!D${i}`);
            spreadsheet.updateCell({ formula: `=SUMIFS(actuals!G4:G1000, actuals!C4:C1000, Variance!B${i}, actuals!D4:D1000, Variance!C${i})` }, `Variance!E${i}`);
            spreadsheet.updateCell({ formula: `=Variance!D${i} - Variance!E${i}` }, `Variance!F${i}`);
            spreadsheet.updateCell({ formula: `=Variance!F${i} / Variance!D${i}` }, `Variance!G${i}`);
            spreadsheet.updateCell({ formula: `=SUMIFS(budgetSheet!L4:L33, budgetSheet!C4:C33, Variance!B${i}, budgetSheet!D4:D33, Variance!C${i})` }, `Variance!H${i}`);
            spreadsheet.updateCell({ formula: `=SUMIFS(actuals!K4:K1000, actuals!C4:C1000, Variance!B${i}, actuals!D4:D1000, Variance!C${i})` }, `Variance!I${i}`);
            spreadsheet.updateCell({ formula: `=SUMIFS(actuals!H4:H1000,actuals!C4:C1000,Variance!B${i}, actuals!D4:D1000 , Variance!C${i},actuals!L4:L1000,"Yes")/SUMIFS(actuals!H4:H1000,actuals!C4:C1000,Variance!B${i}, actuals!D4:D1000 , Variance!C${i})` }, `Variance!J${i}`);
            spreadsheet.updateCell({ formula: `=(budgetSheet!H${i}*0.4)+(Variance!J${i}*0.6)` }, `Variance!K${i}`);
            spreadsheet.updateCell({ formula: `=IF(Variance!K${i}>0.30,"High Risk",IF(Variance!K${i}>0.20,"Moderate Risk","Low Risk"))` }, `Variance!L${i}`);
        }
        spreadsheet.numberFormat('$#,##0.00', 'Variance!D4:F33');
        spreadsheet.numberFormat('$#,##0.00', 'Variance!H4:I33');
        spreadsheet.numberFormat('0.00%', 'Variance!K4:K33');
        spreadsheet.numberFormat('0.00%', 'Variance!G4:G33');
        spreadsheet.numberFormat('0.00%', 'Variance!J4:J33');
        spreadsheet.conditionalFormat({ type: 'FiveQuarters', range: 'Variance!J4:J33' });
        spreadsheet.conditionalFormat({ type: 'ThreeArrows', range: 'Variance!L4:L33' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', cFColor: 'GreenFT', value: 'Low Risk', range: 'Variance!L4:L33' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', cFColor: 'YellowFT', value: 'Moderate Risk', range: 'Variance!L4:L33' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', cFColor: 'RedFT', value: 'High Risk', range: 'Variance!L4:L33' });
        spreadsheet.protectSheet(3, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    };

    const dashboardCalculation = (spreadsheet) => {
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign: 'center' }, 'Dashboard!A2:E2 A4:E4 A9:E9 A21:E21 A27:E27');
        const metrics = ['Total Spend', 'Total Interest'];
        const locations = ['Bengaluru', 'Chennai', 'Coimbatore', 'Delhi', 'Kochi', 'Kolkata', 'Madurai', 'Mumbai', 'Pune', 'Trichy'];
        const cardTypes = ['Gold', 'Classic', 'Platinum'];
        let i = 5;
        metrics.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!A${i}`);
            i++;
        });
        spreadsheet.merge('Dashboard!A4:B4');
        spreadsheet.merge('Dashboard!D4:E4');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'BudgetSheet!B3:L3');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'Actuals!B3:L3');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'Variance!B3:L3');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'BudgetSheet!B3:L33', 'Vertical');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'Actuals!B3:L1000', 'Vertical');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'Variance!B3:L33', 'Vertical');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'Dashboard!A4:B4 D4:E4 A9:E9 A21:E21 A27:B27 D27:E27', 'Outer');
        spreadsheet.setBorder({ border: '1px solid #000000' }, 'Dashboard!A4:B4 D4:E4 A9:E9 A21:E21 A27:B27 D27:E27', 'Vertical');
        spreadsheet.updateCell({ formula: '=SUM(Variance!E4:E33)' }, 'Dashboard!B5');
        spreadsheet.updateCell({ formula: '=SUM(Variance!I4:I33)' }, 'Dashboard!B6');
        const branchSpend = ['Branch', 'Actual Spend', 'Actual Interest', 'Variance', 'Planned Spend'];
        const cardSpend = ['Card Type', 'Actual Spend', 'Actual Interest', 'Variance', 'Planned Spend'];
        const cardRisk = ['Low Risk', 'Moderate Risk', 'High Risk'];
        const deliquentCalc = ['Branch', 'Payment Risk'];
        const deliquentCalcCard = ['Card', 'Payment Risk'];
        branchSpend.forEach((val, i) => {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}9`);
        });
        cardSpend.forEach((val, i) => {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}21`);
        });
        let j = 10;
        let newLocation = 28;
        locations.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!A${j}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!A${newLocation}`);
            spreadsheet.updateCell({ formula: `=AVERAGEIF(Variance!B4:B33,A${newLocation},Variance!J4:J33)` }, `Dashboard!B${newLocation}`);
            newLocation++;
            j++;
        });
        for (let k = 10; k < 20; k++) {
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!B4:B33,Dashboard!A${k},Variance!E4:E33` }, `Dashboard!B${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!B4:B33,Dashboard!A${k},Variance!I4:I33` }, `Dashboard!C${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!B4:B33,Dashboard!A${k},Variance!F4:F33` }, `Dashboard!D${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(budgetSheet!C4:C33,Dashboard!A${k},budgetSheet!J4:J33` }, `Dashboard!E${k}`);
        }
        let cardTrackNumber = 22;
        cardTypes.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!A${cardTrackNumber}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!D${cardTrackNumber + 6}`);
            spreadsheet.updateCell({ formula: `=AVERAGEIF(Variance!C4:C33,D${cardTrackNumber + 6} , Variance!J4:J33)` }, `Dashboard!E${cardTrackNumber + 6}`);
            cardTrackNumber++;
        })
        for (let i = 22; i < 25; i++) {
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C4:C33,Dashboard!A${i},Variance!E4:E33` }, `Dashboard!B${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C4:C33,Dashboard!A${i},Variance!I4:I33` }, `Dashboard!C${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C4:C33,Dashboard!A${i},Variance!F4:F33` }, `Dashboard!D${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(budgetSheet!D4:D33,Dashboard!A${i},budgetSheet!J4:J33` }, `Dashboard!E${i}`);
        }
        deliquentCalc.forEach((val, i) => {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}27`);
        })
        deliquentCalcCard.forEach((val, i) => {
            const col = String.fromCharCode(68 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}27`);
        })
        spreadsheet.updateCell({ value: 'Key Metrics' }, 'Dashboard!A4');
        spreadsheet.updateCell({ value: 'Credit Card Risk State' }, 'Dashboard!D4');
        let cardRiskNumber = 5;
        cardRisk.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!D${cardRiskNumber}`);
            spreadsheet.updateCell({ formula: `=COUNTIF(Variance!L4:L33,Dashboard!D${cardRiskNumber})` }, `Dashboard!E${cardRiskNumber}`);
            cardRiskNumber++;
        });
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B10:E24');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B4:B6');
        spreadsheet.numberFormat('0.00%', 'Dashboard!B28:B37');
        spreadsheet.numberFormat('0.00%', 'Dashboard!E28:E30');
        spreadsheet.insertChart(chart1);
        spreadsheet.insertChart(chart2);
        spreadsheet.insertChart(chart3);
        spreadsheet.insertChart(chart4);
        spreadsheet.insertChart(chart5);
        spreadsheet.insertChart(chart6);
        spreadsheet.protectSheet(0, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    }
    const chart1 = [{ type: 'Column', range: 'Dashboard!A9:C19', title: 'BRANCH WISE SUMMARY', theme: 'Bootstrap5', left: 602, top: 282, width: 1194, height: 416, id: 'Chart1', isSeriesInRows: false, legendSettings: { visible: false } }];
    const chart2 = [{ type: 'StackingLine', range: 'Dashboard!A21:C24', title: 'CARD TYPE SUMMARY', theme: 'Bootstrap5', height: 280, left: 602, top: 701, width: 585, id: 'Chart2', isSeriesInRows: true }];
    const chart3 = [{ type: 'Doughnut', range: 'Dashboard!A5:B6', title: 'TOTAL SPEND VS TOTAL INTEREST', theme: 'Bootstrap5', height: 280, left: 608, top: 0, width: 365, id: 'Chart3' }];
    const chart4 = [{ type: 'Pie', range: 'Dashboard!A27:B37', title: 'PAYMENT RISK BASED ON LOCATION', theme: 'Bootstrap5', height: 279, left: 1407, width: 507, top: 0, id: 'Chart4' }];
    const chart5 = [{ type: 'Area', range: 'Dashboard!D27:E30', title: 'PAYMENT RISK BASED ON CARD', theme: 'Bootstrap5', top: 0, width: 430, left: 976, height: 280, id: 'Chart5' }];
    const chart6 = [{ type: 'Bar', range: 'Dashboard!D5:E7', title: 'RISK INSIGHTS OF CREDIT CARD', theme: 'Bootstrap5', height: 280, left: 1189, top: 701, width: 585, id: 'Chart6' }];
    return (
        <SpreadsheetComponent height={950} ref={(ssObj) => { spreadsheet = ssObj }} created={onCreated.bind(this)} cellSave={onCellSave} actionComplete={actionComplete} actionBegin={actionBegin}
        openUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/open'
        saveUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/save'
        >
            <SheetsDirective>
                <SheetDirective name='Dashboard' showGridLines={false}></SheetDirective>
                <SheetDirective name='BudgetSheet' showGridLines={false}></SheetDirective>
                <SheetDirective name='Actuals' showGridLines={false}></SheetDirective>
                <SheetDirective name='Variance' showGridLines={false}></SheetDirective>
            </SheetsDirective>
        </SpreadsheetComponent>
    );
}