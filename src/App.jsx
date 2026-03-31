import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { CellDirective, CellsDirective, RowDirective, RowsDirective, SheetDirective, SheetsDirective, SpreadsheetComponent, getSheet } from '@syncfusion/ej2-react-spreadsheet';
import './App.css';
import { formatUnit, registerLicense } from '@syncfusion/ej2-base';

export default function App() {
    let spreadsheet;
    let budgetData = [
        { "Scenario": "Base", "Region": "South", "Branch": "Chennai", "Card_Type": "Platinum", "Cards_Issued": 2000, "Activation_Rate": 0.80, "Avg_Spend": 50000, "Revolving_Rate": 0.40, "Interest_Rate": 0.36 },
        { "Scenario": "Base", "Region": "South", "Branch": "Chennai", "Card_Type": "Gold", "Cards_Issued": 3000, "Activation_Rate": 0.75, "Avg_Spend": 30000, "Revolving_Rate": 0.35, "Interest_Rate": 0.32 },
        { "Scenario": "Base", "Region": "South", "Branch": "Chennai", "Card_Type": "Classic", "Cards_Issued": 2500, "Activation_Rate": 0.70, "Avg_Spend": 20000, "Revolving_Rate": 0.30, "Interest_Rate": 0.30 },
        { "Scenario": "Base", "Region": "South", "Branch": "Coimbatore", "Card_Type": "Platinum", "Cards_Issued": 1400, "Activation_Rate": 0.78, "Avg_Spend": 42000, "Revolving_Rate": 0.38, "Interest_Rate": 0.35 },
        { "Scenario": "Base", "Region": "South", "Branch": "Coimbatore", "Card_Type": "Gold", "Cards_Issued": 2000, "Activation_Rate": 0.72, "Avg_Spend": 28000, "Revolving_Rate": 0.34, "Interest_Rate": 0.32 },
        { "Scenario": "Base", "Region": "South", "Branch": "Coimbatore", "Card_Type": "Classic", "Cards_Issued": 2500, "Activation_Rate": 0.70, "Avg_Spend": 20000, "Revolving_Rate": 0.30, "Interest_Rate": 0.30 },
        { "Scenario": "Base", "Region": "South", "Branch": "Madurai", "Card_Type": "Platinum", "Cards_Issued": 1200, "Activation_Rate": 0.76, "Avg_Spend": 40000, "Revolving_Rate": 0.36, "Interest_Rate": 0.34 },
        { "Scenario": "Base", "Region": "South", "Branch": "Madurai", "Card_Type": "Gold", "Cards_Issued": 1500, "Activation_Rate": 0.65, "Avg_Spend": 25000, "Revolving_Rate": 0.38, "Interest_Rate": 0.34 },
        { "Scenario": "Base", "Region": "South", "Branch": "Madurai", "Card_Type": "Classic", "Cards_Issued": 1100, "Activation_Rate": 0.60, "Avg_Spend": 18000, "Revolving_Rate": 0.32, "Interest_Rate": 0.31 },
        { "Scenario": "Base", "Region": "South", "Branch": "Trichy", "Card_Type": "Platinum", "Cards_Issued": 1800, "Activation_Rate": 0.78, "Avg_Spend": 45000, "Revolving_Rate": 0.42, "Interest_Rate": 0.37 },
        { "Scenario": "Base", "Region": "South", "Branch": "Trichy", "Card_Type": "Gold", "Cards_Issued": 900, "Activation_Rate": 0.70, "Avg_Spend": 27000, "Revolving_Rate": 0.35, "Interest_Rate": 0.33 },
        { "Scenario": "Base", "Region": "South", "Branch": "Trichy", "Card_Type": "Classic", "Cards_Issued": 700, "Activation_Rate": 0.62, "Avg_Spend": 16000, "Revolving_Rate": 0.28, "Interest_Rate": 0.29 },
        { "Scenario": "Base", "Region": "West", "Branch": "Pune", "Card_Type": "Platinum", "Cards_Issued": 1600, "Activation_Rate": 0.74, "Avg_Spend": 44000, "Revolving_Rate": 0.36, "Interest_Rate": 0.33 },
        { "Scenario": "Base", "Region": "West", "Branch": "Pune", "Card_Type": "Gold", "Cards_Issued": 2200, "Activation_Rate": 0.72, "Avg_Spend": 28000, "Revolving_Rate": 0.33, "Interest_Rate": 0.31 },
        { "Scenario": "Base", "Region": "West", "Branch": "Pune", "Card_Type": "Classic", "Cards_Issued": 1400, "Activation_Rate": 0.68, "Avg_Spend": 21000, "Revolving_Rate": 0.30, "Interest_Rate": 0.29 },
        { "Scenario": "Base", "Region": "North", "Branch": "Delhi", "Card_Type": "Platinum", "Cards_Issued": 2100, "Activation_Rate": 0.77, "Avg_Spend": 47000, "Revolving_Rate": 0.43, "Interest_Rate": 0.36 },
        { "Scenario": "Base", "Region": "North", "Branch": "Delhi", "Card_Type": "Gold", "Cards_Issued": 1900, "Activation_Rate": 0.70, "Avg_Spend": 29000, "Revolving_Rate": 0.34, "Interest_Rate": 0.32 },
        { "Scenario": "Base", "Region": "North", "Branch": "Delhi", "Card_Type": "Classic", "Cards_Issued": 2700, "Activation_Rate": 0.68, "Avg_Spend": 22000, "Revolving_Rate": 0.29, "Interest_Rate": 0.28 },
        { "Scenario": "Base", "Region": "South", "Branch": "Kochi", "Card_Type": "Platinum", "Cards_Issued": 1000, "Activation_Rate": 0.70, "Avg_Spend": 38000, "Revolving_Rate": 0.34, "Interest_Rate": 0.32 },
        { "Scenario": "Base", "Region": "South", "Branch": "Kochi", "Card_Type": "Gold", "Cards_Issued": 1300, "Activation_Rate": 0.67, "Avg_Spend": 24000, "Revolving_Rate": 0.32, "Interest_Rate": 0.30 },
        { "Scenario": "Base", "Region": "South", "Branch": "Kochi", "Card_Type": "Classic", "Cards_Issued": 1600, "Activation_Rate": 0.66, "Avg_Spend": 21000, "Revolving_Rate": 0.31, "Interest_Rate": 0.29 },
        { "Scenario": "Base", "Region": "South", "Branch": "Bengaluru", "Card_Type": "Platinum", "Cards_Issued": 2400, "Activation_Rate": 0.82, "Avg_Spend": 48000, "Revolving_Rate": 0.44, "Interest_Rate": 0.38 },
        { "Scenario": "Base", "Region": "South", "Branch": "Bengaluru", "Card_Type": "Gold", "Cards_Issued": 2100, "Activation_Rate": 0.78, "Avg_Spend": 32000, "Revolving_Rate": 0.36, "Interest_Rate": 0.34 },
        { "Scenario": "Base", "Region": "South", "Branch": "Bengaluru", "Card_Type": "Classic", "Cards_Issued": 1800, "Activation_Rate": 0.70, "Avg_Spend": 23000, "Revolving_Rate": 0.33, "Interest_Rate": 0.31 },
        { "Scenario": "Base", "Region": "East", "Branch": "Kolkata", "Card_Type": "Platinum", "Cards_Issued": 1500, "Activation_Rate": 0.75, "Avg_Spend": 42000, "Revolving_Rate": 0.38, "Interest_Rate": 0.35 },
        { "Scenario": "Base", "Region": "East", "Branch": "Kolkata", "Card_Type": "Gold", "Cards_Issued": 2000, "Activation_Rate": 0.70, "Avg_Spend": 26000, "Revolving_Rate": 0.34, "Interest_Rate": 0.33 },
        { "Scenario": "Base", "Region": "East", "Branch": "Kolkata", "Card_Type": "Classic", "Cards_Issued": 1700, "Activation_Rate": 0.69, "Avg_Spend": 20000, "Revolving_Rate": 0.30, "Interest_Rate": 0.29 },
        { "Scenario": "Base", "Region": "West", "Branch": "Mumbai", "Card_Type": "Platinum", "Cards_Issued": 2300, "Activation_Rate": 0.79, "Avg_Spend": 50000, "Revolving_Rate": 0.41, "Interest_Rate": 0.37 },
        { "Scenario": "Base", "Region": "West", "Branch": "Mumbai", "Card_Type": "Gold", "Cards_Issued": 2500, "Activation_Rate": 0.76, "Avg_Spend": 33000, "Revolving_Rate": 0.35, "Interest_Rate": 0.34 },
        { "Scenario": "Base", "Region": "West", "Branch": "Mumbai", "Card_Type": "Classic", "Cards_Issued": 1900, "Activation_Rate": 0.70, "Avg_Spend": 21000, "Revolving_Rate": 0.32, "Interest_Rate": 0.30 }
    ];

    const onCreated = () => {
        spreadsheet.updateRange({ dataSource: budgetData, startCell: 'budgetSheet!A3' }, 0);
        spreadsheet.setColumnsWidth(120, ['budgetSheet!A:L', 'actuals!B:I', 'Variance!B:M', 'Dashboard!A:D']);
        spreadsheet.setRowsHeight(40, ['budgetSheet!3', 'actuals!3', 'actuals!B3:I3', 'Dashboard!9', 'Dashboard!21']);
        spreadsheet.setRowsHeight(25,['budgetSheet!4:33', 'actuals!4:10000', 'actuals!B4:I33', 'Variance!4:33','Dashboard!4:8', 'Dashboard!10:20', 'Dashboard!21:24']);
        budgetSheetCalculation(spreadsheet);
        actualsSheetCalculation(spreadsheet);
        varianceSheetCalculation(spreadsheet);
        dashboardCalculation(spreadsheet);
        spreadsheet.addDataValidation({ type: 'List', value1: 'Chennai,Coimbatore,Madurai,Trichy,Pune,Delhi, Kochi, Bengaluru, Kolkata, Mumbai', ignoreBlank: false }, 'budgetSheet!C4:C33');
        spreadsheet.addDataValidation({ type: 'List', value1: 'Platinum,Gold,Classic', ignoreBlank: false }, 'budgetSheet!D4:D33');
        spreadsheet.cellFormat({ textAlign: 'center' }, 'Variance!C4:M100');
        spreadsheet.cellFormat({ textAlign: 'center' }, 'actuals!C4:I1000');
        //#ffffcc
        spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'budgetSheet!A3:L3');
        spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'actuals!B3:I3');
        spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'Variance!B3:M3');
        spreadsheet.cellFormat({ backgroundColor: '#ffffcc' }, 'Dashboard!A3:B3 A9:D9 A21:D21');
        spreadsheet.cellFormat({ backgroundColor: '#b3ffb3' }, 'budgetSheet!A4:A33 C4:C33 E4:E33 G4:G33 I4:I33 K4:K33');
        //spreadsheet.cellFormat({ backgroundColor: '#b3ffb3' }, 'actuals!B4:B10000 D4:D10000 F4:F10000 H4:H10000');
        //spreadsheet.cellFormat({ backgroundColor: '#b3ffb3' }, 'Variance!B4:B33 D4:D33 F4:F33 H4:H33 J4:J33 L4:L33');
        spreadsheet.cellFormat({ backgroundColor: '#b3ffb3' }, 'Dashboard!B4:B7 B10:B19 B22:B24 D10:D19 D22:D24');
        spreadsheet.cellFormat({ backgroundColor: '#ccccff' }, 'Dashboard!A4:A7 A10:A19 A22:A24 C10:C19 C22:C24');
    }

    const budgetSheetCalculation = (spreadsheet) => {
        spreadsheet.updateCell({ value: 'Planned_spend' }, `budgetSheet!J3`);
        spreadsheet.updateCell({ value: 'Revolving_Balance' }, `budgetSheet!K3`);
        spreadsheet.updateCell({ value: 'Planned_Interest' }, `budgetSheet!L3`);
        for (let i = 4; i < 34; i++) {
            spreadsheet.updateCell({ formula: `=E${i}* F${i} * G${i}` }, `budgetSheet!J${i}`);
            spreadsheet.updateCell({ formula: `=J${i}* H${i}` }, `budgetSheet!K${i}`);
            spreadsheet.updateCell({ formula: `=K${i}* I${i}` }, `budgetSheet!L${i}`);
        }
        spreadsheet.numberFormat('$#,##0.00', 'budgetSheet!G1:G100 J1:L100');
        //apply formattings 
        //textAlign: 'center',
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', }, 'budgetSheet!A3:L3');
    }

    const actualsSheetCalculation = (spreadsheet) => {
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle' }, 'actuals!B3:I3');
        spreadsheet.cellFormat({ textAlign: 'left' }, 'actuals!B4:B1000');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!F1:F100');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!H1:H100');
        const branches = ['Chennai', 'Coimbatore', 'Madurai', 'Trichy', 'Pune',
            'Delhi', 'Kochi', 'Bengaluru', 'Kolkata', 'Mumbai'];
        const cardTypes = ['Platinum', 'Gold', 'Classic'];
        const yesNo = ['Yes', 'No'];
        const startDate = new Date(2024, 0, 1).getTime();
        const endDate = new Date(2025, 11, 31).getTime();
        const dateRange = endDate - startDate;
        const actualsData = [];
        for (let i = 0; i < 10000; i++) {
            const randomDate = new Date(startDate + Math.random() * dateRange);
            const branch = branches[Math.floor(Math.random() * branches.length)];
            const cardType = cardTypes[Math.floor(Math.random() * cardTypes.length)];
            const cardId = `CC${String(Math.floor(Math.random() * 9000000) + 1000000)}`;
            const transactionAmt = Math.floor(Math.random() * 99500) + 500;
            const revolvingFlag = yesNo[Math.floor(Math.random() * 2)];
            const interestCharged = revolvingFlag === 'Yes'
                ? parseFloat((transactionAmt * (0.28 + Math.random() * 0.10)).toFixed(2))
                : 0;
            const delinquencyFlag = yesNo[Math.floor(Math.random() * 2)];
            actualsData.push({
                'Date': randomDate.toLocaleDateString('en-IN'),
                'Branch': branch,
                'Card_Type': cardType,
                'Card_ID': cardId,
                'Transaction_Amount': transactionAmt,
                'Revolving_Flag': revolvingFlag,
                'Interest_Charged': interestCharged,
                'Delinquency_Flag': delinquencyFlag
            });
        }
        spreadsheet.updateRange({ dataSource: actualsData, startCell: 'actuals!B3' }, 1);
        spreadsheet.conditionalFormat({ type: 'BWRColorScale', range: 'Actuals!F4:F10000' });
        spreadsheet.conditionalFormat({ type: 'FiveArrows', range: 'Actuals!H4:H10000' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!G4:G10000', cFColor: 'RedT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!G4:G10000', format: { style: { color: 'Green' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!I4:I10000', cFColor: 'RedFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!I4:I10000', cFColor: 'GreenFT' });
    };

    const varianceSheetCalculation = (spreadsheet) => {
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', }, 'Variance!B3:M3');
        const headers = ['Branch', 'Card_Type', 'Planned_Spend', 'Actual_Spend', 'Variance', 'Variance %', 'Planned_Interest', 'Actual_Interest', 'Delinquency %', 'Run_Rate', 'Risk_Score', 'Insight'];
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
            spreadsheet.updateCell({ formula: `=SUMIFS(actuals!F4:F100, actuals!C4:C100, Variance!B${i}, actuals!D4:D100, Variance!C${i})` }, `Variance!E${i}`);
            spreadsheet.updateCell({ formula: `=Variance!D${i} - Variance!E${i}` }, `Variance!F${i}`);
            spreadsheet.updateCell({ formula: `=Variance!F${i} / Variance!D${i}` }, `Variance!G${i}`);
            spreadsheet.updateCell({ formula: `=SUMIFS(budgetSheet!L4:L33, budgetSheet!C4:C33, Variance!B${i}, budgetSheet!D4:D33, Variance!C${i})` }, `Variance!H${i}`);
            spreadsheet.updateCell({ formula: `=SUMIFS(actuals!H4:H100, actuals!C4:C100, Variance!B${i}, actuals!D4:D100, Variance!C${i})` }, `Variance!I${i}`);
            spreadsheet.updateCell({ formula: `=SUMIFS(actuals!F4:F100,actuals!C4:C100,Variance!B${i}, actuals!D4:D100 , Variance!C${i},actuals!I4:I100,"Yes")/SUMIFS(actuals!F4:F100,actuals!C4:C100,Variance!B${i}, actuals!D4:D100 , Variance!C${i})` }, `Variance!J${i}`);
            spreadsheet.updateCell({ formula: `=(Variance!E${i}/10)*30` }, `Variance!K${i}`);
            spreadsheet.updateCell({ formula: `=(budgetSheet!H${i}*0.4)+(Variance!J${i}*0.6)` }, `Variance!L${i}`);
            spreadsheet.updateCell({ formula: `=IF(Variance!L${i}>0.70,"High Risk",IF(Variance!L${i}>0.50,"Moderate Risk","Low Risk"))` }, `Variance!M${i}`);
        }
        spreadsheet.numberFormat('$#,##0.00', 'Variance!D4:F33');
        spreadsheet.numberFormat('$#,##0.00', 'Variance!H4:I33');
        spreadsheet.numberFormat('$#,##0.00', 'Variance!K4:K33');
        spreadsheet.numberFormat('0.00%', 'Variance!G4:G33');
        spreadsheet.numberFormat('0.00%', 'Variance!J4:J33');
        spreadsheet.conditionalFormat({ type: 'FiveQuarters', range: 'Variance!J4:J33' });
        spreadsheet.conditionalFormat({ type: 'ThreeFlags', range: 'Variance!L4:L33' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', cFColor: 'GreenFT', value: 'Low Risk', range: 'Variance!M4:M33'});
        spreadsheet.conditionalFormat({ type: 'EqualTo', cFColor: 'YellowFT', value: 'Moderate Risk', range: 'Variance!M4:M33'});
        spreadsheet.conditionalFormat({ type: 'EqualTo', cFColor: 'RedFT', value: 'High Risk', range: 'Variance!M4:M33'});
        spreadsheet.conditionalFormat({ type: 'BWRColorScale', range: 'Variance!K4:K33' });
    };

    const dashboardCalculation = (spreadsheet) =>{
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', }, 'Dashboard!A3:B3 A9:D9 A21:D21 ');
        const metrics = ['Key Metrics','Total Spend','Total Interest','ForeCast Spend','Delinquency %'];
        const locations = ['Bengaluru', 'Chennai', 'Coimbatore', 'Delhi', 'Kochi', 'Kolkata', 'Madurai', 'Mumbai', 'Pune', 'Trichy'];
        const cardTypes = ['Gold', 'Classic', 'Platinum'];
        let i = 3;
        metrics.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!A${i}`);
            i++;
        });
        spreadsheet.merge('Dashboard!A3:B3');
        spreadsheet.updateCell({formula:'=SUM(Variance!E4:E33)'},'Dashboard!B4');
        spreadsheet.updateCell({formula:'=SUM(Variance!I4:I33)'},'Dashboard!B5');
        spreadsheet.updateCell({formula:'=AVERAGE(Variance!J4:J33)'},'Dashboard!B7');
        spreadsheet.updateCell({formula:'=SUM(Variance!K4:K33)'},'Dashboard!B6');
        const branchSpend = ['Branch','Spend','Interest','Variance'];
        const cardSpend = ['Card Type', 'Spend', 'Interest', 'Variance'];
        branchSpend.forEach((val,i) => {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({value:val},`Dashboard!${col}9`);
        });
        cardSpend.forEach((val, i) => {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}21`);
        });
        let j = 10;
        locations.forEach((val)=>{
            spreadsheet.updateCell({value:val},`Dashboard!A${j}`);
            j++;
        });
        for(let k = 10; k < 20; k++) {
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!B4:B33,Dashboard!A${k},Variance!E4:E33` }, `Dashboard!B${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!B4:B33,Dashboard!A${k},Variance!I4:I33` }, `Dashboard!C${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!B4:B33,Dashboard!A${k},Variance!F4:F33` }, `Dashboard!D${k}`);
        }
        let cardTrackNumber = 22;
        cardTypes.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!A${cardTrackNumber}`);
            cardTrackNumber++;
        })
        for(let i = 22; i < 25 ; i++ ){
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C4:C33,Dashboard!A${i},Variance!E4:E33` }, `Dashboard!B${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C4:C33,Dashboard!A${i},Variance!I4:I33` }, `Dashboard!C${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C4:C33,Dashboard!A${i},Variance!F4:F33` }, `Dashboard!D${i}`);
        }
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B10:D24');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B4:B6');
        spreadsheet.numberFormat('0.00%', 'Dashboard!B7');
        
    }
    const chart1 = [{ type: 'Column', range: 'Dashboard!A9:C19', title: 'BRANCH WISE SUMMARY', theme: 'Bootstrap5', height: 280, width: 1080, top: 292, left: 485, id: 'Chart2', isSeriesInRows: false, legendSettings: { visible: false } }];
    const chart2 = [{ type: 'StackingLine', range: 'Dashboard!A21:C24', title: 'CARD TYPE SUMMARY', theme: 'Bootstrap5', height: 280, width: 540, top: 2, left: 485, id: 'Chart1', isSeriesInRows: true }];
    const chart3= [{ type: 'Doughnut', range: 'Dashboard!A4:B5', title: 'TOTAL SPEND VS TOTAL INTEREST', theme: 'Bootstrap5', height: 280, width: 540, top: 2, left: 1027, id: 'Chart3'}];
    return (
        <SpreadsheetComponent height={850} ref={(ssObj) => { spreadsheet = ssObj }} created={onCreated.bind(this)}>
            <SheetsDirective>
                <SheetDirective name='budgetSheet'></SheetDirective>
                <SheetDirective name='actuals'></SheetDirective>
                <SheetDirective name='Variance'></SheetDirective>
                <SheetDirective name='dashboard'>
                    <RowsDirective>
                        <RowDirective>
                            <CellsDirective>
                                <CellDirective chart={chart2}>
                                </CellDirective>
                                <CellDirective chart={chart1}>
                                </CellDirective>
                                <CellDirective chart={chart3}>
                                </CellDirective>
                            </CellsDirective>
                        </RowDirective>
                    </RowsDirective>
                </SheetDirective>
            </SheetsDirective>
        </SpreadsheetComponent>
    );
}
const root = createRoot(document.getElementById('root'));
root.render(<App />);