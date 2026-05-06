import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { CellDirective, CellsDirective, ColumnDirective, ColumnsDirective, getColumnHeaderText, getSheet, RangeDirective, RangesDirective, RowDirective, RowsDirective, setCell, setColumn, SheetDirective, SheetsDirective, sheetTabs, SpreadsheetComponent,
} from '@syncfusion/ej2-react-spreadsheet';
import { DropDownListComponent } from '@syncfusion/ej2-react-dropdowns';
import { formatUnit, registerLicense } from '@syncfusion/ej2-base';
import { budgetImage, interestImage, varianceImage, moneyImage} from './dataSource'; //incomeGraph,
registerLicense("IAk8BicRIAEqCzQhAR8kAxMHIgRJXmBXf013TmJbYF1xa1xPaVVBRXdVR1RpTHdDFjhoW39cdXVXRGNdUkZzXkpea0B8WHVG");

let budgetData = [
    { "Region": "South", "Branch": "Dallas", "Card Type": "Platinum", "Cards Issued": 2000, "Activation %": 0.80, "Average Spend": 450, "Revolving %": 0.40, "Interest %": 0.36 },
    { "Region": "South", "Branch": "Dallas", "Card Type": "Gold", "Cards Issued": 3000, "Activation %": 0.75, "Average Spend": 410, "Revolving %": 0.35, "Interest %": 0.32 },
    { "Region": "South", "Branch": "Dallas", "Card Type": "Classic", "Cards Issued": 2500, "Activation %": 0.70, "Average Spend": 395, "Revolving %": 0.30, "Interest %": 0.30 },
    { "Region": "South", "Branch": "Austin", "Card Type": "Platinum", "Cards Issued": 1400, "Activation %": 0.78, "Average Spend": 445, "Revolving %": 0.38, "Interest %": 0.35 },
    { "Region": "South", "Branch": "Austin", "Card Type": "Gold", "Cards Issued": 2000, "Activation %": 0.72, "Average Spend": 355, "Revolving %": 0.34, "Interest %": 0.32 },
    { "Region": "South", "Branch": "Austin", "Card Type": "Classic", "Cards Issued": 2500, "Activation %": 0.70, "Average Spend": 390, "Revolving %": 0.30, "Interest %": 0.30 },
    { "Region": "South", "Branch": "San Antonio", "Card Type": "Platinum", "Cards Issued": 1200, "Activation %": 0.76, "Average Spend": 340, "Revolving %": 0.36, "Interest %": 0.34 },
    { "Region": "South", "Branch": "San Antonio", "Card Type": "Gold", "Cards Issued": 1500, "Activation %": 0.65, "Average Spend": 360, "Revolving %": 0.38, "Interest %": 0.34 },
    { "Region": "South", "Branch": "San Antonio", "Card Type": "Classic", "Cards Issued": 1100, "Activation %": 0.60, "Average Spend": 385, "Revolving %": 0.32, "Interest %": 0.31 },
    { "Region": "South", "Branch": "Houston", "Card Type": "Platinum", "Cards Issued": 1800, "Activation %": 0.78, "Average Spend": 450, "Revolving %": 0.42, "Interest %": 0.37 },
    { "Region": "South", "Branch": "Houston", "Card Type": "Gold", "Cards Issued": 900, "Activation %": 0.70, "Average Spend": 410, "Revolving %": 0.35, "Interest %": 0.33 },
    { "Region": "South", "Branch": "Houston", "Card Type": "Classic", "Cards Issued": 700, "Activation %": 0.62, "Average Spend": 390, "Revolving %": 0.28, "Interest %": 0.29 },
    { "Region": "West", "Branch": "Phoenix", "Card Type": "Platinum", "Cards Issued": 1600, "Activation %": 0.74, "Average Spend": 345, "Revolving %": 0.36, "Interest %": 0.33 },
    { "Region": "West", "Branch": "Phoenix", "Card Type": "Gold", "Cards Issued": 2200, "Activation %": 0.72, "Average Spend": 410, "Revolving %": 0.33, "Interest %": 0.31 },
    { "Region": "West", "Branch": "Phoenix", "Card Type": "Classic", "Cards Issued": 1400, "Activation %": 0.68, "Average Spend": 395, "Revolving %": 0.30, "Interest %": 0.29 },
    { "Region": "Northeast", "Branch": "New York", "Card Type": "Platinum", "Cards Issued": 2100, "Activation %": 0.77, "Average Spend": 455, "Revolving %": 0.43, "Interest %": 0.36 },
    { "Region": "Northeast", "Branch": "New York", "Card Type": "Gold", "Cards Issued": 1900, "Activation %": 0.70, "Average Spend": 315, "Revolving %": 0.34, "Interest %": 0.32 },
    { "Region": "Northeast", "Branch": "New York", "Card Type": "Classic", "Cards Issued": 2700, "Activation %": 0.68, "Average Spend": 400, "Revolving %": 0.29, "Interest %": 0.28 },
    { "Region": "South", "Branch": "Tampa", "Card Type": "Platinum", "Cards Issued": 1000, "Activation %": 0.70, "Average Spend": 440, "Revolving %": 0.34, "Interest %": 0.32 },
    { "Region": "South", "Branch": "Tampa", "Card Type": "Gold", "Cards Issued": 1300, "Activation %": 0.67, "Average Spend": 405, "Revolving %": 0.32, "Interest %": 0.30 },
    { "Region": "South", "Branch": "Tampa", "Card Type": "Classic", "Cards Issued": 1600, "Activation %": 0.66, "Average Spend": 395, "Revolving %": 0.31, "Interest %": 0.29 },
    { "Region": "West", "Branch": "San Jose", "Card Type": "Platinum", "Cards Issued": 2400, "Activation %": 0.82, "Average Spend": 460, "Revolving %": 0.44, "Interest %": 0.38 },
    { "Region": "West", "Branch": "San Jose", "Card Type": "Gold", "Cards Issued": 2100, "Activation %": 0.78, "Average Spend": 330, "Revolving %": 0.36, "Interest %": 0.34 },
    { "Region": "West", "Branch": "San Jose", "Card Type": "Classic", "Cards Issued": 1800, "Activation %": 0.70, "Average Spend": 305, "Revolving %": 0.33, "Interest %": 0.31 },
    { "Region": "Midwest", "Branch": "Chicago", "Card Type": "Platinum", "Cards Issued": 1500, "Activation %": 0.75, "Average Spend": 445, "Revolving %": 0.38, "Interest %": 0.35 },
    { "Region": "Midwest", "Branch": "Chicago", "Card Type": "Gold", "Cards Issued": 2000, "Activation %": 0.70, "Average Spend": 410, "Revolving %": 0.34, "Interest %": 0.33 },
    { "Region": "Midwest", "Branch": "Chicago", "Card Type": "Classic", "Cards Issued": 1700, "Activation %": 0.69, "Average Spend": 400, "Revolving %": 0.30, "Interest %": 0.29 },
    { "Region": "West", "Branch": "Los Angeles", "Card Type": "Platinum", "Cards Issued": 2300, "Activation %": 0.79, "Average Spend": 360, "Revolving %": 0.41, "Interest %": 0.37 },
    { "Region": "West", "Branch": "Los Angeles", "Card Type": "Gold", "Cards Issued": 2500, "Activation %": 0.76, "Average Spend": 420, "Revolving %": 0.35, "Interest %": 0.34 },
    { "Region": "West", "Branch": "Los Angeles", "Card Type": "Classic", "Cards Issued": 1900, "Activation %": 0.70, "Average Spend": 405, "Revolving %": 0.32, "Interest %": 0.30 }
];

export default function App() {
    let spreadsheet;
    const getInterestRate = (branch, cardType) => {
        const base = interestRateMatrix[branch]?.[cardType];
        return base;
    };
    const interestRateMatrix = {
        'Dallas': { Platinum: 0.29, Gold: 0.27, Classic: 0.25 },
        'Austin': { Platinum: 0.30, Gold: 0.28, Classic: 0.26 },
        'San Antonio': { Platinum: 0.29, Gold: 0.29, Classic: 0.27 },
        'Houston': { Platinum: 0.31, Gold: 0.28, Classic: 0.26 },
        'Phoenix': { Platinum: 0.28, Gold: 0.26, Classic: 0.25 },
        'New York': { Platinum: 0.30, Gold: 0.28, Classic: 0.24 },
        'Tampa': { Platinum: 0.27, Gold: 0.25, Classic: 0.24 },
        'San Jose': { Platinum: 0.32, Gold: 0.29, Classic: 0.27 },
        'Chicago': { Platinum: 0.29, Gold: 0.28, Classic: 0.25 },
        'Los Angeles': { Platinum: 0.31, Gold: 0.29, Classic: 0.26 }
    };

    const creditLimitMatrix = {
        'Dallas': { Platinum: 150000, Gold: 100000, Classic: 50000 },
        'Austin': { Platinum: 140000, Gold: 90000, Classic: 45000 },
        'San Antonio': { Platinum: 130000, Gold: 85000, Classic: 40000 },
        'Houston': { Platinum: 125000, Gold: 80000, Classic: 40000 },
        'Phoenix': { Platinum: 160000, Gold: 110000, Classic: 55000 },
        'New York': { Platinum: 180000, Gold: 120000, Classic: 60000 },
        'Tampa': { Platinum: 140000, Gold: 90000, Classic: 45000 },
        'San Jose': { Platinum: 200000, Gold: 140000, Classic: 70000 },
        'Chicago': { Platinum: 150000, Gold: 100000, Classic: 50000 },
        'Los Angeles': { Platinum: 220000, Gold: 150000, Classic: 80000 }
    };
    const handleBeforeCellRender = (args) => {
        if (args.cell && args.cell.template === 'plus-icon') {
            const wrapperDiv = createLowRiskIconWrapper();
            args.element.insertBefore(wrapperDiv, args.element.firstChild);
        }
    };
    const createLowRiskIconWrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-icons e-plus e-custom-icon'; //e-low-risk
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };
    const createMediumRiskIconWrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-icons e-medium-risk e-custom-icon';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };
    const createHighRiskIconWrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-icons e-high-risk e-custom-icon';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };
    const onCreated = () => {
        console.time();
        spreadsheet.updateRange({ dataSource: budgetData, startCell: 'Budget!B4' }, 1);
        spreadsheet.updateCell({ template: 'plus-icon' }, 'Budget!B1');
        BudgetCalculation(spreadsheet);
        spreadsheet.updateRange({ dataSource: actualDataSource(), startCell: 'actuals!B4' }, 2);
        actualsSheetCalculation(spreadsheet);
        varianceSheetCalculation(spreadsheet);
        setTimeout(() => {
            dashboardCalculation(spreadsheet);
            const dashboardSheet = getSheet(spreadsheet, 0);
            setCell(41, 7, dashboardSheet, { value: 'TOP 5 DELIQUENT BRANCHES', colSpan: 3, style: { fontFamily: 'Times New Roman' } });
            setCell(41, 2, dashboardSheet, { value: 'LOW PERFORMING BRANCHES', colSpan: 2, style: { fontFamily: 'Times New Roman' } }); //verticalAlign: 'middle', textAlign: 'center',
            spreadsheet.cellFormat({  fontSize: '11pt', verticalAlign: 'middle', textAlign: 'center' }, 'DashboardInsights!B42:K42'); //fontWeight: 'bold',
            const branchSpend = ['Branch', 'Planned Spend', 'Actual Spend', 'Variance'];
            const deliquentHeaders =[ 'Branch','Deliquent Accounts', 'Days Past Due(Avg)'];
            deliquentHeaders.forEach((val, i) => {
                setCell(42, i + 7, dashboardSheet, { value: val, style:{color:'#124B5C', fontWeight:'bold'} });
            }); 
            branchSpend.forEach((val, i) => {
                setCell(42, i + 1 , dashboardSheet, { value: val, style:{color:'#124B5C', fontWeight:'bold'}  });
            });
            spreadsheet.cellFormat({verticalAlign:'middle'},'DashboardInsights!B43:H43');
            //spreadsheet.cellFormat({fontWeight:'bold'},'DashboardInsights!B43:E43 I43:J43'); //, fontFamily:'Axettac Demo'
            let dashboardDataCell = 28;
            let lowPerformanceBranches = 15; //15
            const dataSheet = getSheet(spreadsheet,4);
            for (let i = 43; i <= 47; i++) {
                setCell(i, 1, dashboardSheet, { formula: `=Dashboard!F${lowPerformanceBranches}` }); //value: `${spreadsheet.sheets[4].rows[dashboardDatacell].cells[9].value}`
                setCell(i, 2, dashboardSheet, { formula: `=Dashboard!G${lowPerformanceBranches}` });
                setCell(i, 3, dashboardSheet, { formula: `=Dashboard!H${lowPerformanceBranches}` });
                setCell(i, 4, dashboardSheet, { formula: `=Dashboard!J${lowPerformanceBranches}` });
                //setCell(i, 5, dashboardSheet, { formula: `=Dashboard!J${lowPerformanceBranches}` });
                setCell(i, 7, dashboardSheet, { formula: `=Dashboard!Q${dashboardDataCell}` });
                setCell(i, 8, dashboardSheet, { formula: `=Dashboard!R${dashboardDataCell}` });
                setCell(i, 9, dashboardSheet, { formula: `=Dashboard!S${dashboardDataCell}` });
                dashboardDataCell++;
                lowPerformanceBranches++;
            }
            //spreadsheet.calculateNow('Dashboard!B43:J48',0);
            spreadsheet.cellFormat( { verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt'},'DashboardInsights!C43:E43 I43:J48')
            spreadsheet.cellFormat( { color:'#ff0000'},'DashboardInsights!I44:J48 E44:E49')//I43:J48
            spreadsheet.cellFormat( { verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt'},'DashboardInsights!C44:E48')
            spreadsheet.setBorder({ border: '1px solid #d4cdcdc8' }, 'DashboardInsights!B41:E49 H41:J49 ','Outer'); //K49
            spreadsheet.setBorder({ border: '1px solid #e6e6e6' },'DashboardInsights!B44:E48 H44:J48','Horizontal');
            spreadsheet.numberFormat('$#,##0.00','DashboardInsights!C44:E48')//spreadsheet.conditionalFormat({ type: 'RedDataBar', range: 'DashboardInsights!J44:J48' });
            spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'DashboardInsights!C44:C48' });
            spreadsheet.conditionalFormat({ type: 'GreenDataBar', range: 'DashboardInsights!D44:D48' });
            //spreadsheet.conditionalFormat({ type: 'LessThan', value:'0',range: 'DashboardInsights!E44:E49', format: { style: { color: '#ff0000' } }});
            setCell(0,0,dashboardSheet,{chart:chart1});
            setCell(1,0,dashboardSheet,{chart:chart3});
            setCell(2,0,dashboardSheet,{chart:chart4});
            setCell(3,0,dashboardSheet,{chart:chart5});
            setCell(4,0,dashboardSheet,{chart:chart6});
            setCell(5,0,dashboardSheet,{chart:chart7});
            setCell(6,0,dashboardSheet,{chart:chart8});
            const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
            setCell(0, 1, dashboardSheet, { value: 'CREDIT CARD EXPENSE DASHBOARD', colSpan: 5, style: headerStyle });
            spreadsheet.resize();
        },100);

    }

    const BudgetCalculation = (spreadsheet) => {
        const sheet = getSheet(spreadsheet, 1);
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        const financeOutcomeSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' }
        const segmentSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' }
        const performanceSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2E8B3C', color: '#fff' }
        const riskSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#B03A2E', color: '#fff' }
        spreadsheet.setColumnsWidth(20, ['Budget!A','actuals!A','variance!A','DashboardInsights!F:G']);
        spreadsheet.setColumnsWidth(135, ['Budget!B:L']);
        spreadsheet.setColumnsWidth(135, ['DashboardInsights!B:E','DashboardInsights!N:O','DashboardInsights!I:J']);
        spreadsheet.setColumnsWidth(130, ['Budget!B:L', 'actuals!B:Q', 'Variance!B:Q']);
        spreadsheet.setColumnsWidth(140,['DashboardInsights!B:E','DashboardInsights!H:J']);
        spreadsheet.setColumnsWidth(1,['DashboardInsights!G']);
        spreadsheet.setColumnsWidth(187,['DashboardInsights!H:J']);
        spreadsheet.setColumnsWidth(80,['DashboardInsights!K']);
        spreadsheet.setRowsHeight(40, ['Budget!3', 'actuals!3', 'Variance!3', 'actuals!B3:I3', 'Dashboard!9', 'Dashboard!4', 'Dashboard!21', 'Dashboard!27']);
        spreadsheet.setRowsHeight(30,['DashboardInsights!42:43']);
        spreadsheet.setRowsHeight(25, ['DashboardInsights!44:49','Budget!4:33', 'actuals!4:10000', 'actuals!B4:I33', 'Variance!4:33', 'Dashboard!5:8', 'Dashboard!10:20', 'Dashboard!22:24', 'Dashboard!28:37']);
        setCell(0, 1, sheet, { value: 'CREDIT CARD EXPENSE SUMMARY - BUDGET', colSpan: 11, rowSpan: 1, style: headerStyle });
        setCell(2, 1, sheet, { value: 'Segment', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
        setCell(2, 4, sheet, { value: 'Volume', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' } });
        setCell(2, 5, sheet, { value: 'Performance', colSpan: 2, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' } });
        setCell(2, 7, sheet, { value: 'Risk', colSpan: 2, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#8B1D18', color: '#fff' } });
        setCell(2, 9, sheet, { value: 'Financial Outcomes', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#5E3A78', color: '#fff' } });
        setCell(3, 9, sheet, { value:'Expected Spend', style: financeOutcomeSubHeaderStyle});
        setCell(3, 10, sheet, { value:'Revolving Balance', style: financeOutcomeSubHeaderStyle});
        setCell(3, 11, sheet, { value:'Expected Interest', style: financeOutcomeSubHeaderStyle});
        setCell(3, 4, sheet, { style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' }});
        for (let i = 1; i <= 3; i++) {
            setCell(3, i, sheet, { style: segmentSubHeaderStyle });
        }
        for (let i = 5; i <= 6; i++) {
            setCell(3, i, sheet, { style: performanceSubHeaderStyle });
        }
         for (let i = 7; i <= 8; i++) {
            setCell(3, i, sheet, { style: riskSubHeaderStyle });
        }
        // spreadsheet.updateCell({ value:'Credit Card Expense Budget'},'Budget!B1');
        // spreadsheet.updateCell({ value: 'Expected Spend' }, `Budget!J3`);
        // spreadsheet.updateCell({ value: 'Revolving Balance' }, `Budget!K3`);
        // spreadsheet.updateCell({ value: 'Expected Interest' }, `Budget!L3`);
        for (let i = 4; i < 34; i++) {
            setCell(i, 9, sheet, { formula: `=E${i + 1}* F${i + 1} * G${i + 1}` });
            setCell(i, 10, sheet, { formula: `=J${i + 1}* H${i + 1}` });
            setCell(i, 11, sheet, { formula: `=K${i + 1}* I${i + 1}` });
        }
        // spreadsheet.numberFormat('$#,##0.00', 'Budget!G1:G33');
        // spreadsheet.numberFormat('$#,##0.00', 'Budget!J1:L33');
        // spreadsheet.merge('Budget!B1:L2');
        // spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign:'center' }, 'Budget!B1:L2 B3:L3');
        // spreadsheet.cellFormat({ backgroundColor:'#00b300', color:'#fff'},'Budget!B1:L2');
        // spreadsheet.cellFormat({ backgroundColor:'#F1FAEE'},'Budget!A1:A100 M1:Z100 B34:L100');
        // for (let i = 4; i < 34; i++) {
        //     spreadsheet.updateCell({ formula: `=E${i}* F${i} * G${i}` }, `Budget!J${i}`);
        //     spreadsheet.updateCell({ formula: `=J${i}* H${i}` }, `Budget!K${i}`);
        //     spreadsheet.updateCell({ formula: `=K${i}* I${i}` }, `Budget!L${i}`);
        // }
        //spreadsheet.setBorder()
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Budget!B3:L4');
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'Budget!B5:L34');
        spreadsheet.numberFormat('$#,##0.00','Budget!G4:G34');
        spreadsheet.numberFormat('$#,##0.00','Budget!J4:L34');
        spreadsheet.numberFormat('0.00%', 'Budget!F4:F34');
        spreadsheet.numberFormat('0.00%', 'Budget!H4:I34');
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center'}, 'Budget!F4:L34');
        spreadsheet.cellFormat({ backgroundColor: '#E6F4EA' }, 'Budget!L5:L34')
        spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'Budget!E4:E34' });
        //spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '70%', format: { style: { color: '#375623', fontWeight:'bold' } }, range: 'Budget!F5:F33'});
        //spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '70%', format: { style: { color: '#375623', fontWeight:'bold' } }, range: 'Budget!F5:F33'});
        //spreadsheet.conditionalFormat({ type: 'Between', value: ['60%','70%'], format: { style: { color: '#375623', fontWeight: 'bold' } }, range: 'Budget!F4:F33' });
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'Budget!H4:H34' });
        spreadsheet.conditionalFormat({ type: 'ThreeTrafficLights1', range: 'Budget!F4:F34' });
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'Budget!H4:H34' });
        spreadsheet.conditionalFormat({ type: 'ThreeTrafficLights1', range: 'Budget!I4:I34' });
        // spreadsheet.protectSheet(1, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    }

    const actualsSheetCalculation = (spreadsheet) => {
        //spreadsheet.cellFormat({ backgroundColor:'#F1FAEE'},'actuals!A1:A100 R1:Z1000');
        const sheet = getSheet(spreadsheet, 2);
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        const financeOutcomeSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' }
        const segmentSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' }
        const volumeSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' }
        const riskSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#B03A2E', color: '#fff' }
        spreadsheet.setColumnsWidth(20, ['Budget!A','actuals!A','variance!A']);
        //overall actuals formatting
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center' }, 'actuals!B4:Q1000');
        spreadsheet.cellFormat({ fontWeight: 'bold' }, 'actuals!G4:H1000');
        spreadsheet.cellFormat({ textAlign:'right'}, 'actuals!K4:K1000 O4:O1000');
        setCell(0, 1, sheet, { value: 'CREDIT CARD EXPENSE SUMMARY - ACTUALS', colSpan: 16, rowSpan: 1, style: headerStyle });
        for (let i = 1; i <= 5; i++) {
            setCell(3, i, sheet, { style: financeOutcomeSubHeaderStyle });
        }
        for (let i = 6; i <= 8; i++) {
            setCell(3, i, sheet, { style: volumeSubHeaderStyle });
        }
        for (let i = 9; i <= 12; i++) {
            setCell(3, i, sheet, { style: segmentSubHeaderStyle });
        }
        for (let i = 13; i <= 16; i++) {
            setCell(3, i, sheet, { style: riskSubHeaderStyle });
        }
        //spreadsheet.updateCell({ value:'Credit Card Expense Summary'},'actuals!B1');
        //spreadsheet.merge('actuals!B1:Q2');
        //spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign: 'center', backgroundColor:'#00b300', color:'#fff' }, 'actuals!B1:N2 B3:Q3');
        //spreadsheet.cellFormat({ textAlign: 'left' }, 'actuals!B4:B1000');
        //bold
        //spreadsheet.cellFormat({ fontWeight: 'bold' }, 'actuals!H4:H1000 K4:K1000 J4:J1000');
        //minimum due
        //spreadsheet.cellFormat({ color: '#2563EB' }, 'actuals!J4:J1000');
        //payment made
        //spreadsheet.cellFormat({ color: '#0D9488' }, 'actuals!K4:K1000');
        setCell(2, 1, sheet, { value: 'Segment', colSpan: 5, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
        setCell(2, 6, sheet, { value: 'Credit and Usage', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' } });
        setCell(2, 9, sheet, { value: 'Payment and Behaviour ', colSpan: 4,style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' } });
        setCell(2, 13, sheet, { value: 'Risk and Deliquency', colSpan: 4, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#8B1D18', color: '#fff' } });
        spreadsheet.numberFormat('$#,##0.00', 'actuals!F1:H1000');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!I1:K1000');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!O1:O1000');
        //spreadsheet.updateRange({ dataSource: actualDataSource(), startCell: 'actuals!B3' }, 2);
        //spreadsheet.addDataValidation({ type: 'List', value1: 'Yes,No', ignoreBlank: false }, 'actuals!N4:N1000');
        spreadsheet.addDataValidation({ type: 'List', value1: 'Full,Partial,Missed', ignoreBlank: false }, 'actuals!L4:L1000');
        //spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!H4:H1000', cFColor: 'RedT' });
        //spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!H4:H1000', format: { style: { color: 'Green' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'High', range: 'actuals!Q4:Q1000', cFColor: 'RedFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Low', range: 'actuals!Q4:Q1000', cFColor: 'GreenFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Medium', range: 'actuals!Q4:Q1000', cFColor: 'YellowFT' });
        //deliquency 
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!P4:P1000', format: { style: { backgroundColor:'#f2f2f2', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!P4:P1000', cFColor: 'RedFT' });
        //Balance carried
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!M4:M1000', format:{ style: {color: '#468966'}} });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!M4:M1000', format:{ style: {color: '#000000'}}});
        //Payment made
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Missed', range: 'actuals!L4:L1000', format: { style: { color: '#DA1212', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Full', range: 'actuals!L4:L1000', format: { style: { color: '#468966', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Partial', range: 'actuals!L4:L1000', format: { style: { color: '#230338', fontWeight: 'bold' } } });
        //Days past Due
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'actuals!N4:N1000' });
        spreadsheet.conditionalFormat({ type: 'GreenDataBar', range: 'actuals!O4:O1000' });
        spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'actuals!K4:K1000' });
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'actuals!B3:Q4');
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'actuals!B5:Q1000');
        spreadsheet.protectSheet(2, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    };

    const dashboardCalculation = (spreadsheet) => { 
        const sheet = getSheet(spreadsheet, 4);
        //setCell(0, 1, sheet, { value: 'Dashboard'});
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign: 'center' }, 'Dashboard!A2:E2 A4:E4 A9:E9 A21:E21 A27:E27');
        const metrics = ['Total Spend', 'Total Interest'];
        //const locations = ['Bengaluru', 'Chennai', 'Coimbatore', 'Delhi', 'Kochi', 'Kolkata', 'Madurai', 'Mumbai', 'Pune', 'Trichy'];
        const locations =['Dallas','Austin','San Antonio','Houston','Tampa','San Jose','Phoenix','New York','Chicago','Los Angeles']
        const cardTypes = ['Gold', 'Classic', 'Platinum'];
        let i = 5;
        metrics.forEach((val) => {
            //spreadsheet.updateCell({ value: val }, `Dashboard!A${i}`);
            setCell(i, 1, sheet, { value: val });
            i++;
        });
        spreadsheet.merge('Dashboard!A4:B4');
        spreadsheet.merge('Dashboard!D4:E4');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Budget!B3:L3');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Actuals!B3:P3');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Variance!B3:L3');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Budget!B3:L33', 'Vertical');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Actuals!B3:P1000', 'Vertical');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Variance!B3:L33', 'Vertical');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Dashboard!A4:B4 D4:E4 A9:E9 A21:E21 A27:B27 D27:E27', 'Outer');
        // spreadsheet.setBorder({ border: '1px solid #000000' }, 'Dashboard!A4:B4 D4:E4 A9:E9 A21:E21 A27:B27 D27:E27', 'Vertical');
        setCell(5, 2, sheet, { formula: '=SUM(Variance!E4:E33)' });
        setCell(6, 2, sheet, { formula: '=SUM(Variance!I4:I33)' });
        //spreadsheet.updateCell({ formula: '=SUM(Variance!E4:E33)' }, 'Dashboard!B5');
        //spreadsheet.updateCell({ formula: '=SUM(Variance!I4:I33)' }, 'Dashboard!B6');
        const branchSpend = ['Branch', 'Planned Spend', 'Actual Spend', 'Actual Interest', 'Variance'];
        const cardSpend = ['Card Type', 'Actual Spend', 'Actual Interest', 'Variance', 'Planned Spend'];
        const cardRisk = ['Low', 'Medium', 'High'];
        const deliquentCalc = ['Branch', 'Payment Risk'];
        const topBranch = ['Branch', 'Actual Spend', 'Actual Interest'];
        const deliquentCalcCard = ['Card', 'Payment Risk'];
        setCell(20, 6, sheet, { value: 'Branch' });
        setCell(20, 7, sheet, { value: 'Paid in Full' });
        setCell(20, 8, sheet, { value: 'Partial Payment' });
        setCell(20, 9, sheet, { value: 'Missed Payment' });
        //doughnut chart
        setCell(37, 12, sheet, { value: 'Paid in Full' });
        setCell(37, 13, sheet, { value: 'Partial Payment' });
        setCell(37, 14, sheet, { value: 'Missed Payment' });
        setCell(38, 12, sheet, { formula: '=SUM(Dashboard!M28:M37)' });
        setCell(38, 13, sheet, { formula: '=SUM(Dashboard!N28:N37)' });
        setCell(38, 14, sheet, { formula: '=SUM(Dashboard!O28:O37)' });

        setCell(26, 16, sheet, { value: 'Branch' });
        setCell(26, 17, sheet, { value: 'Deliquent Accounts' });
        setCell(26, 18, sheet, { value: 'Days Past Due(Avg)' });
        topBranch.forEach((val, i) => {
            //setCell(26, 6 + i, sheet, { value: val });
            const col = String.fromCharCode(71 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}27`);
        });
        branchSpend.forEach((val, i) => {
            const col = String.fromCharCode(65 + i);
            const dataCol = String.fromCharCode(70 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${dataCol}9`);
            //setCell(5, 2, sheet, { formula: '=SUM(Variance!E4:E33)' });
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
            spreadsheet.updateCell({ value: val }, `Dashboard!F${j}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!A${newLocation}`);
            //new data 
            spreadsheet.updateCell({ value: val }, `Dashboard!G${newLocation}`);
            spreadsheet.updateCell({ value: val},`Dashboard!J${newLocation}`);
            spreadsheet.updateCell({ value: val},`Dashboard!L${newLocation}`);
            spreadsheet.updateCell({ value: val},`Dashboard!Q${newLocation}`);
            //payment calculations
            // spreadsheet.updateCell({ formula:`=AVERAGEIFS(Actuals!Q4:Q1004,Actuals!D4:D1004,A${newLocation},Actuals!L4:L1004,"Full")` },`Dashboard!M${newLocation}`);
            // spreadsheet.updateCell({ formula:`=AVERAGEIFS(Actuals!Q4:Q1004,Actuals!D4:D1004,A${newLocation},Actuals!L4:L1004,"Partial")` },`Dashboard!N${newLocation}`);
            // spreadsheet.updateCell({ formula:`=AVERAGEIFS(Actuals!Q4:Q1004,Actuals!D4:D1004,A${newLocation},Actuals!L4:L1004,"Missed")` },`Dashboard!O${newLocation}`);
            spreadsheet.updateCell({ formula:`=COUNTIFS(Actuals!D4:D1004,Dashboard!L${newLocation},Actuals!L4:L1004,"Full")` },`Dashboard!M${newLocation}`);
            spreadsheet.updateCell({ formula:`=COUNTIFS(Actuals!D4:D1004,Dashboard!L${newLocation},Actuals!L4:L1004,"Partial")` },`Dashboard!N${newLocation}`);
            spreadsheet.updateCell({ formula:`=COUNTIFS(Actuals!D4:D1004,Dashboard!L${newLocation},Actuals!L4:L1004,"Missed")` },`Dashboard!O${newLocation}`);
            spreadsheet.updateCell({ formula: `=AVERAGEIF(Variance!C8:C37,A${newLocation},Variance!K8:K37)` }, `Dashboard!B${newLocation}`);
            //deliquent account counts
            spreadsheet.updateCell({ formula: `=COUNTIFS(Actuals!D4:D1004,A${newLocation},Actuals!P4:P1004,"Yes")` }, `Dashboard!R${newLocation}`);
            //days past due
            spreadsheet.updateCell({ formula: `=ROUND(AVERAGEIFS(Actuals!N4:N1004,Actuals!D4:D1004,A${newLocation},Actuals!P4:P1004,"Yes"),0)` }, `Dashboard!S${newLocation}`);
            //NEW 
            spreadsheet.updateCell({ formula: `=AVERAGEIF(Variance!C8:C37,Dashboard!J${newLocation},Variance!K8:K37` }, `Dashboard!K${newLocation}`);
            //top 10
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${newLocation},Variance!F8:F37` }, `Dashboard!H${newLocation}`);
            //average for deliquency scores     
            newLocation++;
            j++;
        });
        for (let k = 10; k < 20; k++) {
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k},Variance!F8:F37` }, `Dashboard!C${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k},Variance!J8:J37` }, `Dashboard!D${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k},Variance!G8:G37` }, `Dashboard!E${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Budget!C8:C37,Dashboard!A${k},Budget!J8:J37` }, `Dashboard!B${k}`);
            //
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k},Variance!F8:F37` }, `Dashboard!H${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k},Variance!J8:J37` }, `Dashboard!I${k}`);
            //chart
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k},Variance!J8:J37` }, `Dashboard!I${k + 18}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k},Variance!G8:G37` }, `Dashboard!J${k}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Budget!C8:C37,Dashboard!A${k},Budget!J8:J37` }, `Dashboard!G${k}`);
        }
        let cardTrackNumber = 22;
        cardTypes.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!A${cardTrackNumber}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!G${cardTrackNumber}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!D${cardTrackNumber + 6}`);
            spreadsheet.updateCell({ formula: `=AVERAGEIF(Variance!D8:D37,D${cardTrackNumber + 6} , Variance!K8:K37)` }, `Dashboard!E${cardTrackNumber + 6}`);
            cardTrackNumber++;
        })
        for (let i = 22; i < 25; i++) {
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!D8:D37,Dashboard!A${i},Variance!F8:F37` }, `Dashboard!B${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!D8:D37,Dashboard!A${i},Variance!J8:J37` }, `Dashboard!C${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!D8:D37,Dashboard!A${i},Variance!G8:G37` }, `Dashboard!D${i}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Budget!D8:D37,Dashboard!A${i},Budget!J8:J37` }, `Dashboard!E${i}`);

            //bar chart formulas
            spreadsheet.updateCell({ formula:`=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Full")/1000` },`Dashboard!H${i}`);
            spreadsheet.updateCell({ formula:`=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Partial")/1000` },`Dashboard!I${i}`);
            spreadsheet.updateCell({ formula:`=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Missed")/1000` },`Dashboard!J${i}`);
            //spreadsheet.updateCell({ formula:`=AVERAGEIFS(Actuals!K4:K1004,Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Full")/AVERAGEIFS(Actuals!H4:H1004,Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Full")` },`Dashboard!H${i}`);
            //spreadsheet.updateCell({ formula:`=AVERAGEIFS(Actuals!K4:K1004,Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Partial")/AVERAGEIFS(Actuals!H4:H1004,Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Partial")` },`Dashboard!I${i}`);
            //spreadsheet.updateCell({ formula:`=AVERAGEIFS(Actuals!K4:K1004,Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Missed")/AVERAGEIFS(Actuals!H4:H1004,Actuals!E4:E1004,Dashboard!G${i},Actuals!L4:L1004,"Missed")` },`Dashboard!J${i}`);
        }
        deliquentCalc.forEach((val, i) => {
            const col = String.fromCharCode(65 + i);
            const updatedCol = String.fromCharCode(74 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${updatedCol}27`);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}27`);
        })
        deliquentCalcCard.forEach((val, i) => {
            const col = String.fromCharCode(68 + i);
            spreadsheet.updateCell({ value: val }, `Dashboard!${col}27`);
        })
        spreadsheet.updateCell({ value: 'Key Metrics' }, 'Dashboard!A4');
        spreadsheet.updateCell({ value: 'Credit Card Risk State' }, 'Dashboard!D4');
        //for( let i = 28 ; i < );
        // spreadsheet.copy('Dashboard!J28:K32');
        // setTimeout(() => {    
        // spreadsheet.paste('Dashboard!J39','values');
        // }, 100);
        let cardRiskNumber = 5;
        cardRisk.forEach((val) => {
            spreadsheet.updateCell({ value: val }, `Dashboard!D${cardRiskNumber}`);
            spreadsheet.updateCell({ formula: `=COUNTIF(Variance!M8:M37,Dashboard!D${cardRiskNumber})` }, `Dashboard!E${cardRiskNumber}`);
            cardRiskNumber++;
        });
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B10:E24');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!H28:H37');
        spreadsheet.numberFormat('0.00%', 'Dashboard!K28:K37');
        spreadsheet.numberFormat('0.00%', 'Dashboard!H22:J25');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B4:B6 ');
        spreadsheet.numberFormat('0.00%', 'Dashboard!B28:B37');
        spreadsheet.numberFormat('0.00%', 'Dashboard!E28:E30');
        // spreadsheet.insertChart(chart1);
        // //spreadsheet.insertChart(chart2);
        // spreadsheet.insertChart(chart3);
        // spreadsheet.insertChart(chart4);
        // spreadsheet.insertChart(chart5);
        // spreadsheet.insertChart(chart6);
        // spreadsheet.insertChart(chart7);
        // spreadsheet.insertChart(chart8);
        //spreadsheet.cellFormat({color:'#fff'},'Dashboard!A3:F50');
        // setTimeout(() => {
            spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'H', order: 'Ascending' } }, 'Dashboard!G27:H37');
            // spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'K', order: 'Descending' } }, 'Dashboard!J27:K37');
            spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'S', order: 'Descending' } }, 'Dashboard!Q27:S37');
            spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'J', order: 'Descending' } }, 'Dashboard!F9:J19');
        //});
        //spreadsheet.calculateNow(4);
        //spreadsheet.updateCell({fontWeight:'bold'},'Dashboard!F9:K37');
        //spreadsheet.resize();
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'Variance!B8:M37');
        //spreadsheet.protectSheet(0, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    }
    //2
    const chart3 = [{ type: 'Doughnut', range: 'Dashboard!B6:C7', title: 'TOTAL SPEND VS TOTAL INTEREST', theme: 'Tailwind3', height: 345, left: 55, top: 60, width: 570, id: 'Chart3', legendSettings:{position:'Right'} }];
    const chart8 = [{ type: 'Bar', range: 'Dashboard!G27:H37', title: 'TOP 10 BRANCHES BY SPEND', theme: 'Tailwind3', left: 1265, top: 60, width: 570, height: 345, id: 'Chart8', isSeriesInRows: false}];
    const chart1 = [{ type: 'Column', range: 'Dashboard!G27:I37', title: 'BRANCH WISE SPEND AND INTEREST', theme: 'Tailwind3', left: 640, top: 60, width: 570, height: 345, id: 'Chart1', isSeriesInRows: false }];
    const chart6 = [{ type: 'Doughnut', range: 'Dashboard!M38:O39', title: 'PAYMENT STATUS SUMMARY', theme: 'Tailwind3', height: 345, left: 55, top: 415, width: 570, id: 'Chart6',legendSettings:{position:'Right'}, isSeriesInRows:true, dataLabelSettings:{position:'middle', visible:'true'} }];
    //const chart2 = [{ type: 'StackingLine', range: 'Dashboard!A21:C24', title: 'CARD TYPE SUMMARY', theme: 'Tailwind3', height: 280, left: 608, top: 701, width: 585, id: 'Chart2', isSeriesInRows: true }];
    //1
    //6
    const chart4 = [{ type: 'Pie', range: 'Dashboard!A27:B37', title: 'PAYMENT RISK BASED ON LOCATION', theme: 'Tailwind3', height: 345, left: 1265, width: 570, top: 415, id: 'Chart4',legendSettings:{position:'Right'}, dataLabelSettings:{position:'middle', visible:'true'} }];
    //5
    const chart5 = [{ type: 'StackingBar', range: 'Dashboard!G21:J24', title: 'PAYMENT RISK BASED ON CARD', theme: 'Tailwind3', top: 415, width: 570, left: 640, height: 345, id: 'Chart5',legendSettings:{position:'Top'},dataLabelSettings:{position:'middle', visible:'true' } }]; //
    const chart7 = [{ type: 'Column', range: 'Dashboard!A21:A24 C21:C24', title: 'INTEREST YIELD BY CARD TYPE', theme: 'Tailwind3', height: 240, left: 1265, top: 810, width: 570, id: 'Chart7', dataLabelSettings:{position:'middle', visible:'true' }}];
    const getCreditLimit = (branch, cardType) => {
        const base = creditLimitMatrix[branch]?.[cardType];
        return base;
    };

    const actualDataSource = () => {
    // const branches = [
    //     'Chennai', 'Coimbatore', 'Madurai', 'Trichy',
    //     'Pune', 'Delhi', 'Kochi', 'Bengaluru',
    //     'Kolkata', 'Mumbai'
    // ];
        const regionBranches = {
            South: ['Dallas', 'Austin', 'San Antonio', 'Houston', 'Tampa'],
            West: ['Phoenix', 'San Jose', 'Los Angeles'],
            Northeast: ['New York'],
            Midwest: ['Chicago']
        };
    const branches =['Dallas','Austin','San Antonio','Houston','Tampa','San Jose','Phoenix','New York','Chicago','Los Angeles']
    const cardTypes = ['Platinum', 'Gold', 'Classic'];
    const startDate = new Date(2025, 0, 1).getTime();
    const endDate = new Date(2025, 11, 31).getTime();
    const dateRange = endDate - startDate;
    const actualsData = [];
    for (let i = 0; i < 1000; i++) {
        // 1. Pick a random region
        const regions = Object.keys(regionBranches);
        const region = regions[Math.floor(Math.random() * regions.length)];
        // 2. Pick a random branch ONLY from that region
        const branches = regionBranches[region];
        const randomDate = new Date(startDate + Math.random() * dateRange);
        const branch = branches[Math.floor(Math.random() * branches.length)];
        const cardType = cardTypes[Math.floor(Math.random() * cardTypes.length)];
        const cardId = `CC${Math.floor(Math.random() * 9000000 + 1000000)}`;
        const creditLimit = getCreditLimit(branch,cardType);
        const statementBalance = Math.floor(Math.random() * 50000);
        // 5–10% minimum due
        const minimumDue = +(statementBalance * (0.05 + Math.random() * 0.05)).toFixed(2);
        // ----------------------------
        // Payment Made (static input)
        // ----------------------------
        let paymentMade = 0;
        const r = Math.random();
        if (statementBalance > 0) {
            if (r < 0.3) {
                paymentMade = statementBalance;          // Full
            } else if (r < 0.7) {
                paymentMade = Math.floor(statementBalance * 0.4); // Partial
            } else {
                paymentMade = 0;                         // Missed
            }
        }

        // ----------------------------
        // Payment Status (Derived)
        // ----------------------------
        let paymentStatus = 'Missed';
        if (paymentMade >= statementBalance && statementBalance > 0) {
            paymentStatus = 'Full';
        } else if (paymentMade > 0) {
            paymentStatus = 'Partial';
        }

        // ----------------------------
        // Balance Carried (Derived)
        // ----------------------------
        const balanceCarried =
            paymentMade < statementBalance ? 'Yes' : 'No';

        // ----------------------------
        // Days Past Due (Static / Simulated)
        // ----------------------------
        let daysPastDue = 0;
        if (paymentStatus === 'Missed' && Math.random() < 0.25) {
            daysPastDue = Math.floor(Math.random() * 30) + 1; // 1–30 days
        }
        else if(paymentStatus === 'Missed'){
            daysPastDue = Math.floor(Math.random() * 30) + 1;
        }

        // ----------------------------
        // Interest (Derived)
        // ----------------------------
        const interestRate = getInterestRate(branch, cardType);
        const interestAmount =
            balanceCarried === 'Yes'
                ? +((statementBalance - paymentMade) * interestRate).toFixed(2)
                : 0;
        // ----------------------------
        // Is Delinquent (Derived)
        // ----------------------------
        const isDelinquent =
            daysPastDue > 0 ? 'Yes' : 'No';
        // ----------------------------
        // Risk Flag (Derived)
        // ----------------------------
        let riskFlag = 'Low';
        if (daysPastDue >= 30) {
            riskFlag = 'High';
        } else if (daysPastDue > 0) {
            riskFlag = 'Medium';
        }
        actualsData.push({
            'Date': randomDate.toLocaleDateString('en-IN'),
            'Region': region,
            'Branch': branch,
            'Card Type': cardType,
            'Customer ID': cardId,
            'Credit Limit': creditLimit,
            'Spend': statementBalance,
            'Statement Balance': statementBalance,
            'Minimum Due': minimumDue,
            'Payment Made': paymentMade,
            'Payment Status': paymentStatus,
            'Balance Carried': balanceCarried,
            'Days Past Due': daysPastDue,        // static demo data
            'Interest': interestAmount,
            'Is Delinquent': isDelinquent,
            'Risk Flag': riskFlag
        });
    }
    return actualsData;
};

const varianceSheetCalculation = (spreadsheet) => {
    const sheet = getSheet(spreadsheet,3);    
    //const riskSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' }
    const segmentSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' }
    const volumeSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' }
    const riskSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' }
    const performanceSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' }
    const interestImpactSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#C77700', color: '#fff' }
    //const interestImpactSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#C77700', color: '#fff' }
    const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
    setCell(0, 1, sheet, { value: 'CREDIT CARD EXPENSE SUMMARY - VARIANCE ANALYSIS', colSpan: 12, rowSpan: 1, style: headerStyle });
    for (let i = 1; i <= 3; i++) {
        setCell(6, i, sheet, { style: segmentSubHeaderStyle });
    }
    for (let i = 4; i <= 7; i++) {
        setCell(6, i, sheet, { style: performanceSubHeaderStyle });
    }
    for (let i = 8; i <= 9; i++) {
        setCell(6, i, sheet, { style: interestImpactSubHeaderStyle });
    }
    for (let i = 10; i <= 11; i++) {
        setCell(6, i, sheet, { style: riskSubHeaderStyle });
    }
    setCell(6, 12, sheet, { style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' } });
        //spreadsheet.cellFormat({ backgroundColor:'#F1FAEE'},'Variance!A1:A100 N1:Z100 B34:L100');
        //spreadsheet.updateCell({ value:'Credit Card Expense Variance Analysis'},'Variance!B1');
        //spreadsheet.merge('Variance!B1:M2');
        //spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign:'center', backgroundColor:'#00b300', color:'#fff' }, 'Variance!B1:L2 B3:M3');
        const headers = ['Region','Branch', 'Card Type', 'Expected Spend', 'Actual Spend', 'Variance', 'Variance %', 'Expected Interest', 'Actual Interest', 'Payment Risk %', 'Risk Score', 'Risk Level']; //'Run_Rate',
        const locations =['Dallas','Austin','San Antonio','Houston','Tampa','San Jose','Phoenix','New York','Chicago','Los Angeles']
        setCell(5, 1, sheet, { value: 'Segment', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
        setCell(5, 4, sheet, { value: 'Performance', colSpan: 4, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' } });
        setCell(5, 8, sheet, { value: 'Interest Impact ', colSpan: 2,style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' } });
        setCell(5, 10, sheet, { value: 'Risk Score', colSpan: 2, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' } });
        setCell(5, 12, sheet, { value: 'Action Priority', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' } });
        headers.forEach((val, i) => {
            const col = String.fromCharCode(66 + i);
            spreadsheet.updateCell({ value: val }, `variance!${col}7`);
        });
    //image headers
    setCell(2, 1, sheet, { value: 'Total Expected Spend', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px' }, colSpan:2});
    setCell(2, 4, sheet, { value: 'Total Actual Spend', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px'}, colSpan:2 });
    //setCell(2, 7, sheet, { value: 'Total Variance', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt'}, colSpan:2});
    setCell(2, 7, sheet, { value: 'Total Expected Interest', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px'}, colSpan:2 });
    setCell(2, 10, sheet, { value: 'Total Actual Interest', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px'}, colSpan:2 });
    //setCell(2, 12, sheet, { value: 'Average Risk Score', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt' } });
    //formulas for header values
    setCell(3, 1, sheet, { formula: '=SUM(E8:E37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px' }, colSpan:2 });
    setCell(3, 4, sheet, { formula: '=SUM(F8:F37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px'}, colSpan:2 });
    //setCell(3, 7, sheet, { formula: '=SUM(G8:G37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt'}, colSpan:2 });
    setCell(3, 7, sheet, { formula: '=SUM(I8:I37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px'}, colSpan:2 });
    setCell(3, 10, sheet, { formula: '=SUM(J8:J37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt', textIndent:'30px'}, colSpan:2 });
    //setCell(3, 12, sheet, { formula: '=AVERAGE(L8:L37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt' } });
        const cardTypes = ['Gold', 'Classic', 'Platinum'];
        let row = 37;
        // const sheet = getSheet(spreadsheet,3);
        for( let i = 7; i < row; i++){
            // spreadsheet.updateCell({formula: `=Budget!B${i}`},`Variance!B${i}`);
            // spreadsheet.updateCell({formula: `=Budget!C${i}`},`Variance!C${i}`);
            // spreadsheet.updateCell({formula: `=Budget!D${i}`},`Variance!D${i}`);
            setCell(i, 1, sheet, { formula: `=Budget!B${i - 2}` });
            setCell(i, 2, sheet, { formula: `=Budget!C${i - 2}` });
            setCell(i, 3, sheet, { formula: `=Budget!D${i - 2}` });
        }
        for (let i = 7; i < row; i++) {
            setCell(i, 4, sheet, { formula: `=SUMIFS(Budget!J4:J34, Budget!C4:C34, Variance!C${i+1}, Budget!D4:D34, Variance!D${i+1})` });
            setCell(i, 5, sheet, { formula: `=SUMIFS(actuals!H4:H1000, actuals!D4:D1000, Variance!C${i+1}, actuals!E4:E1000, Variance!D${i+1})` });
            setCell(i, 6, sheet, { formula: `=Variance!F${i+1} - Variance!E${i+1}`, style: {fontWeight:'bold'} });
            setCell(i, 7, sheet, { formula: `=Variance!G${i+1} / Variance!E${i+1}`, style: {fontWeight:'bold'} });
            setCell(i, 8, sheet, { formula: `=SUMIFS(Budget!L4:L34, Budget!C4:C34, Variance!C${i+1}, Budget!D4:D34, Variance!D${i+1})` });
            setCell(i, 9, sheet, { formula: `=SUMIFS(actuals!O4:O1000, actuals!D4:D1000, Variance!C${i+1}, actuals!E4:E1000, Variance!D${i+1})` });
            setCell(i, 10, sheet, { formula: `=COUNTIFS(actuals!D4:D1000, Variance!C${i+1},actuals!E4:E1000, Variance!D${i+1},actuals!N4:N1000, ">0") / COUNTIFS(actuals!D4:D1000, Variance!C${i+1},
            actuals!E4:E1000, Variance!D${i + 1})` });
            setCell(i, 11, sheet, { formula: `=(Budget!H${i + 1} * 0.4) + (Variance!K${i + 1} * 0.6)` });
            setCell(i, 12, sheet, { formula: `=IF(Variance!L${i + 1} > 0.30, "High", IF( Variance!L${i + 1} > 0.20, "Medium","Low"))` });
            
            //
            // spreadsheet.updateCell({ formula: `=SUMIFS(Budget!J4:J33, Budget!C4:C33, Variance!C${i}, Budget!D4:D33, Variance!D${i})` }, `Variance!E${i}`);
            // spreadsheet.updateCell({ formula: `=SUMIFS(actuals!H4:H1000, actuals!D4:D1000, Variance!C${i}, actuals!E4:E1000, Variance!D${i})` }, `Variance!F${i}`);
            // spreadsheet.updateCell({ formula: `=Variance!E${i} - Variance!F${i}` }, `Variance!G${i}`);
            // spreadsheet.updateCell({ formula: `=Variance!G${i} / Variance!E${i}` }, `Variance!H${i}`);
            // spreadsheet.updateCell({ formula: `=SUMIFS(Budget!L4:L33, Budget!C4:C33, Variance!C${i}, Budget!D4:D33, Variance!D${i})` }, `Variance!I${i}`);
            // spreadsheet.updateCell({ formula: `=SUMIFS(actuals!O4:O1000, actuals!D4:D1000, Variance!C${i}, actuals!E4:E1000, Variance!D${i})` }, `Variance!J${i}`);
            // spreadsheet.updateCell({ formula: `=COUNTIFS(actuals!D4:D1000, Variance!C${i},actuals!E4:E1000, Variance!D${i},actuals!N4:N1000, ">0") / COUNTIFS(actuals!D4:D1000, Variance!C${i},
            // actuals!E4:E1000, Variance!D${i})`}, `Variance!K${i}`);
            // spreadsheet.updateCell({ formula: `=(Budget!H${i} * 0.4) + (Variance!K${i} * 0.6)`}, `Variance!L${i}`);
            // spreadsheet.updateCell({ formula: `=IF(Variance!L${i} > 0.30, "High", IF( Variance!L${i} > 0.20, "Medium","Low"))` }, `Variance!M${i}`);
            /*
            // spreadsheet.updateCell({ formula: `=SUMIFS(actuals!H4:H1000,actuals!C4:C1000,Variance!B${i}, actuals!D4:D1000 , Variance!C${i},actuals!L4:L1000,"Yes")/SUMIFS(actuals!H4:H1000,actuals!C4:C1000,Variance!B${i}, actuals!D4:D1000 , Variance!C${i})` }, `Variance!J${i}`);
            // spreadsheet.updateCell({ formula: `=(Budget!H${i}*0.4)+(Variance!J${i}*0.6)` }, `Variance!K${i}`);
            // spreadsheet.updateCell({ formula: `=IF(Variance!K${i}>0.30,"High Risk",IF(Variance!K${i}>0.20,"Moderate Risk","Low Risk"))` }, `Variance!L${i}`);
            */
        }
        spreadsheet.numberFormat('$#,##0.00', 'Variance!E4:G37');
        spreadsheet.numberFormat('$#,##0.00', 'Variance!I4:J37');
        spreadsheet.numberFormat('0.00%', 'Variance!K4:K37');
        spreadsheet.numberFormat('0.00%', 'Variance!H4:H37');
        spreadsheet.numberFormat('0.00%', 'Variance!L4:L37');
        //spreadsheet.conditionalFormat({ type: 'ThreeSymbols2', range: 'Variance!K4:K33' });
        //spreadsheet.conditionalFormat({ type: 'ThreeTriangles', range: 'Variance!L4:L37' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Low', range: 'Variance!M4:M37',  format:{style:{ color: '#375623', fontWeight:'bold' }} });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Medium', range: 'Variance!M4:M37',  format:{style:{ color: '#c7c705', fontWeight:'bold' }}});
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'High', range: 'Variance!M4:M37',  format:{style:{ color: '#ff3333', fontWeight:'bold' }} });
        //spreadsheet.conditionalFormat({ type: 'BWRColorScale', range: 'Variance!G4:G33' });
        spreadsheet.conditionalFormat({ type: 'RedDataBar', range: 'Variance!K8:K37' });
        spreadsheet.conditionalFormat({ type: 'GreaterThan', value:'0',range: 'Variance!G8:G37', format:{style:{ color: '#375623', fontWeight:'bold' }} });
        spreadsheet.conditionalFormat({ type: 'LessThan', value:'0',range: 'Variance!G8:G37', format:{style:{ color: '#ff3333', fontWeight:'bold' }} });
        spreadsheet.conditionalFormat({ type: 'AboveAverage',range: 'Variance!L8:L37', format:{style:{ color: '#ff3333', fontWeight:'bold' }} });
        spreadsheet.conditionalFormat({ type: 'BelowAverage', range: 'Variance!L8:L37', format:{style:{ color: '#375623', fontWeight:'bold' }} });
        // spreadsheet.conditionalFormat({ type: 'AboveAverage',range: 'Variance!H4:H37', format:{style:{ color: '#375623', fontWeight:'bold' }} });
        // spreadsheet.conditionalFormat({ type: 'BelowAverage',range: 'Variance!H4:H37', format:{style:{ color: '#ff3333', fontWeight:'bold' }} });
        spreadsheet.conditionalFormat({ type: 'GreaterThan', value:'0',range: 'Variance!G8:G37', format:{style:{ color: '#375623', fontWeight:'bold' }} });
        spreadsheet.conditionalFormat({ type: 'LessThan', value:'0',range: 'Variance!G8:G37', format:{style:{ color: '#ff3333', fontWeight:'bold' }} });
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Variance!B6:M7');
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Variance!B3:C4 E3:F4 H3:I4 K3:L4','Outer'); //B1:M1
        setCell(0, 0, sheet, { image: budgetImage });
        //setCell(0, 2, sheet, { image: incomeGraph });
        setCell(0, 4, sheet, { image: interestImage });
        setCell(0, 6, sheet, { image: varianceImage });
        setCell(0, 8, sheet, { image: moneyImage });
        //setCell(0, 9, sheet, { image: riskImage });
        spreadsheet.protectSheet(3, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
        console.timeEnd();
    };

    return (
        <SpreadsheetComponent height={950} ref={(ssObj) => { spreadsheet = ssObj }} created={onCreated.bind(this)} beforeCellRender={handleBeforeCellRender}
        openUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/open'
        saveUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/save'
        >
            <SheetsDirective>
                <SheetDirective name='DashboardInsights' showGridLines={false}></SheetDirective>
                <SheetDirective name='Budget' showGridLines={false}></SheetDirective>
                <SheetDirective name='Actuals' showGridLines={false}></SheetDirective>
                <SheetDirective name='Variance' showGridLines={false}></SheetDirective>
                <SheetDirective name='Dashboard' showGridLines={false} state='Hidden'></SheetDirective>
            </SheetsDirective>
        </SpreadsheetComponent>
        
    );
}