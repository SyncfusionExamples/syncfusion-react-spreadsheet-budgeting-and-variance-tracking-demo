import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { CellDirective, CellsDirective, ColumnDirective, ColumnsDirective, getColumnHeaderText, getSheet, RangeDirective, RangesDirective, RowDirective, RowsDirective, setCell, setColumn, SheetDirective, SheetsDirective, sheetTabs, SpreadsheetComponent, wrap,
} from '@syncfusion/ej2-react-spreadsheet';
import { DropDownListComponent } from '@syncfusion/ej2-react-dropdowns';
import { formatUnit, registerLicense } from '@syncfusion/ej2-base';
import './App.css';
import { budgetData } from './dataSource';
import { actualSpend ,interestImage, actualInterestImage, moneyImage, varianceGraph, variancePercentage} from './dataSource';
registerLicense("IAk8BicRIAEqCzQhAR8kAxMHIgRJXmBXf013TmJbYF1xa1xPaVVBRXdVR1RpTHdDFjhoW39cdXVXRGNdUkZzXkpea0B8WHVG");


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
        
        if (args.colIndex == 11) {
            if (args.cell && args.cell.value === 'Full') {
                const wrapperDiv = createThreeSymbols1Wrapper();
                args.element.insertBefore(wrapperDiv, args.element.firstChild);
            } else if (args.cell && args.cell.value === 'Partial') {
                const wrapperDiv = createThreeSymbols2Wrapper();
                args.element.insertBefore(wrapperDiv, args.element.firstChild);
            } else if (args.cell && args.cell.value === 'Missed') {
                const wrapperDiv = createThreeSymbols3Wrapper();
                args.element.insertBefore(wrapperDiv, args.element.firstChild);
            }
        }

        if (args.colIndex == 11) {
            if (args.cell && args.cell.value === 'Low') {
                const wrapperDiv = createLowRiskIconWrapper();
                args.element.insertBefore(wrapperDiv, args.element.firstChild);
            } else if (args.cell && args.cell.value === 'Medium') {
                const wrapperDiv = createMediumRiskIconWrapper();
                args.element.insertBefore(wrapperDiv, args.element.firstChild);
            } else if (args.cell && args.cell.value === 'High') {
                const wrapperDiv = createHighRiskIconWrapper();
                args.element.insertBefore(wrapperDiv, args.element.firstChild);
            }
        }
        
    };

    const createThreeSymbols1Wrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper1';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-3symbols-1';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };

    const createThreeSymbols2Wrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-3symbols-2';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };

    const createThreeSymbols3Wrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-3symbols-3';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };

    const createLowRiskIconWrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-low-risk';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };

    const createMediumRiskIconWrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-medium-risk';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };
    const createHighRiskIconWrapper = () => {
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = 'e-custom-wrapper';
        const iconSpan = document.createElement("span");
        iconSpan.className = 'e-high-risk';
        wrapperDiv.appendChild(iconSpan);
        return wrapperDiv;
    };
    const onCreated = () => {
        spreadsheet.updateRange({ dataSource: budgetData, startCell: 'Budget!B4' }, 1);
        BudgetCalculation(spreadsheet);
        spreadsheet.updateRange({ dataSource: actualDataSource(), startCell: 'actuals!B4' }, 2);
        actualsSheetCalculation(spreadsheet);
        varianceSheetCalculation(spreadsheet);
        setTimeout(() => {
            dashboardCalculation(spreadsheet);
            sheetInsights(spreadsheet);            
        },100);
    }

    const sheetInsights = (spreadsheet) => {
        const dashboardSheet = getSheet(spreadsheet, 0);
        setCell(41, 7, dashboardSheet, { value: 'TOP 5 DELIQUENT BRANCHES', colSpan: 3, style: { fontFamily: 'Times New Roman' } });
        setCell(41, 2, dashboardSheet, { value: 'LOW PERFORMING BRANCHES', colSpan: 2, style: { fontFamily: 'Times New Roman' } });
        spreadsheet.cellFormat({ fontSize: '11pt', verticalAlign: 'middle', textAlign: 'center' }, 'DashboardInsights!B42:K42');
        const branchSpend = ['Branch', 'Planned Spend', 'Actual Spend', 'Variance'];
        const deliquentHeaders = ['Branch', 'Deliquent Accounts', 'Days Past Due(Avg)'];
        deliquentHeaders.forEach((val, i) => {
            setCell(42, i + 7, dashboardSheet, { value: val, style: { color: '#124B5C', fontWeight: 'bold' } });
        });
        branchSpend.forEach((val, i) => {
            setCell(42, i + 1, dashboardSheet, { value: val, style: { color: '#124B5C', fontWeight: 'bold' } });
        });
        spreadsheet.cellFormat({ verticalAlign: 'middle' }, 'DashboardInsights!B43:H43');
        let dashboardDataCell = 28;
        let lowPerformanceBranches = 15;
        const dataSheet = getSheet(spreadsheet, 4);
        for (let i = 43; i <= 47; i++) {
            setCell(i, 1, dashboardSheet, { formula: `=Dashboard!F${lowPerformanceBranches}` });
            setCell(i, 2, dashboardSheet, { formula: `=Dashboard!G${lowPerformanceBranches}` });
            setCell(i, 3, dashboardSheet, { formula: `=Dashboard!H${lowPerformanceBranches}` });
            setCell(i, 4, dashboardSheet, { formula: `=Dashboard!J${lowPerformanceBranches}` });
            setCell(i, 7, dashboardSheet, { formula: `=Dashboard!Q${dashboardDataCell}` });
            setCell(i, 8, dashboardSheet, { formula: `=Dashboard!R${dashboardDataCell}` });
            setCell(i, 9, dashboardSheet, { formula: `=Dashboard!S${dashboardDataCell}` });
            dashboardDataCell++;
            lowPerformanceBranches++;
        }
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt' }, 'DashboardInsights!C43:E43 I43:J48')
        spreadsheet.cellFormat({ color: '#ff0000' }, 'DashboardInsights!I44:J48 E44:E49')
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center', fontSize: '11pt' }, 'DashboardInsights!C44:E48')
        spreadsheet.setBorder({ border: '1px solid #d4cdcdc8' }, 'DashboardInsights!B41:E49 H41:J49 ', 'Outer');
        spreadsheet.setBorder({ border: '1px solid #e6e6e6' }, 'DashboardInsights!B44:E48 H44:J48', 'Horizontal');
        spreadsheet.numberFormat('$#,##0.00', 'DashboardInsights!C44:E48');
        spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'DashboardInsights!C44:C48' });
        spreadsheet.conditionalFormat({ type: 'GreenDataBar', range: 'DashboardInsights!D44:D48' });
        setCell(0, 0, dashboardSheet, { chart: chart1 });
        setCell(1, 0, dashboardSheet, { chart: chart3 });
        setCell(2, 0, dashboardSheet, { chart: chart4 });
        setCell(3, 0, dashboardSheet, { chart: chart5 });
        setCell(4, 0, dashboardSheet, { chart: chart6 });
        setCell(5, 0, dashboardSheet, { chart: chart7 });
        setCell(6, 0, dashboardSheet, { chart: chart8 });
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        setCell(0, 1, dashboardSheet, { value: 'CREDIT CARD EXPENSE DASHBOARD', colSpan: 5, style: headerStyle });
        spreadsheet.resize();
    }

    const BudgetCalculation = (spreadsheet) => {
        const sheet = getSheet(spreadsheet, 1);
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        const financeOutcomeSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' }
        const segmentSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' }
        const performanceSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2E8B3C', color: '#fff' }
        const riskSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#B03A2E', color: '#fff' }
        spreadsheet.setColumnsWidth(20, ['Budget!A','actuals!A','variance!A','DashboardInsights!F:G']);
        spreadsheet.setColumnsWidth(100,['Budget!B:D']);
        spreadsheet.setColumnsWidth(145, ['Budget!E:L']);
        spreadsheet.setColumnsWidth(135, ['DashboardInsights!B:E','DashboardInsights!N:O','DashboardInsights!I:J']);
        spreadsheet.setColumnsWidth(130, ['Variance!B:Q']);
        spreadsheet.setColumnsWidth(100, ['actuals!B:Q']);
        spreadsheet.setColumnsWidth(140,['DashboardInsights!B:E','DashboardInsights!H:J','Variance!I']);
        spreadsheet.setColumnsWidth(1,['DashboardInsights!G']);
        spreadsheet.setColumnsWidth(187,['DashboardInsights!H:J']);
        spreadsheet.setColumnsWidth(80,['DashboardInsights!K']);
        spreadsheet.setRowsHeight(40, ['Budget!3:4', 'actuals!3:4', 'Variance!6:7', 'actuals!B3:I3', 'Dashboard!9', 'Dashboard!4', 'Dashboard!21', 'Dashboard!27']);
        spreadsheet.setRowsHeight(30,['DashboardInsights!42:43', 'Variance!3:4']);
        spreadsheet.setRowsHeight(25, ['DashboardInsights!44:49','Budget!5:33', 'actuals!5:10000', 'actuals!B4:I33', 'Variance!8:37', 'Dashboard!5:8', 'Dashboard!10:20', 'Dashboard!22:24', 'Dashboard!28:37']);
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
        for (let i = 4; i < 34; i++) {
            setCell(i, 9, sheet, { formula: `=E${i + 1}* F${i + 1} * G${i + 1}` });
            setCell(i, 10, sheet, { formula: `=J${i + 1}* H${i + 1}` });
            setCell(i, 11, sheet, { formula: `=K${i + 1}* I${i + 1}` });
        }
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Budget!B3:L4');
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'Budget!B5:L34');
        spreadsheet.numberFormat('$#,##0.00','Budget!G4:G34');
        spreadsheet.numberFormat('$#,##0.00','Budget!J4:L34');
        spreadsheet.numberFormat('0.00%', 'Budget!F4:F34');
        spreadsheet.numberFormat('0.00%', 'Budget!H4:I34');
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center'}, 'Budget!F4:L34');
        spreadsheet.cellFormat({ backgroundColor: '#E6F4EA' }, 'Budget!L5:L34')
        spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'Budget!E4:E34' });
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'Budget!H4:H34' });
        spreadsheet.conditionalFormat({ type: 'ThreeTrafficLights1', range: 'Budget!F4:F34' });
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'Budget!H4:H34' });
        spreadsheet.conditionalFormat({ type: 'ThreeTrafficLights1', range: 'Budget!I4:I34' });
    }

    const actualsSheetCalculation = (spreadsheet) => {
        const sheet = getSheet(spreadsheet, 2);
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        const financeOutcomeSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' }
        const segmentSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' }
        const volumeSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' }
        const riskSubHeaderStyle =  { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#B03A2E', color: '#fff' }
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center' }, 'actuals!B4:Q1000');
        spreadsheet.cellFormat({ fontWeight: 'bold' }, 'actuals!G4:H1000');
        spreadsheet.cellFormat({ textAlign:'right'}, 'actuals!K4:K1000 O4:O1000');
        setCell(0, 1, sheet, { value: 'CREDIT CARD EXPENSE SUMMARY - ACTUALS', colSpan: 16, rowSpan: 1, style: headerStyle });
        for (let i = 1; i <= 5; i++) {
            setCell(3, i, sheet, { style: financeOutcomeSubHeaderStyle, wrap: true });
        }
        for (let i = 6; i <= 8; i++) {
            setCell(3, i, sheet, { style: volumeSubHeaderStyle, wrap: true });
        }
        for (let i = 9; i <= 12; i++) {
            setCell(3, i, sheet, { style: segmentSubHeaderStyle, wrap: true });
        }
        for (let i = 13; i <= 16; i++) {
            setCell(3, i, sheet, { style: riskSubHeaderStyle, wrap: true });
        }
        setCell(2, 1, sheet, { value: 'Segment', colSpan: 5, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
        setCell(2, 6, sheet, { value: 'Credit and Usage', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' } });
        setCell(2, 9, sheet, { value: 'Payment and Behaviour ', colSpan: 4,style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' } });
        setCell(2, 13, sheet, { value: 'Risk and Deliquency', colSpan: 4, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#8B1D18', color: '#fff' } });
        spreadsheet.numberFormat('$#,##0.00', 'actuals!F1:H1000');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!I1:K1000');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!O1:O1000');
        spreadsheet.addDataValidation({ type: 'List', value1: 'Full,Partial,Missed', ignoreBlank: false }, 'actuals!L4:L1000');
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
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '13pt', verticalAlign: 'middle', textAlign: 'center' }, 'Dashboard!A2:E2 A4:E4 A9:E9 A21:E21 A27:E27');
        const metrics = ['Total Spend', 'Total Interest'];
        const locations =['Dallas','Austin','San Antonio','Houston','Tampa','San Jose','Phoenix','New York','Chicago','Los Angeles']
        const cardTypes = ['Gold', 'Classic', 'Platinum'];
        let i = 5;
        metrics.forEach((val) => {
            setCell(i, 1, sheet, { value: val });
            i++;
        });
        spreadsheet.merge('Dashboard!A4:B4');
        spreadsheet.merge('Dashboard!D4:E4');
        setCell(5, 2, sheet, { formula: '=SUM(Variance!E4:E33)' });
        setCell(6, 2, sheet, { formula: '=SUM(Variance!I4:I33)' });
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
            setCell(26, 6 + i, sheet, { value: val });
        });
        branchSpend.forEach((val, i) => {
            setCell(8, 5 + i, sheet, { value: val });
            setCell(8, 0 + i, sheet, { value: val });
        });
        cardSpend.forEach((val, i) => {
            setCell(20, i, sheet, { value: val });
        });
        let j = 9;
        let newLocation = 27;
        locations.forEach((val) => {
            setCell(j, 0, sheet, { value: val });
            setCell(j, 5, sheet, { value: val });
            setCell(newLocation, 0, sheet, { value: val });
            setCell(newLocation, 6, sheet, { value: val });
            setCell(newLocation, 9, sheet, { value: val });
            setCell(newLocation, 11, sheet, { value: val });
            setCell(newLocation, 16, sheet, { value: val });
            setCell(newLocation, 12, sheet, { formula: `=COUNTIFS(Actuals!D4:D1004,Dashboard!L${newLocation+1},Actuals!L4:L1004,"Full")` });
            setCell(newLocation, 13, sheet, { formula: `=COUNTIFS(Actuals!D4:D1004,Dashboard!L${newLocation+1},Actuals!L4:L1004,"Partial")` });
            setCell(newLocation, 14, sheet, { formula: `=COUNTIFS(Actuals!D4:D1004,Dashboard!L${newLocation+1},Actuals!L4:L1004,"Missed")` });
            setCell(newLocation, 1, sheet, { formula: `=AVERAGEIF(Variance!C8:C37,A${newLocation+1},Variance!K8:K37)` });
            setCell(newLocation, 17, sheet, { formula: `=COUNTIFS(Actuals!D4:D1004,A${newLocation+1},Actuals!P4:P1004,"Yes")` });
            setCell(newLocation, 18, sheet, { formula: `=ROUND(AVERAGEIFS(Actuals!N4:N1004,Actuals!D4:D1004,A${newLocation+1},Actuals!P4:P1004,"Yes"),0)` });
            setCell(newLocation, 10, sheet, { formula: `=AVERAGEIF(Variance!C8:C37,Dashboard!J${newLocation+1},Variance!K8:K37)` });
            setCell(newLocation, 7, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${newLocation+1},Variance!F8:F37)` });
            newLocation++;
            j++;
        });
        
        for (let k = 9; k < 19; k++) {
            setCell(k, 2, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!F8:F37` });
            setCell(k, 3, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!J8:J37` });
            setCell(k, 4, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!G8:G37` });
            setCell(k, 1, sheet, { formula: `=SUMIF(Budget!C5:C34,Dashboard!A${k + 1},Budget!J5:J34` });
            setCell(k, 7, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!F8:F37` });
            setCell(k, 8, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!J8:J37` });
            setCell(k + 18, 8, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!J8:J37` });
            setCell(k, 9, sheet, { formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!G8:G37` });
            setCell(k, 6, sheet, { formula: `=SUMIF(Budget!C5:C34,Dashboard!A${k + 1},Budget!J5:J34` });
        }
    
        let cardTrackNumber = 21;
        cardTypes.forEach((val) => {
            setCell(cardTrackNumber, 0, sheet, { value: val });
            setCell(cardTrackNumber, 6, sheet, { value: val });
            setCell(cardTrackNumber + 6, 3, sheet, { value: val });
            setCell(cardTrackNumber + 6, 4, sheet, { formula: `=AVERAGEIF(Variance!D8:D37,D${cardTrackNumber + 7} , Variance!K8:K37)` });
            cardTrackNumber++;
        })

        for (let i = 21; i < 24; i++) {
            
            setCell(i, 1, sheet, { formula: `=SUMIF(Variance!D8:D37,Dashboard!A${i + 1},Variance!F8:F37` });
            setCell(i, 2, sheet, { formula: `=SUMIF(Variance!D8:D37,Dashboard!A${i + 1},Variance!J8:J37` });
            setCell(i, 3, sheet, { formula: `=SUMIF(Variance!D8:D37,Dashboard!A${i + 1},Variance!G8:G37` });
            setCell(i, 4, sheet, { formula: `=SUMIF(Budget!D8:D37,Dashboard!A${i + 1},Budget!J8:J37` });
            setCell(i, 7, sheet, { formula: `=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i + 1},Actuals!L4:L1004,"Full")/1000` });
            setCell(i, 8, sheet, { formula: `=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i + 1},Actuals!L4:L1004,"Partial")/1000` });
            setCell(i, 9, sheet, { formula: `=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i + 1},Actuals!L4:L1004,"Missed")/1000` });
        }

        deliquentCalc.forEach((val, i) => {
            const rowIndex = 26;
            setCell(rowIndex, 9 + i, sheet, { value: val });
            setCell(rowIndex, i, sheet, { value: val });
        })

        deliquentCalcCard.forEach((val, i) => {
            setCell(26, 3 + i, sheet, { value: val });
        })

        setCell(3, 0, sheet, { value: 'Key Metrics' });
        setCell(3, 3, sheet, { value: 'Credit Card Risk State' });
        let cardRiskNumber = 4;
        cardRisk.forEach((val) => {
            setCell(cardRiskNumber, 3, sheet, { value: val });
            setCell(cardRiskNumber, 4, sheet, { formula: `=COUNTIF(Variance!M8:M37,Dashboard!D${cardRiskNumber + 1})` });
            cardRiskNumber++;
        });

        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B10:E24');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!H28:H37');
        spreadsheet.numberFormat('0.00%', 'Dashboard!K28:K37');
        spreadsheet.numberFormat('0.00%', 'Dashboard!H22:J25');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B4:B6 ');
        spreadsheet.numberFormat('0.00%', 'Dashboard!B28:B37');
        spreadsheet.numberFormat('0.00%', 'Dashboard!E28:E30');
        spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'H', order: 'Ascending' } }, 'Dashboard!G27:H37');
        spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'S', order: 'Descending' } }, 'Dashboard!Q27:S37');
        spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'J', order: 'Descending' } }, 'Dashboard!F9:J19');
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'Variance!B8:L37');
        spreadsheet.protectSheet(0, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    }

    //chart initialization
    const chart3 = [{ type: 'Doughnut', range: 'Dashboard!B6:C7', title: 'TOTAL SPEND VS TOTAL INTEREST', theme: 'Tailwind3', height: 345, left: 55, top: 60, width: 570, id: 'Chart3', legendSettings:{position:'Right'} }];
    const chart8 = [{ type: 'Bar', range: 'Dashboard!G27:H37', title: 'TOP 10 BRANCHES BY SPEND', theme: 'Tailwind3', left: 1265, top: 60, width: 570, height: 345, id: 'Chart8', isSeriesInRows: false}];
    const chart1 = [{ type: 'Column', range: 'Dashboard!G27:I37', title: 'BRANCH WISE SPEND AND INTEREST', theme: 'Tailwind3', left: 640, top: 60, width: 570, height: 345, id: 'Chart1', isSeriesInRows: false }];
    const chart6 = [{ type: 'Doughnut', range: 'Dashboard!M38:O39', title: 'PAYMENT STATUS SUMMARY', theme: 'Tailwind3', height: 345, left: 55, top: 415, width: 570, id: 'Chart6',legendSettings:{position:'Right'}, isSeriesInRows:true, dataLabelSettings:{position:'middle', visible:'true'} }];
    const chart4 = [{ type: 'Pie', range: 'Dashboard!A27:B37', title: 'PAYMENT RISK BASED ON LOCATION', theme: 'Tailwind3', height: 345, left: 1265, width: 570, top: 415, id: 'Chart4',legendSettings:{position:'Right'}, dataLabelSettings:{position:'middle', visible:'true'} }];
    const chart5 = [{ type: 'StackingBar', range: 'Dashboard!G21:J24', title: 'PAYMENT RISK BASED ON CARD', theme: 'Tailwind3', top: 415, width: 570, left: 640, height: 345, id: 'Chart5',legendSettings:{position:'Top'},dataLabelSettings:{position:'middle', visible:'true' } }]; //
    const chart7 = [{ type: 'Column', range: 'Dashboard!A21:A24 C21:C24', title: 'INTEREST YIELD BY CARD TYPE', theme: 'Tailwind3', height: 240, left: 1265, top: 810, width: 570, id: 'Chart7', dataLabelSettings:{position:'middle', visible:'true' }}];
    
    const getCreditLimit = (branch, cardType) => {
        const base = creditLimitMatrix[branch]?.[cardType];
        return base;
    };

    const actualDataSource = () => {
        const regionBranches = {
            South: ['Dallas', 'Austin', 'San Antonio', 'Houston', 'Tampa'],
            West: ['Phoenix', 'San Jose', 'Los Angeles'],
            Northeast: ['New York'],
            Midwest: ['Chicago']
        };
        const branches = ['Dallas', 'Austin', 'San Antonio', 'Houston', 'Tampa', 'San Jose', 'Phoenix', 'New York', 'Chicago', 'Los Angeles']
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
            const creditLimit = getCreditLimit(branch, cardType);
            const statementBalance = Math.floor(Math.random() * 50000);
            // 5–10% minimum due
            const minimumDue = +(statementBalance * (0.05 + Math.random() * 0.05)).toFixed(2);
            // Payment Made (static input)
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
            // Payment Status (Derived)
            let paymentStatus = 'Missed';
            if (paymentMade >= statementBalance && statementBalance > 0) {
                paymentStatus = 'Full';
            } else if (paymentMade > 0) {
                paymentStatus = 'Partial';
            }
            // Balance Carried (Derived)
            const balanceCarried = paymentMade < statementBalance ? 'Yes' : 'No';
            // Days Past Due
            let daysPastDue = 0;
            if (paymentStatus === 'Missed' && Math.random() < 0.25) {
                daysPastDue = Math.floor(Math.random() * 30) + 1; // 1–30 days
            }
            else if (paymentStatus === 'Missed') {
                daysPastDue = Math.floor(Math.random() * 30) + 1;
            }
            const interestRate = getInterestRate(branch, cardType);
            const interestAmount = balanceCarried === 'Yes' ? +((statementBalance - paymentMade) * interestRate).toFixed(2) : 0;
            const isDelinquent = daysPastDue > 0 ? 'Yes' : 'No';
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
    const segmentSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' }
    const riskSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' }
    const performanceSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' }
    const interestImpactSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#C77700', color: '#fff' }
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
    setCell(6, 10, sheet, { style: riskSubHeaderStyle });
    setCell(6, 11, sheet, { style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' } });
    const headers = ['Region', 'Branch', 'Card Type', 'Expected Spend', 'Actual Spend', 'Variance (Actual - Expected)', 'Variance % (vs Expected)', 'Expected Interest', 'Actual Interest', 'Payment Risk %', 'Risk Level']; //'Run_Rate',
    const locations = ['Dallas', 'Austin', 'San Antonio', 'Houston', 'Tampa', 'San Jose', 'Phoenix', 'New York', 'Chicago', 'Los Angeles']
    setCell(5, 1, sheet, { value: 'Segment', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
    setCell(5, 4, sheet, { value: 'Performance', colSpan: 4, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' } });
    setCell(5, 8, sheet, { value: 'Interest Impact ', colSpan: 2, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#C77700', color: '#fff' } });
    setCell(5, 10, sheet, { value: 'Risk Score', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' } });
    setCell(5, 11, sheet, { value: 'Action Priority', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' } });
    headers.forEach((val, i) => {
        setCell(6, 1 + i, sheet, { value: val });
    });
    //image headers
    setCell(2, 1, sheet, { value: 'Estimated Spend', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
    setCell(2, 3, sheet, { value: 'Actual Spend', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
    setCell(2, 5, sheet, { value: 'Estimated Interest', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent:'30px'} });
    setCell(2, 7, sheet, { value: `Actual Interest`, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent:'30px'} });
    setCell(2, 9, sheet, { value: 'Total Variance', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt',textIndent:'30px' } });
    setCell(2, 11, sheet, { value: 'Total Variance%', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt',textIndent:'30px' } });
    //formulas for image header values
    setCell(3, 1, sheet, { formula: '=SUM(E8:E37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent:'30px' } });
    setCell(3, 3, sheet, { formula: '=SUM(F8:F37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent:'30px'}});
    setCell(3, 5, sheet, { formula: '=SUM(I8:I37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent:'30px'}});
    setCell(3, 7, sheet, { formula: '=SUM(J8:J37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent:'30px'}});
    setCell(3, 9, sheet, { formula: '=AVERAGE(G8:G37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt' }, format:'Percentage' });
    setCell(3, 11, sheet, { formula: '=AVERAGE(H8:H37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt' } });
    const cardTypes = ['Gold', 'Classic', 'Platinum'];
    let row = 37;
    for (let i = 7; i < row; i++) {
        setCell(i, 1, sheet, { formula: `=Budget!B${i - 2}` });
        setCell(i, 2, sheet, { formula: `=Budget!C${i - 2}` });
        setCell(i, 3, sheet, { formula: `=Budget!D${i - 2}` });
    }
    for (let i = 7; i < row; i++) {
        setCell(i, 4, sheet, { formula: `=SUMIFS(Budget!J4:J34, Budget!C4:C34, Variance!C${i + 1}, Budget!D4:D34, Variance!D${i + 1})` });
        setCell(i, 5, sheet, { formula: `=SUMIFS(actuals!H4:H1000, actuals!D4:D1000, Variance!C${i + 1}, actuals!E4:E1000, Variance!D${i + 1})` });
        setCell(i, 6, sheet, { formula: `=Variance!F${i + 1} - Variance!E${i + 1}`, style: { fontWeight: 'bold' } });
        setCell(i, 7, sheet, { formula: `=Variance!G${i + 1} / Variance!E${i + 1}`, style: { fontWeight: 'bold' } });
        setCell(i, 8, sheet, { formula: `=SUMIFS(Budget!L4:L34, Budget!C4:C34, Variance!C${i + 1}, Budget!D4:D34, Variance!D${i + 1})` });
        setCell(i, 9, sheet, { formula: `=SUMIFS(actuals!O4:O1000, actuals!D4:D1000, Variance!C${i + 1}, actuals!E4:E1000, Variance!D${i + 1})` });
        setCell(i, 10, sheet, {
            formula: `=COUNTIFS(actuals!D4:D1000, Variance!C${i + 1},actuals!E4:E1000, Variance!D${i + 1},actuals!N4:N1000, ">0") / COUNTIFS(actuals!D4:D1000, Variance!C${i + 1},
            actuals!E4:E1000, Variance!D${i + 1})`
        });
        setCell(i, 11, sheet, { formula: `=IF((Budget!H${i + 1}*0.4 + Variance!K${i + 1}*0.6) > 0.30, "High", IF((Budget!H${i + 1}*0.4 + Variance!K${i + 1}*0.6) > 0.20, "Medium", "Low"))` });
    }
    spreadsheet.numberFormat('$#,##0.00', 'Variance!E4:G37');
    spreadsheet.numberFormat('$#,##0.00', 'Variance!I4:J37');
    spreadsheet.numberFormat('0.00%', 'Variance!K4:K37');
    spreadsheet.numberFormat('0.00%', 'Variance!H4:H37');
    spreadsheet.numberFormat('0.00%', 'Variance!L4:L37');
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Low', range: 'Variance!ML4:L37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Medium', range: 'Variance!L4:L37', format: { style: { color: '#c7c705', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'High', range: 'Variance!L4:L37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'RedDataBar', range: 'Variance!K8:K37' });
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0', range: 'Variance!G8:G37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '0', range: 'Variance!G8:G37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'AboveAverage', range: 'Variance!L8:L37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'BelowAverage', range: 'Variance!L8:L37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0', range: 'Variance!G8:G37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '0', range: 'Variance!G8:G37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
    spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Variance!B6:L7');
    spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Variance!B3:B4 D3:D4 F3:F4 H3:H4 J3:J4 L3:L4', 'Outer');
    setCell(6, 6, sheet, { style: { fontSize: '11pt', backgroundColor: '#1E6B2D', color:'#fff',fontWeight: 'bold',textAlign:'center',verticalAlign:'middle' }, wrap: true, value:'Variance (Actual - Expected)' });
    setCell(6, 7, sheet, { style: { fontSize: '11pt', backgroundColor: '#1E6B2D', color:'#fff',fontWeight: 'bold',textAlign:'center',verticalAlign:'middle'  }, wrap: true, value:'Variance% (vs Expected)' });
    setCell(0, 0, sheet, { image: actualSpend });
    setCell(0, 2, sheet, { image: varianceGraph });
    setCell(0, 4, sheet, { image: moneyImage });
    setCell(0, 6, sheet, { image: actualInterestImage });
    setCell(0, 8, sheet, { image: interestImage });
    setCell(0, 9, sheet, { image: variancePercentage });
    spreadsheet.protectSheet(3, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
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
                <SheetDirective name='Dashboard' showGridLines={false}  state='VeryHidden' ></SheetDirective> 
               
            </SheetsDirective>
        </SpreadsheetComponent>
        
    );
}