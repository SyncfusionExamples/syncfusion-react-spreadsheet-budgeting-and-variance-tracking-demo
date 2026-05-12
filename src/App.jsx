import * as React from 'react';
import { createRoot } from 'react-dom/client';
import {
    CellDirective, CellsDirective, ColumnDirective, ColumnsDirective, getColumnHeaderText, getSheet, RangeDirective, RangesDirective, RowDirective, RowsDirective, setCell, setColumn, SheetDirective, SheetsDirective, sheetTabs, SpreadsheetComponent, wrap,
} from '@syncfusion/ej2-react-spreadsheet';
import { DropDownListComponent } from '@syncfusion/ej2-react-dropdowns';
import './App.css';
import { budgetData, actualDataSource, actualData } from './data-util';
import {
    actualSpend, interestImage, actualInterestImage, moneyImage, varianceGraph, variancePercentage, createLowRiskIconWrapper, createMediumRiskIconWrapper, createHighRiskIconWrapper,
    createThreeSymbols1Wrapper, createThreeSymbols2Wrapper, createThreeSymbols3Wrapper
} from './data-util';

export default function App() {
    //initializing spreadsheet
    let spreadsheet;
    let isdataupdate = '';
    //Adding custom icon to Payment status column
    const handleBeforeCellRender = (args) => {
        //applying icon to payment status column
        if (args.colIndex == 13) {
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

        //applying icon to risk level column
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
    
    const dataSourceChanged = (args) => {
        if (args.action === 'dataSourceChanged') {
            if (isdataupdate === 'budget') {
                budgetCalculation(spreadsheet);
                isdataupdate = 'actual';
                // load actuals after budget calculation
                spreadsheet.updateRange({ dataSource: actualDataSource(), startCell: 'actuals!B4' }, 2);
            }
            else if (isdataupdate === 'actual') {
                //actuals sheet calculations
                actualsSheetCalculation(spreadsheet);
                isdataupdate = '';
                //variance sheet calculations
                varianceSheetCalculation(spreadsheet);
                setTimeout(() => {
                    //chart data population
                    chartDataPopulation(spreadsheet);
                    //Dashboard Creation
                    dashboardInsights(spreadsheet);
                }, 100);
            }
        }
    };


    //created event
    const onCreated = () => {
        setRowColumnSize(spreadsheet);
        isdataupdate = 'budget';
        //adding datasource to budget sheet
        spreadsheet.updateRange({ dataSource: budgetData, startCell: 'Budget!B4' }, 1);
    }

    const setRowColumnSize = (spreadsheet) => {
        spreadsheet.setColumnsWidth(20, ['Budget!A', 'actuals!A', 'variance!A', 'Dashboard Insights!F:G']);
        spreadsheet.setColumnsWidth(100, ['Budget!B:D']);
        spreadsheet.setColumnsWidth(145, ['Budget!E:L']);
        spreadsheet.setColumnsWidth(135, ['Dashboard Insights!B:E', 'Dashboard Insights!N:O', 'Dashboard Insights!I:J']);
        spreadsheet.setColumnsWidth(130, ['Variance!B:Q','actuals!M']);
        spreadsheet.setColumnsWidth(131, ['Variance!G']);
        spreadsheet.setColumnsWidth(110, ['actuals!J']);
        spreadsheet.setColumnsWidth(100, ['actuals!B:I','actuals!K:L','actuals!N:S']);
        spreadsheet.setColumnsWidth(140, ['Dashboard Insights!H:J', 'Variance!I']);
        spreadsheet.setColumnsWidth(186, ['Dashboard Insights!B:D']);
        spreadsheet.setColumnsWidth(15, ['Dashboard Insights!E']);
        spreadsheet.setColumnsWidth(94, ['Dashboard Insights!F:J']);
        spreadsheet.setColumnsWidth(33, ['Dashboard Insights!K:M']);
        spreadsheet.setRowsHeight(40, ['actuals!3']);
        spreadsheet.setRowsHeight(49, ['actuals!4']);
        spreadsheet.setRowsHeight(40, ['Budget!3:4', 'Variance!6', 'actuals!B3:I3', 'Dashboard!9', 'Dashboard!4', 'Dashboard!21', 'Dashboard!27']);
        spreadsheet.setRowsHeight(46, ['Variance!7']);
        spreadsheet.setRowsHeight(30, ['Variance!3:4']);
        spreadsheet.setRowsHeight(38, ['Dashboard Insights!42:49']);
        spreadsheet.setRowsHeight(25, ['Budget!5:33', 'actuals!5:10000', 'actuals!B4:I33', 'Variance!8:37', 'Dashboard!5:8', 'Dashboard!10:20', 'Dashboard!22:24', 'Dashboard!28:37']);
    }

    const dashboardInsights = (spreadsheet) => {
        //get the dashboardsheet data
        const dashboardSheet = getSheet(spreadsheet, 0);
        //custom sort to get Top performing branches in order and Top 5 deliquent branches in order
        setTimeout(() => {
            spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'H', order: 'Ascending' } }, 'Dashboard!G27:H37')
                .then(() => spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'R', order: 'Descending' } }, 'Dashboard!Q27:S37'))
                .then(() => spreadsheet.sort({ containsHeader: true, sortDescriptors: { field: 'J', order: 'Descending' } }, 'Dashboard!F9:K19')).then(() => {
                    let dashboardDataCell = 28;
                    let lowPerformanceBranches = 19;
                    //Populating data for Table in Dashboard sheet
                    for (let i = 43; i <= 47; i++) {
                        setCell(i, 1, dashboardSheet, { formula: `=Dashboard!Q${dashboardDataCell}` });
                        setCell(i, 2, dashboardSheet, { formula: `=Dashboard!R${dashboardDataCell}` });
                        setCell(i, 3, dashboardSheet, { formula: `=Dashboard!S${dashboardDataCell}` });
                        setCell(i, 5, dashboardSheet, { formula: `=Dashboard!F${lowPerformanceBranches}` });
                        setCell(i, 6, dashboardSheet, { formula: `=Dashboard!G${lowPerformanceBranches}` });
                        setCell(i, 7, dashboardSheet, { formula: `=Dashboard!H${lowPerformanceBranches}` });
                        setCell(i, 8, dashboardSheet, { formula: `=Dashboard!J${lowPerformanceBranches}` });
                        setCell(i, 9, dashboardSheet, { formula: `=Dashboard!K${lowPerformanceBranches}` });
                        //risk score
                        setCell(i, 11, dashboardSheet, { formula: `=ROUND(50-(J${i + 1}*100*1.2),0)` });
                        dashboardDataCell++;
                        lowPerformanceBranches--;
                    }
                    //applying cell formats, borders, number formats and conditional formats to dashboard sheet
                    spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center', fontSize: '10pt', fontWeight: 'bold' }, 'Dashboard Insights!G43:L48')
                    spreadsheet.cellFormat({ textAlign: 'right', fontSize: '10pt', fontWeight: 'bold', verticalAlign: 'middle',fontFamily: 'Inter' }, 'Dashboard Insights!C44:D48');
                    spreadsheet.cellFormat({ textAlign: 'left', fontSize: '10pt', fontWeight: 'bold', verticalAlign: 'middle',fontFamily: 'Inter' }, 'Dashboard Insights!B43:B48 F43:F48');
                    spreadsheet.numberFormat('$#,##0.00', 'Dashboard Insights!G44:I48');
                    spreadsheet.numberFormat('0%', 'Dashboard Insights!J44:J48');
                    spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'Dashboard Insights!L44:L48' });
                    spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'Dashboard Insights!C44:C48' });
                    spreadsheet.conditionalFormat({ type: 'OrangeDataBar', range: 'Dashboard Insights!D44:D48' });
                    spreadsheet.cellFormat({ color: '#ff0000' }, 'Dashboard Insights!I44:J48')
                    spreadsheet.setBorder({ border: '1px solid #d4cdcdc8' }, 'Dashboard Insights!B41:D49 F41:M49 ', 'Outer');
                    spreadsheet.setBorder({ border: '1px solid #e6e6e6' }, 'Dashboard Insights!B44:D48 F44:M48', 'Horizontal');
                    spreadsheet.cellFormat({ backgroundColor: '#fff' }, 'Dashboard Insights!B41:D49 F41:M49');
                    //applying sheet protection
                    spreadsheet.protectSheet(0, { formatCells: false, formatRows: false, formatColumns: false, insertLink: false, selectCells: false });
                    //refreshing spreadsheet to update UI
                    spreadsheet.resize();
                })
                .catch(err => console.error(err));
        }, 50);
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        setCell(0, 1, dashboardSheet, { value: 'CREDIT CARD EXPENSE DASHBOARD', colSpan: 5, style: headerStyle });
        setCell(41, 1, dashboardSheet, { value: 'TOP 5 DELIQUENT BRANCHES', colSpan: 3, style: { fontFamily: 'Inter', fontSize: '11pt', verticalAlign: 'middle', textAlign: 'center' } });
        setCell(41, 5, dashboardSheet, { value: 'LOW PERFORMING BRANCHES', colSpan: 8, style: { fontFamily: 'Inter', fontSize: '11pt', verticalAlign: 'middle', textAlign: 'center' } });
        //updating Table data in Dashboard sheet
        const branchSpend = ['Branch', 'Planned Spend', 'Actual Spend', 'Variance', 'Variance %', 'Risk Score'];
        const deliquentHeaders = ['Branch', 'Deliquent Accounts', 'Days Past Due(Avg)'];
        deliquentHeaders.forEach((val, i) => {
            setCell(42, i + 1, dashboardSheet, { value: val, style: { color: '#124B5C', fontWeight: 'bold', verticalAlign: 'middle', fontFamily: 'Inter', fontSize: '10pt' } });
        });
        branchSpend.forEach((val, i) => {
            if (i == 5) {
                setCell(42, 10, dashboardSheet, { value: val, colSpan: 3, style: { color: '#124B5C', fontWeight: 'bold', verticalAlign: 'middle', fontFamily: 'Inter', fontSize: '10pt' } });
            }
            else {
                setCell(42, 5 + i, dashboardSheet, { value: val, style: { color: '#124B5C', fontWeight: 'bold', verticalAlign: 'middle', fontFamily: 'Inter', fontSize: '10pt' } });
            }
        });
        //inserting chart in the dashboard sheet
        setCell(0, 0, dashboardSheet, { chart: chart1 });
        setCell(1, 0, dashboardSheet, { chart: chart2 });
        setCell(2, 0, dashboardSheet, { chart: chart3 });
        setCell(3, 0, dashboardSheet, { chart: chart4 });
        setCell(4, 0, dashboardSheet, { chart: chart5 });
        setCell(5, 0, dashboardSheet, { chart: chart6 });
        setCell(6, 0, dashboardSheet, { chart: chart7 });
    }

    const budgetCalculation = (spreadsheet) => {
        const sheet = getSheet(spreadsheet, 1);
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        const financeOutcomeSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#7A4FA3', color: '#fff' }
        const segmentSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#124B5C', color: '#fff' }
        const performanceSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#2E8B3C', color: '#fff' }
        const riskSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#B03A2E', color: '#fff' }
        //updating segment headers and column header values for budget sheet
        setCell(0, 1, sheet, { value: 'CREDIT CARD EXPENSE SUMMARY - BUDGET', colSpan: 11, rowSpan: 1, style: headerStyle });
        setCell(2, 1, sheet, { value: 'SEGMENT', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
        setCell(2, 4, sheet, { value: 'VOLUME', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' } });
        setCell(2, 5, sheet, { value: 'PERFORMANCE', colSpan: 2, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' } });
        setCell(2, 7, sheet, { value: 'RISK', colSpan: 2, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#8B1D18', color: '#fff' } });
        setCell(2, 9, sheet, { value: 'FINANCIAL OUTCOMES', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#5E3A78', color: '#fff' } });
        setCell(3, 9, sheet, { value: 'Expected Spend', style: financeOutcomeSubHeaderStyle });
        setCell(3, 10, sheet, { value: 'Revolving Balance', style: financeOutcomeSubHeaderStyle });
        setCell(3, 11, sheet, { value: 'Expected Interest', style: financeOutcomeSubHeaderStyle });
        spreadsheet.updateCell({ style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#4B79C9', color: '#fff' } }, 'Budget!E4');
        for (let i = 1; i <= 3; i++) {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ style: segmentSubHeaderStyle }, `Budget!${col}4`);
        }
        for (let i = 5; i <= 6; i++) {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ style: performanceSubHeaderStyle }, `Budget!${col}4`);
        }
        for (let i = 7; i <= 8; i++) {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ style: riskSubHeaderStyle }, `Budget!${col}4`);
        }
        for (let i = 4; i < 34; i++) {
            const row1Based = i + 1;
            spreadsheet.updateCell({ formula: `=E${row1Based}* F${row1Based} * G${row1Based}` }, `Budget!J${row1Based}`);
            spreadsheet.updateCell({ formula: `=J${row1Based}* H${row1Based}` }, `Budget!K${row1Based}`);
            spreadsheet.updateCell({ formula: `=K${row1Based}* I${row1Based}` }, `Budget!L${row1Based}`);
        }
        //applying cell formats, borders, number formats and conditional formats to budget sheet
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Budget!B3:L4');
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'Budget!B5:L34');
        spreadsheet.numberFormat('$#,##0.00', 'Budget!G4:G34');
        spreadsheet.numberFormat('$#,##0.00', 'Budget!J4:L34');
        spreadsheet.numberFormat('0.00%', 'Budget!F4:F34');
        spreadsheet.numberFormat('0.00%', 'Budget!H4:I34');
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center' }, 'Budget!B5:D34 F4:L34');
        spreadsheet.cellFormat({backgroundColor:'#fff'}, 'Budget!B5:K34');
        spreadsheet.cellFormat({ backgroundColor: '#E6F4EA' }, 'Budget!L5:L34');
        spreadsheet.conditionalFormat({ type: 'Top10Items', range: 'Budget!L5:L34', value: '5', format: { style: { backgroundColor: '#c6efce' } } }); //#c6efce
        spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'Budget!E4:E34' });
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'Budget!H4:H34' });
        spreadsheet.conditionalFormat({ type: 'ThreeTrafficLights1', range: 'Budget!F4:F34' });
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'Budget!H4:H34' });
        spreadsheet.conditionalFormat({ type: 'ThreeTrafficLights1', range: 'Budget!I4:I34' });
        //applying sheet protection
        spreadsheet.protectSheet(1, { formatCells: false, formatRows: false, formatColumns: false, insertLink: false, selectCells: true });
    }

    const actualsSheetCalculation = (spreadsheet) => {
        const sheet = getSheet(spreadsheet, 2);
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        const financeOutcomeSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#124B5C', color: '#fff' }
        const segmentSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#2B64B3', color: '#fff' }
        const volumeSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#1E6B2D', color: '#fff' }
        const riskSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#B03A2E', color: '#fff' }
        //applying sheet name with proper styles for actuals
        setCell(0, 1, sheet, { value: 'CREDIT CARD EXPENSE SUMMARY - ACTUALS', colSpan: 16, rowSpan: 1, style: headerStyle });
        for (let i = 1; i <= 5; i++) {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ style: financeOutcomeSubHeaderStyle }, `actuals!${col}4`);
        }
        for (let i = 6; i <= 9; i++) {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ style: volumeSubHeaderStyle }, `actuals!${col}4`);
        }
        for (let i = 10; i <= 14; i++) {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ style: segmentSubHeaderStyle }, `actuals!${col}4`);
        }
        for (let i = 15; i <= 18; i++) {
            const col = String.fromCharCode(65 + i);
            spreadsheet.updateCell({ style: riskSubHeaderStyle }, `actuals!${col}4`);
        }
        setCell(2, 1, sheet, { value: 'SEGMENT', colSpan: 5, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
        setCell(2, 6, sheet, { value: 'CREDIT AND USAGE', colSpan: 4, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' } });
        setCell(2, 10, sheet, { value: 'PAYMENT BEHAVIOUR', colSpan: 5, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#2B64B3', color: '#fff' } });
        setCell(2, 15, sheet, { value: 'RISK AND DELIQUENCY', colSpan: 4, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#8B1D18', color: '#fff' } });
        //applying cell formats, borders, number formats and conditional formats to actuals sheet
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center' }, 'actuals!B4:S1004 ');
        spreadsheet.cellFormat({ fontWeight: 'bold' }, 'actuals!H4:I1004');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!F1:H1004');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!I5:I1004');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!K5:L1004');
        spreadsheet.numberFormat('$#,##0.00', 'actuals!Q1:Q1004');
        spreadsheet.numberFormat('0%', 'actuals!J5:J1004');
        spreadsheet.numberFormat('0%', 'actuals!M5:M1004');
        spreadsheet.addDataValidation({ type: 'List', value1: 'Full,Partial,Missed', ignoreBlank: false }, 'actuals!N5:N1004');
        spreadsheet.addDataValidation({ type: 'List', value1: 'Yes,No', ignoreBlank: false }, 'actuals!O5:N1004');
        spreadsheet.addDataValidation({ type: 'List', value1: 'Yes,No', ignoreBlank: false }, 'actuals!R5:R1004');
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'High', range: 'actuals!S5:S1004', cFColor: 'RedFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Low', range: 'actuals!S5:S1004', cFColor: 'GreenFT' });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Medium', range: 'actuals!S5:S1004', cFColor: 'YellowFT' });
        //deliquency 
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!R5:R1004', format: { style: { backgroundColor: '#f2f2f2', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!R5:R1004', format: { style: { color: '#ff3333', fontWeight: 'bold' }} });
        //Balance carried
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Yes', range: 'actuals!O5:O1004', format: { style: { color: '#ff3333', fontWeight: 'bold' }} });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'No', range: 'actuals!O5:O1004', cFColor: 'GreenFT' });
        //Payment made
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Missed', range: 'actuals!N5:N1004', format: { style: { color: '#DA1212', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Full', range: 'actuals!N5:N1004', format: { style: { color: '#468966', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Partial', range: 'actuals!N5:N1004', format: { style: { color: '#230338', fontWeight: 'bold' } } });
        //Days past Due
        spreadsheet.conditionalFormat({ type: 'GYRColorScale', range: 'actuals!M5:M1004' });
        spreadsheet.conditionalFormat({ type: 'RWColorScale', range: 'actuals!P5:P1004' });
        spreadsheet.conditionalFormat({ type: 'GreenDataBar', range: 'actuals!J5:J1004' });
        spreadsheet.conditionalFormat({ type: 'RedDataBar', range: 'actuals!Q5:Q1004' });
        spreadsheet.conditionalFormat({ type: 'BlueDataBar', range: 'actuals!L5:L1004' });
        spreadsheet.cellFormat({backgroundColor:'#fff'}, 'actuals!B5:S1004');
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'actuals!B3:S4');
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'actuals!B5:S1004');
        //applying sheet protection
        spreadsheet.protectSheet(2, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    };

    //Data to create chart in Dashboard Sheet
    const chartDataPopulation = (spreadsheet) => {
        const sheet = getSheet(spreadsheet, 4);
        const metrics = ['Total Spend', 'Total Interest'];
        const locations = ['Dallas', 'Austin', 'San Antonio', 'Houston', 'Tampa', 'San Jose', 'Phoenix', 'New York', 'Chicago', 'Los Angeles']
        const cardTypes = ['Gold', 'Classic', 'Platinum'];
        let cellAddress = 5; let branchData = 9; let paymentRiskData = 27;
        metrics.forEach((val) => {
            setCell(cellAddress, 1, sheet, { value: val });
            cellAddress++;
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
        setCell(8, 10, sheet, { value: 'Variance %' });
        setCell(8, 11, sheet, { value: 'Risk Score' });
        setCell(20, 6, sheet, { value: 'Branch' });
        setCell(20, 7, sheet, { value: 'Paid in Full' });
        setCell(20, 8, sheet, { value: 'Partial Payment' });
        setCell(20, 9, sheet, { value: 'Missed Payment' });
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
            setCell(8, 12 + i, sheet, { value: val });
        });

        branchSpend.forEach((val, i) => {
            setCell(8, 5 + i, sheet, { value: val });
            setCell(8, 0 + i, sheet, { value: val });
        });

        cardSpend.forEach((val, i) => {
            setCell(20, i, sheet, { value: val });
        });

        locations.forEach((val) => {
            const branchRow = branchData + 1;
            const prRow = paymentRiskData + 1;
            spreadsheet.updateCell({ value: val }, `Dashboard!A${branchRow}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!F${branchRow}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!M${branchRow}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!A${prRow}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!G${prRow}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!J${prRow}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!L${prRow}`);
            spreadsheet.updateCell({ value: val }, `Dashboard!Q${prRow}`);
            spreadsheet.updateCell({ formula: `=COUNTIFS(Actuals!D4:D1004,Dashboard!L${prRow},Actuals!N4:N1004,"Full")` }, `Dashboard!M${prRow}`);
            spreadsheet.updateCell({ formula: `=COUNTIFS(Actuals!D4:D1004,Dashboard!L${prRow},Actuals!N4:N1004,"Partial")` }, `Dashboard!N${prRow}`);
            spreadsheet.updateCell({ formula: `=COUNTIFS(Actuals!D4:D1004,Dashboard!L${prRow},Actuals!N4:N1004,"Missed")` }, `Dashboard!O${prRow}`);
            spreadsheet.updateCell({ formula: `=AVERAGEIF(Variance!C8:C37,A${prRow},Variance!K8:K37)` }, `Dashboard!B${prRow}`);
            spreadsheet.updateCell({ formula: `=COUNTIFS(Actuals!D4:D1004,A${prRow},Actuals!R4:R1004,"Yes")` }, `Dashboard!R${prRow}`);
            spreadsheet.updateCell({ formula: `=ROUND(AVERAGEIFS(Actuals!P4:P1004,Actuals!D4:D1004,A${prRow},Actuals!R4:R1004,"Yes"),0)` }, `Dashboard!S${prRow}`);
            spreadsheet.updateCell({ formula: `=AVERAGEIF(Variance!C8:C37,Dashboard!J${prRow},Variance!K8:K37)` }, `Dashboard!K${prRow}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${prRow},Variance!F8:F37` }, `Dashboard!H${prRow}`);
            paymentRiskData++;
            branchData++;
        });

        for (let k = 9; k < 19; k++) {
            //data population
            spreadsheet.updateCell({ formula: `=SUMIF(Budget!C5:C34,Dashboard!A${k + 1},Budget!J5:J34)` }, `Dashboard!B${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!F8:F37)` }, `Dashboard!C${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!J8:J37)` }, `Dashboard!D${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!A${k + 1},Variance!G8:G37)` }, `Dashboard!E${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Budget!C5:C34,Dashboard!F${k + 1},Budget!J5:J34)` }, `Dashboard!G${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!F${k + 1},Variance!F8:F37)` }, `Dashboard!H${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!F${k + 1},Variance!J8:J37)` }, `Dashboard!I${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!F${k + 1},Variance!G8:G37)` }, `Dashboard!J${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!F${k + 1},Variance!G8:G37) /SUMIF(Variance!C8:C37,Dashboard!F${k + 1},Variance!E8:E37)` }, `Dashboard!K${k + 1}`);
            spreadsheet.updateCell({ formula: `=Dashboard!C${k + 1}` }, `Dashboard!N${k + 1}`);
            spreadsheet.updateCell({ formula: `=Dashboard!D${k + 1}` }, `Dashboard!O${k + 1}`);
            spreadsheet.updateCell({ formula: `=SUMIF(Variance!C8:C37,Dashboard!G${k + 19},Variance!J8:J37)` }, `Dashboard!I${k + 19}`); // corresponds to setCell(k+18,8,...)
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
            //payment risk
            setCell(i, 7, sheet, { formula: `=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i + 1},Actuals!N4:N1004,"Full")/1000` });
            setCell(i, 8, sheet, { formula: `=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i + 1},Actuals!N4:N1004,"Partial")/1000` });
            setCell(i, 9, sheet, { formula: `=COUNTIFS(Actuals!E4:E1004,Dashboard!G${i + 1},Actuals!N4:N1004,"Missed")/1000` });
        }

        deliquentCalc.forEach((val, i) => {
            const rowIndex = 26;
            setCell(rowIndex, 9 + i, sheet, { value: val });
            setCell(rowIndex, i, sheet, { value: val });
        })

        deliquentCalcCard.forEach((val, i) => {
            setCell(26, 3 + i, sheet, { value: val });
        })
        
        let cardRiskNumber = 4;
        cardRisk.forEach((val) => {
            setCell(cardRiskNumber, 3, sheet, { value: val });
            setCell(cardRiskNumber, 4, sheet, { formula: `=COUNTIF(Variance!M8:M37,Dashboard!D${cardRiskNumber + 1})` });
            cardRiskNumber++;
        });

        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B10:E24');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!H28:H37');
        spreadsheet.numberFormat('0%', 'Dashboard!K28:K37');
        spreadsheet.numberFormat('0%', 'Dashboard!H22:J25');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!B4:B6 ');
        spreadsheet.numberFormat('$#,##0.00', 'Dashboard!N10:O19 ');
        spreadsheet.numberFormat('0%', 'Dashboard!B28:B37');
        spreadsheet.numberFormat('0%', 'Dashboard!E28:E30');
        spreadsheet.setBorder({ border: '1px solid #f2f2f2' }, 'Variance!B8:L37');
    }

    //chart initialization
    const chart1 = [{ type: 'Column', range: 'Dashboard!M9:O19', title: 'BRANCH WISE SPEND AND INTEREST', theme: 'Tailwind3', left: 640, top: 60, width: 570, height: 345, id: 'Chart1', isSeriesInRows: false }];
    const chart2 = [{ type: 'Doughnut', range: 'Dashboard!B6:C7', title: 'TOTAL SPEND VS TOTAL INTEREST', theme: 'Tailwind3', height: 345, left: 55, top: 60, width: 570, id: 'Chart2', legendSettings: { position: 'Right' } }];
    const chart3 = [{ type: 'Pie', range: 'Dashboard!A27:B37', title: 'PAYMENT RISK BASED ON LOCATION', theme: 'Tailwind3', height: 345, left: 1225, width: 570, top: 415, id: 'Chart3', legendSettings: { position: 'Right' }, dataLabelSettings: { position: 'middle', visible: 'true' } }];
    const chart4 = [{ type: 'StackingBar', range: 'Dashboard!G21:J24', title: 'PAYMENT RISK BASED ON CARD', theme: 'Tailwind3', top: 415, width: 570, left: 640, height: 345, id: 'Chart4', legendSettings: { position: 'Top' }, dataLabelSettings: { position: 'middle', visible: 'true' } }];
    const chart5 = [{ type: 'Doughnut', range: 'Dashboard!M38:O39', title: 'PAYMENT STATUS SUMMARY', theme: 'Tailwind3', height: 345, left: 55, top: 415, width: 570, id: 'Chart5', legendSettings: { position: 'Right' }, isSeriesInRows: true, dataLabelSettings: { position: 'middle', visible: 'true' } }];
    const chart6 = [{ type: 'Column', range: 'Dashboard!A21:A24 C21:C24', title: 'INTEREST YIELD BY CARD TYPE', theme: 'Tailwind3', height: 330, left: 1225, top: 815, width: 570, id: 'Chart6', dataLabelSettings: { position: 'middle', visible: 'true' } }];
    const chart7 = [{ type: 'Bar', range: 'Dashboard!G27:H37', title: 'TOP 10 BRANCHES BY SPEND', theme: 'Tailwind3', left: 1225, top: 60, width: 570, height: 345, id: 'Chart7', isSeriesInRows: false, dataLabelSettings: { position: 'Outer', visible: 'true' } }];

    const varianceSheetCalculation = (spreadsheet) => {
        const sheet = getSheet(spreadsheet, 3);
        //applying Header style
        const segmentSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#124B5C', color: '#fff' }
        const riskSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#7A4FA3', color: '#fff' }
        const performanceSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#1E6B2D', color: '#fff' }
        const interestImpactSubHeaderStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#C77700', color: '#fff' }
        const headerStyle = { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '22pt', color: '#1E6B2D' };
        const segmentHeader = ['Region', 'Branch', 'Card Type'];
        const performanceHeader = ['Expected Spend', 'Actual Spend', 'Variance\n(Actual - Expected)', 'Variance%\n(vs Expected)'];
        const interestImpactHeader = ['Expected Interest', 'Actual Interest'];
        setCell(0, 1, sheet, { value: 'CREDIT CARD EXPENSE SUMMARY - VARIANCE ANALYSIS', colSpan: 12, rowSpan: 1, style: headerStyle });
        setCell(6, 10, sheet, { style: riskSubHeaderStyle, value: 'Payment Risk' });
        setCell(6, 11, sheet, { style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '12pt', backgroundColor: '#124B5C', color: '#fff' }, value: 'Risk Level' });
        setCell(5, 1, sheet, { value: 'SEGMENT', colSpan: 3, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#0E3A4A', color: '#fff' } });
        setCell(5, 4, sheet, { value: 'PERFORMANCE', colSpan: 4, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#1E6B2D', color: '#fff' } });
        setCell(5, 8, sheet, { value: 'INTEREST IMPACT', colSpan: 2, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#C77700', color: '#fff' } });
        setCell(5, 10, sheet, { value: 'RISK SCORE', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#7A4FA3', color: '#fff' } });
        //segment header style
        for (let i = 1; i <= 3; i++) {
            setCell(6, i, sheet, { style: segmentSubHeaderStyle, value: segmentHeader[i - 1] });
        }
        
        //performance header style
        for (let i = 4; i <= 7; i++) {
            setCell(6, i, sheet, { style: performanceSubHeaderStyle, value: performanceHeader[i - 4] });
        }
        
        //Interest Impact style
        for (let i = 8; i <= 9; i++) {
            setCell(6, i, sheet, { style: interestImpactSubHeaderStyle, value: interestImpactHeader[i - 8] });
        }
        
        //Action Priority style
        setCell(5, 11, sheet, { value: 'Action Priority', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'center', fontSize: '14pt', backgroundColor: '#124B5C', color: '#fff' } });
        
        //image headers
        setCell(2, 1, sheet, { value: 'Estimated Spend', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(2, 3, sheet, { value: 'Actual Spend', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(2, 5, sheet, { value: 'Estimated Interest', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(2, 7, sheet, { value: `Actual Interest`, style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(2, 9, sheet, { value: 'Total Variance', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(2, 11, sheet, { value: 'Total Variance%', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        
        //formulas for image header values
        setCell(3, 1, sheet, { formula: '=SUM(E8:E37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(3, 3, sheet, { formula: '=SUM(F8:F37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(3, 5, sheet, { formula: '=SUM(I8:I37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(3, 7, sheet, { formula: '=SUM(J8:J37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt', textIndent: '30px' } });
        setCell(3, 9, sheet, { formula: '=AVERAGE(G8:G37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt' }, format: 'Percentage' });
        setCell(3, 11, sheet, { formula: '=AVERAGE(H8:H37)', style: { fontWeight: 'bold', verticalAlign: 'middle', textAlign: 'right', fontSize: '9pt' } });
        for (let i = 7; i < 37; i++) {
            setCell(i, 1, sheet, { formula: `=Budget!B${i - 2}` });
            setCell(i, 2, sheet, { formula: `=Budget!C${i - 2}` });
            setCell(i, 3, sheet, { formula: `=Budget!D${i - 2}` });
        }
        
        //formula calculation
        const varianceSheetColumnStyle = { verticalAlign: 'middle', textAlign: 'center' };
        const varianceColumnStyle = { verticalAlign: 'middle', textAlign: 'center', fontWeight: 'bold' };
        for (let i = 7; i < 37; i++) {
            setCell(i, 4, sheet, { formula: `=SUMIFS(Budget!J4:J34, Budget!C4:C34, Variance!C${i + 1}, Budget!D4:D34, Variance!D${i + 1})`, style: varianceSheetColumnStyle });
            setCell(i, 5, sheet, { formula: `=SUMIFS(actuals!H4:H1004, actuals!D4:D1004, Variance!C${i + 1}, actuals!E4:E1004, Variance!D${i + 1})`, style: varianceSheetColumnStyle });
            setCell(i, 6, sheet, { formula: `=Variance!F${i + 1} - Variance!E${i + 1}`, style: varianceColumnStyle });
            setCell(i, 7, sheet, { formula: `=Variance!G${i + 1} / Variance!E${i + 1}`, style: varianceColumnStyle });
            setCell(i, 8, sheet, { formula: `=SUMIFS(Budget!L4:L34, Budget!C4:C34, Variance!C${i + 1}, Budget!D4:D34, Variance!D${i + 1})`, style: varianceSheetColumnStyle });
            setCell(i, 9, sheet, { formula: `=SUMIFS(actuals!Q4:Q1004, actuals!D4:D1004, Variance!C${i + 1}, actuals!E4:E1004, Variance!D${i + 1})`, style: varianceSheetColumnStyle });
            setCell(i, 10, sheet, {
                formula: `=COUNTIFS(actuals!D4:D1004, Variance!C${i + 1},actuals!E4:E1004, Variance!D${i + 1},actuals!P4:P1004, ">0") / COUNTIFS(actuals!D4:D1004, Variance!C${i + 1},
            actuals!E4:E1004, Variance!D${i + 1})`
            });
            setCell(i, 11, sheet, { formula: `=IF((Budget!H${i + 1}*0.4 + Variance!K${i + 1}*0.6) > 0.30, "High", IF((Budget!H${i + 1}*0.4 + Variance!K${i + 1}*0.6) > 0.20, "Medium", "Low"))`, style: varianceSheetColumnStyle });
        }
        
        //applying number format, conditional format, setborder
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center' },'Variance!B8:D37');
        spreadsheet.numberFormat('$#,##0.00', 'Variance!E8:G37');
        spreadsheet.numberFormat('$#,##0.00', 'Variance!I8:J37');
        spreadsheet.numberFormat('0%', 'Variance!K8:K37');
        spreadsheet.numberFormat('0.00%', 'Variance!H8:H37');
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Low', range: 'Variance!L8:L37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Medium', range: 'Variance!L8:L37', format: { style: { color: '#c7c705', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'High', range: 'Variance!L8:L37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'RedDataBar', range: 'Variance!K8:K37' });
        spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0', range: 'Variance!G8:H37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'LessThan', value: '0', range: 'Variance!G8:H37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'ThreeArrows', range: 'Variance!H8:H37' });
        spreadsheet.conditionalFormat({ type: 'AboveAverage', range: 'Variance!L8:L37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'BelowAverage', range: 'Variance!L8:L37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0', range: 'Variance!G8:G37', format: { style: { color: '#375623', fontWeight: 'bold' } } });
        spreadsheet.conditionalFormat({ type: 'LessThan', value: '0', range: 'Variance!G8:G37', format: { style: { color: '#ff3333', fontWeight: 'bold' } } });
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Variance!B6:L7');
        spreadsheet.setBorder({ border: '1px solid #cccccc' }, 'Variance!B3:B4 D3:D4 F3:F4 H3:H4 J3:J4 L3:L4', 'Outer');
        spreadsheet.cellFormat({backgroundColor:'#fff'}, 'Variance!B3:B4 D3:D4 F3:F4 H3:H4 J3:J4 L3:L4 B8:L37');

        //Image insertion
        setCell(0, 0, sheet, { image: actualSpend });
        setCell(0, 2, sheet, { image: varianceGraph });
        setCell(0, 4, sheet, { image: moneyImage });
        setCell(0, 6, sheet, { image: actualInterestImage });
        setCell(0, 8, sheet, { image: interestImage });
        setCell(0, 9, sheet, { image: variancePercentage });
        //applying sheet protection
        spreadsheet.protectSheet(3, { selectCells: true, selectUnLockedCells: true, formatCells: true, formatRows: true, formatColumns: true, insertLink: false });
    };

    return (
        <SpreadsheetComponent height={950} ref={(ssObj) => { spreadsheet = ssObj }} created={onCreated.bind(this)} beforeCellRender={handleBeforeCellRender}  dataSourceChanged={dataSourceChanged}
        cellStyle={{ backgroundColor: '#fafafa' }}
            openUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/open'
            saveUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/save' >
            <SheetsDirective>
                <SheetDirective name='Dashboard Insights' showGridLines={false}></SheetDirective>
                <SheetDirective name='Budget' showGridLines={false}></SheetDirective>
                <SheetDirective name='Actuals' showGridLines={false}></SheetDirective>
                <SheetDirective name='Variance' showGridLines={false}></SheetDirective>
                <SheetDirective name='Dashboard' showGridLines={false} state='VeryHidden'></SheetDirective>
            </SheetsDirective>
        </SpreadsheetComponent>
    );
}