import { getColumn, getFormatFromType, getSheet, setCell, SheetDirective, SheetsDirective, Spreadsheet, SpreadsheetComponent, type CellStyleModel, type SheetModel } from '@syncfusion/ej2-react-spreadsheet';
import { employeeData, timeData } from './Data'

export default function App() {
  
  //initializing spreadsheet
  let spreadsheet: Spreadsheet;
  //Global variable to load data based on sheet
  let dataLoaded: string = 'EmployeeMaster';
  let employeeSheetName: string = 'EmployeeMaster';
  let timeSheetName: string = 'Timesheet';
  let payrollSheetName: string = 'Payroll';
  let dataSheetName: string = 'Data';
  //Header cell style
  const headerStyle: CellStyleModel = { fontSize: '14pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#001A4A', color: '#fff', fontWeight: 'bold' };
  const columnHeaderStyle: CellStyleModel = { fontSize: '12pt', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#001A4A', color: '#fff', fontWeight: 'bold' };
  //created function
  const onCreated = (): void => {
    setFormats();
    setRowColumnSize();
    spreadsheet.updateRange({ dataSource: employeeData as any, startCell: 'A5' }, 1);
  }

  //datasource changed event
  const dataSourceChanged = (args: any): void => {
    if (args.action === 'dataSourceChanged') {
      if (dataLoaded === 'EmployeeMaster') {
        
        initiateEmployeeSheet();
        dataLoaded = 'TimeSheet';
        spreadsheet.updateRange({ dataSource: timeData as any, startCell: 'A5' }, 2);
      }
      else if (dataLoaded === 'TimeSheet') {
        initiateTimeSheet();
        initiatePayrollSheet();
        initiateDataSheet();
      }
    }
  }
 
  const setRowColumnSize: Function = (): void => {
    //set the row height
    spreadsheet.setRowsHeight(30, [`${employeeSheetName}!1:57`, `${timeSheetName}!1:1505`, `${payrollSheetName}!1:56`]);
    //set the column width
    spreadsheet.setColumnsWidth(100, [`${employeeSheetName}!A:H`,`${timeSheetName}!A:J`,`${payrollSheetName}!A:L`]);
  }

  const setFormats: Function = (): void => {
    //setting sheetmodels
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    const timeSheet: SheetModel = getSheet(spreadsheet, 2);
    const payrollSheet: SheetModel = getSheet(spreadsheet, 3);
    
    //Employee Sheet Formats
    //set header value
    setCell(0, 0, employeeSheet, { value: 'EMPLOYEE MASTER SHEET- HR PAYROLL AND TIMESHEET APPLICATION', colSpan: 8, rowSpan: 3, style: headerStyle });
    //updating styles to headers
    for (let columnHeader = 0; columnHeader < 8; columnHeader++) {
      setCell(4, columnHeader, employeeSheet, { style: columnHeaderStyle });
    }

    //Time Sheet Formats
    //updating styles to timesheet headers
    for (let columnHeader = 0; columnHeader < 10; columnHeader++) {
      setCell(4, columnHeader, timeSheet, { style: columnHeaderStyle });
    }

    //set header value
    setCell(0, 0, timeSheet, { value: 'TIME SHEET - HR PAYROLL AND TIMESHEET APPLICATION', colSpan: 10, rowSpan: 3, style: headerStyle });
    //Payroll Sheet Formats
    //set header value
    setCell(0, 0, payrollSheet, { value: 'PAYROLL SHEET - HR PAYROLL AND TIMESHEET APPLICATION', colSpan: 12, rowSpan: 3, style: headerStyle });
  }

  //EmployeeSheet Calculations
  const initiateEmployeeSheet: Function = (): void => {
    //get the sheet
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    //Employee ID Columns
    spreadsheet.cellFormat({ color: '#2549BE', fontWeight: 'bold' }, `${employeeSheetName}!A6:A${employeeSheet.rows.length}`);
    //department color conditional formatttings
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'HR', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D5DEE9'}} });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Finance', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D7E7C7'}} });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'IT', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D5C2E6'}} });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Operations', range: `${employeeSheetName}!C6:C${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#DADADA'}} });
    //work location conditional formattings
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Dallas', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D6DDE6'}} });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'San Francisco', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#F2E3B6'}} });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Austin', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D7E6DB'}} });
    spreadsheet.conditionalFormat({ type: 'EqualTo', value: 'Seattle', range: `${employeeSheetName}!E6:E${employeeSheet.rows.length}`, format: { style: { backgroundColor: '#D9D6E6'}} });
    //set data bars to base salary
    spreadsheet.conditionalFormat({ type: 'LightBlueDataBar', range: `${employeeSheetName}!F6:F${employeeSheet.rows.length}` });
    spreadsheet.conditionalFormat({ type: 'GWRColorScale', range: `${employeeSheetName}!G6:G${employeeSheet.rows.length}` });
    //formula update for OT rate
    for (let row = 5; row < employeeSheet.rows.length - 1; row++) {
      setCell(row, 6, employeeSheet, { formula: `=(F${row + 1}/(30*8))*1.5`});
    }
  }

  //TimeSheet Calculations
  const initiateTimeSheet: Function = (): void => {
    //get the sheet
    const timeSheet: SheetModel = getSheet(spreadsheet, 2);
    //conditional formattings
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '9:30', range: `${timeSheetName}!C6:C${timeSheet.rows.length}`, format: { style: { color: '#ff0000', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '9:31', range: `${timeSheetName}!C6:C${timeSheet.rows.length}`, format: { style: { color: '#4CAF50', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '5:31', range: `${timeSheetName}!D6:D${timeSheet.rows.length}`, format: { style: { color: '#ff0000', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '5:30', range: `${timeSheetName}!D6:D${timeSheet.rows.length}`, format: { style: { color: '#4CAF50', fontWeight: 'bold' } } });
    //Lunch
    spreadsheet.cellFormat({backgroundColor:'#f1deab'},`${timeSheetName}!E6:E${timeSheet.rows.length}`);
    //Permission
    spreadsheet.conditionalFormat({ type: 'LessThan', value: '1:01', range: `${timeSheetName}!F6:F${timeSheet.rows.length}`, format: { style: { color: '#4CAF50', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'Between', value: '1:01,2:01', range: `${timeSheetName}!F6:F${timeSheet.rows.length}`, format: { style: { color: '#FFC107', fontWeight: 'bold' } } });
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '2:00', range: `${timeSheetName}!F6:F${timeSheet.rows.length}`, format: { style: { color: '#FF0000', fontWeight: 'bold' } } });
    //Break
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0.014', range: `${timeSheetName}!G6:G${timeSheet.rows.length}`, format: { style: { color: '#FF0000', fontWeight: 'bold' } }});
    //work hours
    spreadsheet.conditionalFormat({ type: 'LightBlueDataBar', range: `${timeSheetName}!H6:H${timeSheet.rows.length}` });
    //overtime hours
    spreadsheet.conditionalFormat({ type: 'OrangeDataBar', range: `${timeSheetName}!I6:I${timeSheet.rows.length}` });
    //formula update for permissions, work hours and overtime
    for (let row = 5; row <= timeSheet.rows.length - 1; row++) {
      //permission formula
      setCell(row, 5, timeSheet, { formula: `=IF(C${row + 1}>TIME(9,30,0), TIME(0,30,0), 0)`, format: 'h:mm' });
      //workhour formula
      setCell(row, 7, timeSheet, { formula: `=(D${row + 1}-C${row + 1})-(MAX(E${row + 1}-TIME(1,0,0),0)+F${row + 1}+MAX(G${row + 1}-TIME(0,20,0),0))`, format: 'h:mm' });
      //overtime formula
      setCell(row, 8, timeSheet, { formula: `=IF(H${row + 1}>TIME(8,0,0),H${row + 1}-TIME(8,0,0),0)`, format: 'h:mm' });
    }
  }

  //Payroll Calculations
  const initiatePayrollSheet: Function = (): void => {
    //get the employee sheet
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    //get the payroll sheet
    const payrollSheet: SheetModel = getSheet(spreadsheet, 3);
    //payroll sheet header update
    const payrollHeaders: string[] = ['Employee Id', 'Base Salary', 'Total Hours', 'OT Hours', 'OT Pay', 'Leave Days', 'Leave Deduction', 'Social Contribution %', 'Social Contribution', 'Tax%', 'Tax', 'Net Salary'];
    payrollHeaders.forEach((value, index) => {
      setCell(4, index, payrollSheet, { value: value, style:columnHeaderStyle });
    });
    //formula update
    for (let row = 5; row <= employeeSheet.rows.length - 1; row++) {
      //employee id
      setCell(row, 0, payrollSheet, { formula: `=${employeeSheetName}!A${row + 1}` });
      //base salary
      setCell(row, 1, payrollSheet, { formula: `=VLOOKUP(A${row + 1},${employeeSheetName}!A1:G56,6,FALSE)`, format:getFormatFromType('Currency') });
      //total hours
      setCell(row, 2, payrollSheet, { formula: `=SUMIFS(${timeSheetName}!H1:H1500,${timeSheetName}!B1:B1500,A${row + 1})`, format: '[h]:mm' });
      //ot hours
      setCell(row, 3, payrollSheet, { formula: `=SUMIFS(Timesheet!I6:I1500,Timesheet!B6:B1500,A${row + 1})`, format: '[h]:mm' });
      //ot pay
      setCell(row, 4, payrollSheet, { formula: `=D${row + 1}*24*((B${row + 1}/240)*1.5)`});
      //leave days
      setCell(row, 5, payrollSheet, { formula: `=COUNTIFS(Timesheet!J6:J1500,"LOP",Timesheet!B6:B1500,A${row + 1})`});
      //leave deduction
      setCell(row, 6, payrollSheet, { formula: `=(B${row + 1}/30)*F${row + 1}`});
      //social contribution %
      setCell(row, 7, payrollSheet, { value: '0.12', format:`${getFormatFromType('Percentage')}`});
      //Contribution
      setCell(row, 8, payrollSheet, { formula: `=B${row + 1}*H${row + 1}`, format:`${getFormatFromType('Currency')}`});
      //Tax %
      setCell(row, 9, payrollSheet, { formula: `=IF(B${row + 1}*12<=17700,0.10,IF(B${row + 1}*12<=67450,0.12,IF(B${row + 1}*12<=105700,0.22,IF(B${row + 1}*12<=201750,0.24,IF(B${row + 1}*12<=256200,0.32,IF(B${row + 1}*12<=640600,0.35,0.37))))))`, format:`${getFormatFromType('Percentage')}`});
      //Tax
      setCell(row, 10, payrollSheet, { formula: `=B${row + 1}*J${row + 1}`, format:`${getFormatFromType('Currency')}`});
      //net salary
      setCell(row, 11, payrollSheet, { formula: `=B${row + 1}+E${row + 1}-(G${row + 1}+I${row + 1}+K${row + 1})`});
    }

    //Formatting
    spreadsheet.cellFormat({ color: '#2549BE', fontWeight: 'bold' }, `${payrollSheetName}!A6:A${payrollSheet.rows.length}`);
    //Data bar
    spreadsheet.conditionalFormat({ type: 'GreenDataBar', range: `${payrollSheetName}!B6:B${payrollSheet.rows.length}` });
    //Leave Deduction
    spreadsheet.conditionalFormat({ type: 'GreaterThan', value: '0', range: `${payrollSheetName}!G6:G${payrollSheet.rows.length}`, format: { style: { color: '#FF0000', fontWeight: 'bold' } }});
    //PF
    spreadsheet.conditionalFormat({ type: 'ThreeSymbols', range: `${payrollSheetName}!H6:H${payrollSheet.rows.length}` });
    
  }

  //DataSheet Calculations
  const initiateDataSheet: Function = (): void => {
    //get employee sheet model
    const employeeSheet: SheetModel = getSheet(spreadsheet, 1);
    const employeeSheetLastRow: number = employeeSheet.rows.length;
    //get timesheet model
    const timeSheet: SheetModel = getSheet(spreadsheet, 2);
    const timeSheetLastRow: number = timeSheet.rows.length;
    //get data sheet model
    const dataSheet: SheetModel = getSheet(spreadsheet, 4);
    //populating data for payroll cost and OT hours by department for charts
    const department: string[] = ['IT','HR','Finance','Operations'];
    department.forEach((value: string, index: number) => {
      setCell(3 + index, 0, dataSheet, { value: value });
      //formula for the standard payroll cost
      setCell(3 + index, 1, dataSheet, { formula:`=SUMIF(${employeeSheetName}!C1:C${employeeSheetLastRow}, ${dataSheetName}!A${3 + index}, ${employeeSheetName}!F1:F${employeeSheetLastRow})` });
      setCell(9 + index, 0, dataSheet, { value: value });
    });
    //OT Hrs Distribution
    const otHrs: string[] = ['0-1Hrs','1-2Hrs','2-4Hrs'];
    otHrs.forEach((value: string, index: number) => {
      setCell(3 + index, 2, dataSheet, { value: value });
    });

    //formulas for ot hrs distribution
    setCell(3 , 3, dataSheet, { formula: `=COUNTIFS(${timeSheetName}!I6:I${timeSheetLastRow},">=0",${timeSheetName}!I6:I${timeSheetLastRow},"<=01:00:00")` });
    setCell(4, 3, dataSheet, { formula: `=COUNTIFS(${timeSheetName}!I6:I${timeSheetLastRow},">01:00:00",${timeSheetName}!I6:I${timeSheetLastRow},"<=02:00:00")` });
    setCell(5, 3, dataSheet, { formula: `=COUNTIFS(${timeSheetName}!I6:I${timeSheetLastRow},">02:00:00",${timeSheetName}!I6:I${timeSheetLastRow},"<=04:00:00")` });
    //formulas for ot hrs
    setCell(9 , 1, dataSheet, { formula: `=SUMIFS(I6:I${timeSheetLastRow}, G6:G${timeSheetLastRow}, A10)` });
    setCell(10 , 1, dataSheet, { formula: `=SUMIFS(I6:I${timeSheetLastRow}, G6:G${timeSheetLastRow}, A11)` });
    setCell(11, 1, dataSheet, { formula: `=SUMIFS(I6:I${timeSheetLastRow}, G6:G${timeSheetLastRow}, A12)` });
    setCell(12 , 1, dataSheet, { formula: `=SUMIFS(I6:I${timeSheetLastRow}, G6:G${timeSheetLastRow}, A13)` });
    //Employee Average working %
    const employeeAverageWorking = ['<6Hrs','6-8Hrs','8-10Hrs','>10Hrs'];
    employeeAverageWorking.forEach((value: string, index: number) => {
      setCell(9 + index, 2, dataSheet, { value: value });
    });

    //formulas for average working %
    setCell(9, 3, dataSheet, { formula: `=COUNTIFS(H4:H${employeeSheetLastRow},"<0.25")`, format:'h:mm' });
    setCell(10, 3, dataSheet, { formula: `=COUNTIFS(H4:H${employeeSheetLastRow},">=0.25",H4:H${employeeSheetLastRow},"<=0.3333333")`, format:'h:mm' });
    setCell(11, 3, dataSheet, { formula: `=COUNTIFS(H4:H${employeeSheetLastRow},">0.3333333",H4:H${employeeSheetLastRow},"<=0.4166667")`, format:'h:mm' });
    setCell(12, 3, dataSheet, { formula: `=COUNTIFS(H4:H${employeeSheetLastRow},">0.4166667")`, format:'h:mm' });

    //Employee Data Summarized
    const employeeMetric = ['EmpId','Department','Avg Work Hrs','Total OT Hrs','Leaves Taken','Late Login','Risk Score'];
    employeeMetric.forEach((value: string, index: number) => {
      setCell(4, 5 + index, dataSheet, { value: value });
    });

    for(let employeeRow = 5 ; employeeRow < employeeSheetLastRow ; employeeRow++){
      //get employeeID
      setCell(employeeRow, 5, dataSheet, { formula:`=${employeeSheetName}!A${employeeRow + 1}` });
      //get department
      setCell(employeeRow, 6, dataSheet, { formula:`=${employeeSheetName}!C${employeeRow + 1}` });
      //get avg work hrs
      setCell(employeeRow, 7, dataSheet, { formula:`=AVERAGEIF(${timeSheetName}!B6:B${timeSheetLastRow}, F${employeeRow + 1}, ${timeSheetName}!H6:H${timeSheetLastRow})`, format: 'h:mm'});
      //get total ot hrs
      setCell(employeeRow, 8, dataSheet, { formula:`=SUMIF(${timeSheetName}!B6:B${timeSheetLastRow},F${employeeRow + 1},${timeSheetName}!I6:I${timeSheetLastRow})` });
      //get leaves taken
      setCell(employeeRow, 9, dataSheet, { formula:`=COUNTIFS(${timeSheetName}!B6:B${timeSheetLastRow},F${employeeRow + 1},${timeSheetName}!J6:J${timeSheetLastRow},"<>None")` });
      //get late logins
      setCell(employeeRow, 10, dataSheet, { formula:`=COUNTIFS(${timeSheetName}!B6:B${timeSheetLastRow},F${employeeRow + 1},${timeSheetName}!F6:F${timeSheetLastRow},">0")` });
      //get risk score
      setCell(employeeRow, 11, dataSheet, { formula:`=MIN(I${employeeRow + 1}*24/60*35,35)+MIN(ABS(H${employeeRow + 1}*24-8)/4*25,25)+MIN(J${employeeRow + 1}/10*20,20)+MIN(K${employeeRow + 1}/10*20,20)` });
    }

    //Leave Distribution
    const leaveDays: string[] = ['Casual Leave', 'Sick Leave', 'Loss of Pay'];
    leaveDays.forEach((value: string, index: number) => {
      setCell(14 + index, 0 , dataSheet, { value: value });
    });

    //formulas for leave
    setCell(14, 1 , dataSheet, { formula:`=COUNTIFS(${timeSheetName}!J6:J${timeSheetLastRow},"CL")` });
    setCell(15, 1 , dataSheet, { formula:`=COUNTIFS(${timeSheetName}!J6:J${timeSheetLastRow},"SL")` });
    setCell(16, 1 , dataSheet, { formula:`=COUNTIFS(${timeSheetName}!J6:J${timeSheetLastRow},"LOP")` });
  }

  return (<SpreadsheetComponent
    height='650px'
    openUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/open'
    saveUrl='https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/save'
    ref={(ssObj: Spreadsheet) => { spreadsheet = ssObj }} created={onCreated.bind(this)}
    dataSourceChanged={dataSourceChanged}
  >
  <SheetsDirective>
    <SheetDirective name='Dashboard' showGridLines={false}></SheetDirective>
    <SheetDirective name='EmployeeMaster' showGridLines={false}></SheetDirective>
    <SheetDirective name='Timesheet' showGridLines={false}></SheetDirective>
    <SheetDirective name='Payroll' showGridLines={false}></SheetDirective>
    <SheetDirective name='Data'></SheetDirective>
  </SheetsDirective>
  </SpreadsheetComponent>);
}