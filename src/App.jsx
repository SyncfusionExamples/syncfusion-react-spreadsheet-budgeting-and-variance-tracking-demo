import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { CellDirective, CellsDirective, ColumnDirective, ColumnsDirective, getColumnHeaderText, getSheet, RangeDirective, RangesDirective, RowDirective, RowsDirective, SheetDirective, SheetsDirective, SpreadsheetComponent,
} from '@syncfusion/ej2-react-spreadsheet';
import { budgetSheet, actualSheet } from './dataSource';
export default function App() {
    let spreadsheet;
    const dataBound = () => {
        spreadsheet.cellFormat({ backgroundColor: '#c1d3d9' }, 'Variance!A1:A40');
        //updating formula for budget sheet and perform auto-fill
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'left' }, 'Variance!B7:B22');
        spreadsheet.updateCell({formula:`=SUM(Budget!B2:M2)`},'Budget!N2');
        spreadsheet.autoFill('Budget!N3:N13', 'Budget!N2', 'Down', 'FillWithoutFormatting');
        //updating formula for actual sheet and perform auto-fill
        spreadsheet.updateCell({formula:`=SUM(Actual!B2:M2)`},'Actual!N2');
        spreadsheet.autoFill('Actual!N3:N13', 'Actual!N2', 'Down', 'FillWithoutFormatting');
        //variance calculation
        spreadsheet.updateCell({formula:`=Variance!D7-Variance!C7`},'Variance!E7');
        spreadsheet.autoFill('Variance!E8:E18', 'Variance!E7', 'Down', 'FillSeries');
        //variance percentage
        spreadsheet.updateCell({formula:`=Variance!E7/Variance!C7`},'Variance!F7');
        spreadsheet.autoFill('Variance!F8:F18', 'Variance!F7', 'Down', 'FillSeries');
        //applying number format to variance % 
        spreadsheet.numberFormat('0.00%','Variance!F3:F22');
        //Total variance and variance %
        spreadsheet.updateCell({ formula: `=SUM(E7:E18)` }, 'Variance!E4');
        spreadsheet.updateCell({ formula: `=E4/B4` }, 'Variance!F4');
        // Budget Sheet Formatting
        spreadsheet.cellFormat({backgroundColor: '#4472C4', color: '#FFFFFF', fontWeight: 'bold',textAlign:'center',verticalAlign:'middle',fontSize:'16pt'}, 'Budget!A1:N1');
        spreadsheet.cellFormat({backgroundColor: '#4472C4', color: '#FFFFFF', fontWeight: 'bold',textAlign:'center',verticalAlign:'middle',fontSize:'16pt'}, 'Actual!A1:N1');
        spreadsheet.cellFormat({backgroundColor: '#D9E1F2'}, 'Budget!A2:N2 A4:N4 A6:N6 A8:N8 A10:N10 A12:N12');
        spreadsheet.cellFormat({backgroundColor: '#D9E1F2'}, 'Actual!A2:N2 A4:N4 A6:N6 A8:N8 A10:N10 A12:N12');
        spreadsheet.cellFormat({backgroundColor: '#ffcccc'}, 'Actual!A3:N3 A5:N5 A7:N7 A9:N9 A11:N11 A13:N13');
        spreadsheet.cellFormat({backgroundColor: '#ffcccc'}, 'Budget!A3:N3 A5:N5 A7:N7 A9:N9 A11:N11 A13:N13');
        spreadsheet.cellFormat({backgroundColor: '#4472C4', color: '#FFFFFF', fontWeight: 'bold',}, 'Actual!A1:N1');
        //Variance Sheet Formatting
        spreadsheet.merge('Variance!B1:F1');
        spreadsheet.cellFormat({ fontWeight: 'bold', backgroundColor: '#b3ffb3' }, 'Variance!B1:F1');
        spreadsheet.cellFormat({ color: 'black', fontSize: '23pt', fontWeight: 'bold', textAlign: 'center' }, 'Variance!B1:D1');
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center' }, `Variance!A1:K20`);
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '15pt' }, `Variance!A6:F6`);
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!B1:F1', 'Outer');
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!B3:C4 E3:F4', 'Inner');
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!B3:C4 E3:F4', 'Outer');
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!C6:F19', 'Vertical');
        spreadsheet.setBorder({ borderBottom: '1px solid black' }, 'Variance!B6:F19', 'Horizontal');
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!B6:F6', 'Horizontal');
        spreadsheet.cellFormat({ backgroundColor: '#ccccff',fontWeight:'bold',fontSize:'14pt'}, 'Variance!B3:C4 E3:F4');
        spreadsheet.conditionalFormat({ type: 'ThreeSymbols', range: `Variance!F6:F22` });
        spreadsheet.conditionalFormat({ type: 'ThreeTriangles', range: `Variance!E6:E22` });
        spreadsheet.numberFormat('$#,##0.00','Variance!C7:E22');
        spreadsheet.numberFormat('$#,##0.00','Variance!B4:E4');
         spreadsheet.numberFormat('$#,##0.00','Budget!B2:N13');
        spreadsheet.numberFormat('$#,##0.00','Actual!B2:N13');
        spreadsheet.numberFormat('0.00%','Variance!F4');
        spreadsheet.setRowsHeight(40, ['Budget!1', 'Actual!1']);
        spreadsheet.setRowsHeight(30, ['Budget!2:22', 'Actual!2:22']);
    };
    const pieChart = [{ type: 'Pie', range: 'Variance!E7:E18', title: 'VARIANCE', theme: 'Bootstrap5', height: 290, width: 720, top: 300, left: 773, id: 'Chart2', isSeriesInRows: false, legendSettings: { visible: false } }];
    const columnChart = [{ type: 'Column', range: 'Variance!B6:D18', title: 'PLANNED BUDGET VS ACTUAL SPENT', theme: 'Bootstrap5',height: 290, width: 720, top: 2, left: 773, id: 'Chart1', isSeriesInRows: true }];
    return (<SpreadsheetComponent height={780} width={1550}
        ref={(ssObj) => { spreadsheet = ssObj }} dataBound={dataBound}>
        <SheetsDirective>
            <SheetDirective name='Variance' showGridLines={false}>
            <RowsDirective>
                <RowDirective height={40}>
                    <CellsDirective>
                        <CellDirective index={1} value='BUDGET AND VARIANCE TRACKING'></CellDirective>
                        <CellDirective index={2} chart={columnChart}></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={30}></RowDirective>
                <RowDirective height={30}>
                    <CellsDirective>
                        <CellDirective index={1} value='TOTAL BUDGET'></CellDirective>
                        <CellDirective index={2} value='TOTAL ACTUAL'></CellDirective>
                        <CellDirective index={4} value='TOTAL VARIANCE'></CellDirective>
                        <CellDirective index={5} value='VARIANCE %'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={30}>
                    <CellsDirective>
                        <CellDirective index={1} formula='=SUM(C6:C17)'></CellDirective>
                        <CellDirective index={2} formula='=SUM(D6:D17)'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective></RowDirective>
                <RowDirective height={35}>
                    <CellsDirective>
                        <CellDirective index={1} value='CATEGORY'></CellDirective>
                        <CellDirective index={2} value='BUDGET'></CellDirective>
                        <CellDirective index={3} value='ACTUALS'></CellDirective>
                        <CellDirective index={4} value='VARIANCE'></CellDirective>
                        <CellDirective index={5} value='VARIANCE %'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='MARKETING'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N2'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N2'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='DIGITAL ADS'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N3'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N3'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='EVENTS'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N4'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N4'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='SALARIES'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N5'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N5'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='HIRING COSTS'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N6'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N6'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='TRAINING COSTS'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N7'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N7'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='SOFTWARE LICENSES'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N8'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N8'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='CLOUD HOSTING'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N9'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N9'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='RENT'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N10'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N10'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='UTILITIES'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N11'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N11'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='INSURANCE'></CellDirective>
                        <CellDirective index={2} formula='=Budget!N12'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N12'></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={25}>
                    <CellsDirective>
                        <CellDirective index={1} value='OFFICE SUPPLIES' chart={pieChart}></CellDirective>
                        <CellDirective index={2} formula='=Budget!N13'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N13'></CellDirective>
                    </CellsDirective>
                </RowDirective>
            </RowsDirective>
            <ColumnsDirective>
            <ColumnDirective index= {0} width={20}></ColumnDirective>
            <ColumnDirective index= {1} width={150}></ColumnDirective>
            <ColumnDirective index= {2} width={150}></ColumnDirective>
            <ColumnDirective index= {3} width={150}></ColumnDirective>
            <ColumnDirective index= {4} width={150}></ColumnDirective>
            <ColumnDirective index= {5} width={150}></ColumnDirective>
            <ColumnDirective index= {6} width={150}></ColumnDirective>
            </ColumnsDirective>
            </SheetDirective>
            <SheetDirective name='Budget' showGridLines={false}>
                <RangesDirective>
                    <RangeDirective dataSource={budgetSheet}> </RangeDirective>
                </RangesDirective>
                <RowsDirective>
                    <RowDirective index={0}>
                        <CellsDirective>
                            <CellDirective value='Total' index={13}></CellDirective>
                        </CellsDirective>
                    </RowDirective>
                </RowsDirective>
                <ColumnsDirective>
                <ColumnDirective width={180}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                </ColumnsDirective>
            </SheetDirective>
             <SheetDirective name='Actual' showGridLines={false}>
                <RangesDirective>
                    <RangeDirective dataSource={actualSheet}> </RangeDirective>
                </RangesDirective>
                <RowsDirective>
                    <RowDirective index={0}>
                        <CellsDirective>
                            <CellDirective value='Total' index={13}></CellDirective>
                        </CellsDirective>
                    </RowDirective>
                </RowsDirective>
                <ColumnsDirective>
                <ColumnDirective width={180}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                <ColumnDirective width={100}></ColumnDirective>
                </ColumnsDirective>
            </SheetDirective>
        </SheetsDirective>
    </SpreadsheetComponent>)
}