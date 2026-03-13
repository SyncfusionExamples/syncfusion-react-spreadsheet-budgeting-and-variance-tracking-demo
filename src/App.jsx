import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { CellDirective, CellsDirective, ColumnDirective, ColumnsDirective, getColumnHeaderText, getSheet, RangeDirective, RangesDirective, RowDirective, RowsDirective, SheetDirective, SheetsDirective, SpreadsheetComponent,
} from '@syncfusion/ej2-react-spreadsheet';
import { budgetSheet, actualSheet } from './dataSource';

export default function App() {
    let spreadsheet;
    const onCreated = () =>{
    }
    const dataBound = () => {
        spreadsheet.cellFormat({ backgroundColor: '#c1d3d9' }, 'Variance!A1:A40');
        //updating formula for budget sheet and perform auto-fill
        spreadsheet.updateCell({formula:`=SUM(Budget!B2:M2)`},'Budget!N2');
        spreadsheet.autoFill('Budget!N3:N13', 'Budget!N2', 'Down', 'FillSeries');
        //updating formula for actual sheet and perform auto-fill
        spreadsheet.updateCell({formula:`=SUM(Actual!B2:M2)`},'Actual!N2');
        spreadsheet.autoFill('Actual!N3:N13', 'Actual!N2', 'Down', 'FillSeries');
        //variance calculation
        spreadsheet.updateCell({formula:`=Variance!D6-Variance!C6`},'Variance!E6');
        spreadsheet.autoFill('Variance!E7:E17', 'Variance!E6', 'Down', 'FillSeries');
        //variance percentage
        spreadsheet.updateCell({formula:`=Variance!E6/Variance!C6`},'Variance!F6');
        spreadsheet.autoFill('Variance!F7:F17', 'Variance!F6', 'Down', 'FillSeries');
        //applying number format to variance % 
        spreadsheet.numberFormat('0.00%','Variance!F3:F17');
        //Total variance and variance %
        spreadsheet.updateCell({ formula: `=SUM(E6:E17)` }, 'Variance!D3');
        spreadsheet.updateCell({ formula: `=D3/B3` }, 'Variance!E3');
        // Budget Sheet Formatting
        spreadsheet.cellFormat({backgroundColor: '#4472C4', color: '#FFFFFF', fontWeight: 'bold'}, 'Budget!A1:N1');
        spreadsheet.cellFormat({backgroundColor: '#D9E1F2', fontWeight: 'bold'}, 'Budget!A2:A13');
        spreadsheet.cellFormat({backgroundColor: '#4472C4', color: '#FFFFFF', fontWeight: 'bold',}, 'Actual!A1:N1');
        spreadsheet.cellFormat({backgroundColor: '#D9E1F2', fontWeight: 'bold'}, 'Budget!A2:N13');
        spreadsheet.cellFormat({backgroundColor: '#D9E1F2', fontWeight: 'bold'}, 'Actual!A2:N13');
        //Variance Sheet Formatting
        spreadsheet.merge('Variance!B1:F1');
        spreadsheet.cellFormat({ fontWeight: 'bold', backgroundColor: '#b3ffb3' }, 'Variance!B1:F1');
        spreadsheet.cellFormat({ color: 'black', fontSize: '23pt', fontWeight: 'bold', textAlign: 'center' }, 'Variance!B1:D1');
        spreadsheet.cellFormat({ fontWeight: 'bold' }, 'Variance!A6:A17');
        spreadsheet.cellFormat({ verticalAlign: 'middle', textAlign: 'center' }, `Variance!A1:K20`);
        spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '17pt' }, `Variance!A5:F5`);
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!B1:F1', 'Outer');
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!B1:F3', 'Outer');
        spreadsheet.setBorder({ border: '1px solid black' }, 'Variance!C5:F17', 'Vertical');
        spreadsheet.cellFormat({ backgroundColor: '#ccccff',fontWeight:'bold',fontSize:'14pt'}, 'Variance!B2:F3');
        spreadsheet.conditionalFormat({ type: 'ThreeSymbols', range: 'Variance!F6:F17' });
        spreadsheet.conditionalFormat({ type: 'ThreeTriangles', range: 'Variance!E6:E17' });
        spreadsheet.numberFormat('$#,##0.00','Variance!C6:D17');
        spreadsheet.numberFormat('$#,##0.00','Variance!B3:C3');
        spreadsheet.numberFormat('0.00%','Variance!E3');
        spreadsheet.setRowsHeight(30, ['Budget!1:20', 'Actual!1:20']);
        //Budget Sheet Formatting
    };
    const pieChart = [{ type: 'Pie', range: 'Variance!E5:E17', title: 'VARIANCE', theme: 'Bootstrap5', height: 265, width: 555, top: 10, left: 940, id: 'Chart2', isSeriesInRows: false, legendSettings: { visible: false }, dataLabelSettings: { visible: true, position: 'Outer' } }];
    const columnChart = [{ type: 'Column', range: 'Variance!B5:D17', title: 'PLANNED BUDGET VS ACTUAL SPENT', theme: 'Bootstrap5', height: 265, width: 555, top: 300, left: 940, id: 'Chart1', isSeriesInRows: true, dataLabelSettings: { visible: true, position: 'Outer' } }];
    return (<SpreadsheetComponent height={780} width={1550}
        ref={(ssObj) => { spreadsheet = ssObj }} dataBound={dataBound} created={onCreated.bind(this)}>
        <SheetsDirective>
            <SheetDirective name='Variance' showGridLines={false}>
            <RowsDirective>
                <RowDirective height={40}>
                    <CellsDirective>
                        <CellDirective index={1} value='BUDGET AND VARIANCE TRACKING'></CellDirective>
                        <CellDirective index={2} chart={pieChart}></CellDirective>
                    </CellsDirective>
                </RowDirective>
                <RowDirective height={30}>
                    <CellsDirective>
                        <CellDirective index={1} value='TOTAL BUDGET'></CellDirective>
                        <CellDirective index={2} value='TOTAL ACTUAL'></CellDirective>
                        <CellDirective index={3} value='TOTAL VARIANCE'></CellDirective>
                        <CellDirective index={4} value='VARIANCE %'></CellDirective>
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
                        <CellDirective index={2} value='BUDGET PLANNED'></CellDirective>
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
                        <CellDirective index={1} value='OFFICE SUPPLIES' chart={columnChart}></CellDirective>
                        <CellDirective index={2} formula='=Budget!N13'></CellDirective>
                        <CellDirective index={3} formula='=Actual!N13'></CellDirective>
                    </CellsDirective>
                </RowDirective>
            </RowsDirective>
            <ColumnsDirective>
            <ColumnDirective index= {0} width={20}></ColumnDirective>
            <ColumnDirective index= {1} width={180}></ColumnDirective>
            <ColumnDirective index= {2} width={190}></ColumnDirective>
            <ColumnDirective index= {3} width={180}></ColumnDirective>
            <ColumnDirective index= {4} width={180}></ColumnDirective>
            <ColumnDirective index= {5} width={180}></ColumnDirective>
            <ColumnDirective index= {6} width={120}></ColumnDirective>
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
                </ColumnsDirective>
            </SheetDirective>
        </SheetsDirective>
    </SpreadsheetComponent>)
}