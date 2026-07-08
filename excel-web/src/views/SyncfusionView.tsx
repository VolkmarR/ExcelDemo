import { SpreadsheetComponent } from '@syncfusion/ej2-react-spreadsheet';

export default function SyncfusionView() {
    return (<SpreadsheetComponent openUrl='http://localhost:5269/api/spreadsheet/open'
                                  allowFiltering={true}
                                  allowEditing={false}
                                  allowInsert={false}
                                  allowDelete={false} />);
}