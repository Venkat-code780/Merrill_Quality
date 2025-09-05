import * as React from "react";
import * as XLSX from "xlsx";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faFileExcel } from '@fortawesome/free-solid-svg-icons';

interface ExportExcelInterface {
    tableData:any;
    columns:any;
    filename:any;
}

const ExportExcel: React.FC<ExportExcelInterface> = ({ tableData,columns, filename }) => {
  
    const exportExcelFile = (workbook:any) => {
        return XLSX.writeFile(workbook, `${filename}.xlsx`);
    };
    const exportToexcel = (dataTable:any,) => {
        let newjson: any[]=[];
        dataTable.map((selItem:any, index:number) => {
            // var obj: {} ={};
            let obj: { [key: string]: any } = {};
            columns.map((column:any,i:number) => {
                let name =column.name;
                if(name != 'Edit' && name !="View")
                {
                let seletor =column.selector;
                obj[name] =selItem[seletor];
                }
            });
            newjson.push(obj);
        });

        var workbook = XLSX.utils.book_new();
        var worksheet_data = XLSX.utils.json_to_sheet(newjson);

        workbook.SheetNames.push(filename);
        workbook.Sheets[filename] = worksheet_data;

        exportExcelFile(workbook);
    };
    
    return (
        <a type="button" id="btnDownloadFile" className="icon-export-b" title="Export to Excel" onClick={(e) => exportToexcel(tableData)}>
            <FontAwesomeIcon icon={faFileExcel} className='icon-export-b'></FontAwesomeIcon>
        </a>
    );
};

export default ExportExcel;