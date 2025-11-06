import * as React from "react";
import { useEffect,useState } from "react";
import DataTable from "react-data-table-component";
import Search from "./Search";
import ExportExcel from "./ExportExcel";

// import ExportListItemsToPDF from "./ExportListItemsToPDF";

const customStyles = {
    rows: {
        style: {
          minHeight: '70px', // override the row height
        }
    },
    headCells: {
        style: {
            paddingLeft: '11px', // override the cell padding for head cells
            // paddingRight: '3px',
            color: '#572ba7',
            fontSize: '.9rem',
            background: 'linear-gradient(rgb(228 228 228),rgb(191 191 191))',
            borderTop: '0!important',
            borderBottom: '2px solid #dee2e6;',
            verticalAlign: 'bottom'
        },
    },
    cells: {
        style: {
            paddingLeft: '3px', // override the cell padding for data cells
            paddingRight: '3px',

        },
    }
}

interface TableGeneratorProps {
    columns: any;
    data: any;
    fileName: string;
    showExportExcel?: boolean;
    ExportExcelCustomisedColumns?:any;
    ExportExcelCustomisedData?:any;
    prvPageNumber?: number;
    prvSort?:any;
    prvDirection?:boolean;
    onChange?:any;
    onSortChange?:any;
    onSortDirection?:any;
    showPagination?:boolean;
    onRowClick?:any;
    className?:string;
}

const TableGenerator = ({ columns, data, fileName,showExportExcel, ExportExcelCustomisedColumns,ExportExcelCustomisedData, prvPageNumber,prvSort,prvDirection,onChange,onSortChange,onRowClick, onSortDirection, showPagination, className }: TableGeneratorProps) =>{

    const [totalData, setData] = useState([]);
    const [search, setSearchText] = useState('');

    prvPageNumber = prvPageNumber!= undefined && prvPageNumber!= null?prvPageNumber:1;
    prvSort = (prvSort!= undefined && prvSort!= null)?prvSort:"";
    prvDirection = (prvDirection!= undefined && prvDirection!= null)?prvDirection:false;

    useEffect(() => {
        ExportExcelCustomisedData =(ExportExcelCustomisedData== undefined && ExportExcelCustomisedData!= null)? ExportExcelCustomisedData : data;
        ExportExcelCustomisedColumns=(ExportExcelCustomisedColumns== undefined && ExportExcelCustomisedColumns!= null)? ExportExcelCustomisedColumns : columns;

        let totaldata = data;
        if (search) {
            var allKeys = Object.keys(data[0]);
            totaldata = totaldata.filter((l:any) => allKeys.some(field => {
            return (l[field] && l[field].toString().toLowerCase().includes(search.toLowerCase()));
            }));
            setData(totaldata);
        } else {
            setData(data);
        }
    }, [data, search]);

    return(
      <div className="table-responsive dataTables_wrapper-overflow">
          <div className="form-border-box p-2 mx-1 my-2">
            <div className={showExportExcel ? '' : 'float-right'}>
              
                { data.length > 0 &&
                <Search onSearch={value => {
                  setSearchText(value);
                }} ></Search>
              }
            

              {showExportExcel &&
                <div className="col-6 text-right">
                  <ExportExcel tableData={ExportExcelCustomisedData ? ExportExcelCustomisedData : data} filename={fileName} columns={ExportExcelCustomisedColumns ? ExportExcelCustomisedColumns : columns}></ExportExcel>
                </div> 
              }
            </div>

        {/* <div>
          <ExportListItemsToPDF listName={fileName} columns={columns} data={data}></ExportListItemsToPDF>
        </div> */}

            <div className="mt-2" id="tablePDF">
              <DataTable
                noHeader
                columns={columns}
                data={totalData}
                striped={true}
                pagination={showPagination}
                actions
                customStyles={customStyles}
                paginationDefaultPage={1}
                persistTableHead={true}
                onChangePage={onChange}
                onSort={onSortChange}
                defaultSortFieldId={prvSort}
                defaultSortAsc={prvDirection}
                onRowClicked={onRowClick}
                className={className}
              />
            </div>
          </div>
      </div>
    )

}

export default TableGenerator;