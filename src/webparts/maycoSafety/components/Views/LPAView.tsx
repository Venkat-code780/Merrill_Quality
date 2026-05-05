import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
// import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { NavLink } from "react-router-dom";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit } from "@fortawesome/free-solid-svg-icons";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import { useNavigate } from "react-router-dom";
import AGGridDataTable from "../Shared/AGGridDataTable";
// import { spfi, SPFx } from "@pnp/sp";

export interface LPAFormProps {
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
}

const LPAView: React.FC<LPAFormProps> = (props) => {
  const currentSiteURL = props.spContext.webAbsoluteUrl;
  const { getListItems } = initCommonFunctions(props.context, props.siteURL);
  const listName = "LPA";
    const navigate = useNavigate();
  // const sp = spfi().using(SPFx(props.context));

  const [data, setData] = useState<any[]>([]);
  // const [pageNumber, setPageNumber] = useState(1);

  useEffect(() => {
    highlightCurrentNav("liLPA");
    loadData();
  }, []);

  const loadData = async () => {

    try {
      showLoader();
        const items= await getListItems(listName,currentSiteURL,"*,Auditor/Title","Auditor","")
    //   const items = await sp.web.lists
    //     .getByTitle(listName)
    //     .items
    //     .select(
    //       "LPA_x0020_Category",
    //       "LPA_x0020_Subcategory",
    //        "Plant",
    //       "Department",
    //       "Zone",
    //       "Status",
    //       "Remarks",
    //       "Machine",
    //       "Date",
    //       "Auditor/Title",
    //     )
    //     .expand("Auditor")
    //     .orderBy("Modified", false)
    //     .top(2000)();
    const sortedItems = items.sort(
  (a: any, b: any) =>
    new Date(b.Modified).getTime() - new Date(a.Modified).getTime()
);

      const tableData = sortedItems.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Date: item.Date
          ? new Date(item.Date).toLocaleDateString()
          : "",
        Department: item.Department,
        Zone_x0009_: item.Zone_x0009_,
         Machine: item.Machine,
        Tool_x0020_Number:item.Tool_x0020_Number,
        Operators:item.Operators,
        Supervisor:item.Supervisor,
        Comments:item.Comments,
        Year:item.Year,
        Auditor: item.Auditor?.Title || "",
        Remarks: item.Remarks,
      }));

      setData(tableData);

    } catch (e) {
      console.error(e);
    } finally {
      hideLoader();
    }
  };




// const loadData = async () => {
//   try {
//     showLoader();
//     let allItems: any[] = [];
    
//     // 1. Setup the initial URL with all your selects and expands
//     // Using the internal spHttpClient logic via PnPjs for maximum control
//     let nextUrl = `${props.siteURL}/_api/web/lists/getbytitle('${listName}')/items?` + 
//                   `$select=Id,Title,Date,Department,Zone_x0009_,Machine,Tool_x0020_Number,Operators,Supervisor,Comments,Year,Remarks,Modified,Auditor/Title` +
//                   `&$expand=Auditor` +
//                   `&$top=5000`; // Max chunk size

//     // 2. The manual loop
//     while (nextUrl) {
//       // Fetch the current page
//       const response = await props.spHttpClient.get(nextUrl, SPHttpClient.configurations.v1);
//       const jsonResponse = await response.json();

//       if (jsonResponse.value) {
//         // Push the results into our master array
//         allItems = [...allItems, ...jsonResponse.value];
//         console.log(`Current count: ${allItems.length}`);
//       }

//       // Check if SharePoint provided a next link for more data
//       // SharePoint returns this as "odata.nextLink" or "@odata.nextLink"
//       nextUrl = jsonResponse["odata.nextLink"] || jsonResponse["@odata.nextLink"] || null;
//     }

//     // 3. Process the data for AG-Grid
//     const tableData = allItems.map((item: any) => ({
//       Id: item.Id,
//       Title: item.Title,
//       Date: item.Date ? new Date(item.Date).toLocaleDateString() : "",
//       Department: item.Department,
//       Zone_x0009_: item.Zone_x0009_,
//       Machine: item.Machine,
//       Tool_x0020_Number: item.Tool_x0020_Number,
//       Operators: item.Operators,
//       Supervisor: item.Supervisor,
//       Comments: item.Comments,
//       Year: item.Year,
//       Auditor: item.Auditor?.Title || "",
//       Remarks: item.Remarks,
//       Modified: item.Modified
//     }));

//     // 4. Client-side Sort (Required because SP API sorting breaks over 5k)
//     tableData.sort((a, b) => new Date(b.Modified).getTime() - new Date(a.Modified).getTime());

//     setData(tableData);

//   } catch (e) {
//     console.error("Error loading massive dataset:", e);
//   } finally {
//     hideLoader();
//   }
// };














  const handleRowClicked = (row: any) => {
  console.log("Row clicked:", row);

  // navigate to edit page
  navigate(`/LPAForm/${row.Id}`);
};

  /* ---------------- TABLE COLUMNS ---------------- */

  // const columns = [
  //    {
  //   name: "Edit",
  //   selector: (row: any) => row.Id,
  //   cell: (row: any) => (
  //     <NavLink to={`/LPAForm/${row.Id}`} title="Edit">
  //       <FontAwesomeIcon icon={faEdit} />
  //     </NavLink>
  //   ),
  //   width: "70px"
  // },

  //   {
  //     name: "Plant",
  //     selector: (row: any) => row.Title,
  //     sortable: true
  //   },
  //    {
  //     name: "Date",
  //     selector: (row: any) => row.Date,
  //     sortable: true
  //   },
  //     {
  //     name: "Department",
  //     selector: (row: any) => row.Department,
  //     sortable: true
  //   },
  //   {
  //     name: "Zone",
  //     selector: (row: any) => row.Zone_x0009_,
  //     sortable: false
  //   },
  //   {
  //     name: "Machine",
  //     selector: (row: any) => row.Machine,
  //     sortable: true
  //   },
  //   {
  //     name: "Tool Number",
  //     selector: (row: any) => row.Tool_x0020_Number,
  //     sortable: true
  //   },
  //   {
  //     name: "Operators",
  //     selector: (row: any) => row.Operators,
  //     sortable: true
  //   },
  //    {
  //     name: "Supervisor's Name",
  //     selector: (row: any) => row.Supervisor,
  //     sortable: true
  //   },
  //     {
  //     name: "Comments",
  //     selector: (row: any) => row.Comments,
  //     sortable: true
  //   },
  //      {
  //     name: "Year",
  //     selector: (row: any) => row.Year,
  //     sortable: true
  //   },
  //      {
  //     name: "Auditor's Name",
  //     selector: (row: any) => row.Auditor,
  //     sortable: true
  //   },

  

  // ];


  const columns = [
  {
    headerName: "Edit",
    field: "Id",
    width: 70,
    sortable: false,
    filter: false,
    cellRenderer: (params: any) => {
      const row = params.data;

      return (
        <NavLink to={`/LPAForm/${row.Id}`} title="Edit">
          <FontAwesomeIcon icon={faEdit} />
        </NavLink>
      );
    },
  },

  {
    headerName: "Plant",
    field: "Title",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Date",
    field: "Date",
    sortable: true,
    filter: "agDateColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Department",
    field: "Department",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Zone",
    field: "Zone_x0009_",
    sortable: false,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Machine",
    field: "Machine",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Tool Number",
    field: "Tool_x0020_Number",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Operators",
    field: "Operators",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Supervisor's Name",
    field: "Supervisor",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Comments",
    field: "Comments",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Year",
    field: "Year",
    sortable: true,
    filter: "agNumberColumnFilter",
    resizable: true,
    flex: 1,
  },

  {
    headerName: "Auditor's Name",
    field: "Auditor",
    sortable: true,
    filter: "agTextColumnFilter",
    resizable: true,
    flex: 1,
  },
];
  return (

    <div className="container-fluid">

      <div className="light-box border-box-shadow">

        <div className="div-form-title">
          <div className="form-title">LPA Header View</div>
        </div>



        <div className="p-2 mx-1 ViewTable">
           <AGGridDataTable
  data={data}
  columns={columns}
  showExportExcel={false}
  showAddButton={false}
  customBtnClass="px-1 text-right"
  btnDivID=""
  btnSpanID=""
  btnTitle=""
  searchBoxLeft={true}
  onRowClicked={(event: any) => handleRowClicked(event.data)}
  domLayout="normal"
  suppressColumnVirtualization={true}
  ensureDomOrder={true}
  pagination={true}
  suppressHorizontalScroll={false}
  suppressSizeToFit={true}
  suppressColumnHiding={true}
  suppressAutoSize={true}
  suppressColumnMoveAnimation={true}
  suppressMovableColumns={true}
/>
          {/* <TableGenerator
            columns={columns}
            data={data}
            onChange={setPageNumber}
            onRowClick={handleRowClicked}
            prvPageNumber={pageNumber}
            prvDirection={false}
            fileName={"LPA"}
            className="sp-Datatable-hh"
            showPagination={true}
          /> */}

        </div>

      </div>

    </div>

  );
};

export default LPAView;

