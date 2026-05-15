
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
// import { initCommonFunctions } from "../Utilities/CommonFunctions";
import { useNavigate } from "react-router-dom";
import AGGridDataTable from "../Shared/AGGridDataTable";
import SearchableDropdown from "../Shared/Dropdown";
// import { showToast } from "../Shared/Toaster";
// import { spfi, SPFx } from "@pnp/sp";

export interface LPAFormProps {
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
}

const LPAView: React.FC<LPAFormProps> = (props) => {
  const currentSiteURL = props.spContext.webAbsoluteUrl;
  // const { getListItems } = initCommonFunctions(props.context, props.siteURL);
  const listName = "LPA";
  const navigate = useNavigate();
  // const sp = spfi().using(SPFx(props.context));

  const [data, setData] = useState<any[]>([]);
  const currentYear = String(new Date().getFullYear());
  const [filters, setFilters] = useState({
    year: currentYear,
    month: "All"
  });

  const uniqueYears = Array.from(
    new Set(data.map(x => x.Year).filter(Boolean))
  );


  const yearOptions = uniqueYears.map(y => ({
    label: y,
    value: y
  }));
  //   const uniqueMonths = Array.from(
  //   new Set(data.map(x => x.YearMonth).filter(Boolean))
  // );

  // const monthOptions = uniqueMonths.map(m => ({
  //   label: m,
  //   value: m
  // }));

  const uniqueMonths = Array.from(
    new Set(data.map(x => x.YearMonth).filter(Boolean))
  )
    .map(m => String(m).padStart(2, "0"))
    .sort((a, b) => Number(a) - Number(b));

  // ✅ ADD "All" dynamically at the TOP
  const monthOptions = [
    { label: "All", value: "All" },
    ...uniqueMonths.map(m => ({
      label: m,
      value: m
    }))
  ];
  // const [pageNumber, setPageNumber] = useState(1);

  useEffect(() => {
    highlightCurrentNav("liLPA");
    loadData();
  }, []);





  // const loadData = async () => {

  //   try {
  //     showLoader();
  //     const items = await getListItems(listName, currentSiteURL, "*,Auditor/Title", "Auditor", "")

  //     const sortedItems = items.sort(
  //       (a: any, b: any) =>
  //         new Date(b.Modified).getTime() - new Date(a.Modified).getTime()
  //     );

  //     const tableData = sortedItems.map((item: any) => ({
  //       Id: item.Id,
  //       Title: item.Title,
  //       Date: item.Date
  //         ? new Date(item.Date).toLocaleDateString("en-US")
  //         : "",
  //       Department: item.Department,
  //       Zone_x0009_: item.Zone_x0009_,
  //       Machine: item.Machine,
  //       Tool_x0020_Number: item.Tool_x0020_Number,
  //       Operators: item.Operators,
  //       Supervisor: item.Supervisor,
  //       Comments: item.Comments,
  //       Year: item.Year,
  //       YearMonth: item.YearMonth ? item.YearMonth.toString().padStart(2, "0") : "",
  //       Auditor: item.Auditor?.Title || "",
  //       Remarks: item.Remarks,
  //     }));


  //     setData(tableData);

  //   } catch (e) {
  //     console.error(e);
  //   } finally {
  //     hideLoader();
  //   }
  // };

  const loadData = async () => {
    try {
      showLoader();

      const url =
        `${currentSiteURL}/_api/web/lists/getbytitle('${listName}')/items` +
        `?$select=Id,Title,Date,Department,Zone_x0009_,Machine,Tool_x0020_Number,Operators,Supervisor,Comments,Year,YearMonth,Modified,Auditor/Title` +
        `&$expand=Auditor` +
        `&$orderby=Modified desc` +
        `&$top=2000`;

      const response = await props.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      const data = await response.json();

      console.log("SharePoint response:", data);

      const tableData = (data.value || []).map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Date: item.Date
          ? new Date(item.Date).toLocaleDateString("en-US")
          : "",
        Department: item.Department,
        Zone_x0009_: item.Zone_x0009_,
        Machine: item.Machine,
        Tool_x0020_Number: item.Tool_x0020_Number,
        Operators: item.Operators,
        Supervisor: item.Supervisor,
        Comments: item.Comments,
        Year: item.Year,
        YearMonth: item.YearMonth
          ? item.YearMonth.toString().padStart(2, "0")
          : "",
        Auditor: item.Auditor?.Title || "",
        Remarks: item.Remarks,
        Modified: item.Modified
      }));

      setData(tableData);
    } catch (e) {
      console.error("LoadData Error:", e);
    } finally {
      hideLoader();
    }
  };



  const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
    const name = actionMeta?.name;

    if (!name) return;

    setFilters(prev => ({
      ...prev,
      [name]: selectedOption ? selectedOption.value : null
    }));
  };


  // const filteredData = React.useMemo(() => {
  //   if ((filters.year && !filters.month) || (!filters.year && filters.month)) {
  //     showToast("error", "Please select both Year and Month");
  //     return data;
  //   }
  //   if (!filters.year || !filters.month) return data;

  //   return data.filter((item) => {
  //     return (
  //       String(item.Year) === String(filters.year) &&
  //       String(item.YearMonth) === String(filters.month)
  //     );
  //   });
  // }, [data, filters]);


  const filteredData = React.useMemo(() => {
    let result = data;

    if (filters.year) {
      result = result.filter(item =>
        String(item.Year) === String(filters.year)
      );
    }

    if (filters.month && filters.month !== "All") {
      result = result.filter(item =>
        String(item.YearMonth) === String(filters.month)
      );
    }

     return result.sort(
    (a: any, b: any) =>
      new Date(b.Modified).getTime() -
      new Date(a.Modified).getTime()
  );
  }, [data, filters]);



  const handleRowClicked = (row: any) => {
    console.log("Row clicked:", row);

    // navigate to edit page
    navigate(`/LPAForm/${row.Id}`);
  };






  /* ---------------- TABLE COLUMNS ---------------- */



  const columns = [
    {
      headerName: "Edit",
      field: "Id",
      width: 80,
      minWidth: 80,
      maxWidth: 80,
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
      minWidth: 160
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
      minWidth: 140,
    },

    {
      headerName: "Tool Number",
      field: "Tool_x0020_Number",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
      minWidth: 140,
    },

    {
      headerName: "Operators",
      field: "Operators",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
      minWidth: 140,
    },

    {
      headerName: "Supervisor's Name",
      field: "Supervisor",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
      minWidth: 160,
    },

    {
      headerName: "Comments",
      field: "Comments",
      sortable: true,
      filter: "agTextColumnFilter",
      resizable: true,
      flex: 1,
      minWidth: 150,
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
      headerName: "Month",
      field: "YearMonth",
      sortable: true,
      filter: "agTextColumnFilter",
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
      minWidth: 180,
    },
  ];
  return (

    <div className="container-fluid">

      <div className="light-box border-box-shadow">

        <div className="div-form-title">
          <div className="form-title">LPA</div>
        </div>

        <div className="row p-1">
          <div className="col-md-3">
            <div className="light-text">
              <SearchableDropdown
                label="Year"
                name="year"
                selectedValue={filters.year}
                OptionsList={yearOptions}
                OnChange={handleChangeDropdown}
                isRequired={false}
                Title="Year"
                id="ddlYear"
                className=""
                disabled={false}
              />
            </div>
          </div>
          <div className="col-md-3">
            <div className="light-text">
              <SearchableDropdown
                label="Month"
                name="month"
                selectedValue={filters.month}
                OptionsList={monthOptions}
                OnChange={handleChangeDropdown}
                isRequired={false}
                Title="Month"
                id="ddlMonth"
                className=""
                disabled={false}
              />
            </div>
          </div>
        </div>

        <div className="p-2 mx-1 ViewTable">
          <AGGridDataTable
            data={filteredData}
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



// import * as React from "react";
// import { useState, useEffect, useMemo } from "react";
// import { SPHttpClient } from "@microsoft/sp-http";

// import { hideLoader, showLoader } from "../Shared/Loader";
// import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
// import { NavLink } from "react-router-dom";
// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faEdit } from "@fortawesome/free-solid-svg-icons";
// import { useNavigate } from "react-router-dom";

// import AGGridDataTable from "../Shared/AGGridDataTable";
// import SearchableDropdown from "../Shared/Dropdown";

// export interface LPAFormProps {
//   spHttpClient: SPHttpClient;
//   context: any;
//   siteURL: string;
//   spContext: any;
// }

// const LPAView: React.FC<LPAFormProps> = (props) => {
//   const currentSiteURL = props.spContext.webAbsoluteUrl;
//   const listName = "LPA";
//   const navigate = useNavigate();

//   const [data, setData] = useState<any[]>([]);
//   const [filters, setFilters] = useState({
//     year: null as any,
//     month: null as any
//   });

//   // -----------------------------
//   // NAV HIGHLIGHT
//   // -----------------------------
//   useEffect(() => {
//     highlightCurrentNav("liLPA");
//     loadData();
//   }, []);

//   // -----------------------------
//   // 🔥 40K+ SAFE LOADER (PAGING)
//   // -----------------------------
//   const loadData = async () => {
//     try {
//       showLoader();

//       let allItems: any[] = [];

//       let nextUrl: string | null =
//         `${currentSiteURL}/_api/web/lists/getbytitle('${listName}')/items` +
//         `?$select=Id,Title,Year,YearMonth,Date,Department,Zone_x0009_,Machine,Tool_x0020_Number,Operators,Supervisor,Comments,Auditor/Title,Modified` +
//         `&$expand=Auditor` +
//         `&$orderby=Modified desc` +
//         `&$top=5000`;

//       while (nextUrl) {
//         const res = await props.spHttpClient.get(
//           nextUrl,
//           SPHttpClient.configurations.v1
//         );

//         const json = await res.json();

//         if (json.value?.length) {
//           allItems = [...allItems, ...json.value];
//         }

//         nextUrl =
//           json["@odata.nextLink"] || json["odata.nextLink"] || null;
//       }

//       // -----------------------------
//       // TRANSFORM FOR GRID
//       // -----------------------------
//       const tableData = allItems.map((item: any) => ({
//         Id: item.Id,
//         Title: item.Title,
//         Date: item.Date
//           ? new Date(item.Date).toLocaleDateString("en-US")
//           : "",
//         Department: item.Department,
//         Zone_x0009_: item.Zone_x0009_,
//         Machine: item.Machine,
//         Tool_x0020_Number: item.Tool_x0020_Number,
//         Operators: item.Operators,
//         Supervisor: item.Supervisor,
//         Comments: item.Comments,
//         Year: item.Year,
//         YearMonth: item.YearMonth
//           ? item.YearMonth.toString().padStart(2, "0")
//           : "",
//         Auditor: item.Auditor?.Title || "",
//         Modified: item.Modified
//       }));

//       // optional sort (latest first)
//       tableData.sort(
//         (a, b) =>
//           new Date(b.Modified).getTime() -
//           new Date(a.Modified).getTime()
//       );

//       setData(tableData);
//     } catch (e) {
//       console.error("Load error:", e);
//     } finally {
//       hideLoader();
//     }
//   };






//   // -----------------------------
//   // FILTER HANDLER
//   // -----------------------------
//   const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
//     const name = actionMeta?.name;
//     if (!name) return;

//     setFilters((prev) => ({
//       ...prev,
//       [name]: selectedOption ? selectedOption.value : null
//     }));
//   };

//   // -----------------------------
//   // FILTERED DATA
//   // -----------------------------
//   const filteredData = useMemo(() => {
//     if ((filters.year && !filters.month) || (!filters.year && filters.month)) {
//       return data;
//     }

//     if (!filters.year || !filters.month) return data;

//     return data.filter(
//       (item) =>
//         String(item.Year) === String(filters.year) &&
//         String(item.YearMonth) === String(filters.month)
//     );
//   }, [data, filters]);

//   // -----------------------------
//   // ROW CLICK
//   // -----------------------------
//   const handleRowClicked = (row: any) => {
//     navigate(`/LPAForm/${row.Id}`);
//   };

//   // -----------------------------
//   // GRID COLUMNS
//   // -----------------------------
//   const columns = [
//     {
//       headerName: "Edit",
//       field: "Id",
//       width: 70,
//       cellRenderer: (params: any) => (
//         <NavLink to={`/LPAForm/${params.data.Id}`}>
//           <FontAwesomeIcon icon={faEdit} />
//         </NavLink>
//       )
//     },
//     { headerName: "Plant", field: "Title" },
//     { headerName: "Date", field: "Date" },
//     { headerName: "Department", field: "Department" },
//     { headerName: "Machine", field: "Machine" },
//     { headerName: "Year", field: "Year" },
//     { headerName: "Month", field: "YearMonth" },
//     { headerName: "Auditor", field: "Auditor" }
//   ];

//   // -----------------------------
//   // UI
//   // -----------------------------
//   return (
//     <div className="container-fluid">

//       <div className="light-box border-box-shadow">

//         <div className="form-title">
//           LPA Header View (40K+ Optimized)
//         </div>

//         {/* FILTERS */}
//         <div className="row mb-2">
//           <div className="col-md-3">
//             <SearchableDropdown
//               name="year"
//               selectedValue={filters.year}
//               OptionsList={[]}
//               OnChange={handleChangeDropdown}
//               label="Year"
//               Title=""
//               id="year"
//               className=""
//               isRequired={false}
//               disabled={false}
//             />
//           </div>

//           <div className="col-md-3">
//             <SearchableDropdown
//               name="month"
//               selectedValue={filters.month}
//               OptionsList={[]}
//               OnChange={handleChangeDropdown}
//               label="Month"
//               Title=""
//               id="month"
//               className=""
//               isRequired={false}
//               disabled={false}
//             />
//           </div>
//         </div>

//         {/* GRID */}
//         <div className="p-2">
//           <AGGridDataTable
//             data={filteredData}
//             columns={columns}
//             pagination={false}
//             searchBoxLeft={true}
//             onRowClicked={(e: any) =>
//               handleRowClicked(e.data)
//             }
//           />
//         </div>

//       </div>
//     </div>
//   );
// };

// export default LPAView;