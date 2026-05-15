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
import SearchableDropdown from "../Shared/Dropdown";
export interface QAMatrixViewProps {
  spHttpClient: SPHttpClient;
  context: any;
  siteURL: string;
  spContext: any;
}

const QAMatrixView: React.FC<QAMatrixViewProps> = (props) => {
  const currentSiteURL = props.spContext.webAbsoluteUrl;
  const { getListItems } = initCommonFunctions(props.context, props.siteURL);
  const listName = "QA Matrix";
  const navigate = useNavigate();

  //   const sp = spfi().using(SPFx(props.context));

  const [data, setData] = useState<any[]>([]);
  //  const currentYear = String(new Date().getFullYear());
     const [filters, setFilters] = useState({
       year: "All",
     });
       const uniqueYears = Array.from(
    new Set(data.map(x => x.Year).filter(Boolean))
  );
  //   const yearOptions = uniqueYears.map(y => ({
  //   label: y,
  //   value: y
  // }));

  const yearOptions = [
  { label: "All", value: "All" },
  ...uniqueYears.map(y => ({
    label: y,
    value: y
  }))
];
  // const [pageNumber, setPageNumber] = useState(1);

  useEffect(() => {
    highlightCurrentNav("liLPA");
    loadData();
  }, []);

  const loadData = async () => {

    try {
      showLoader();
      const items = await getListItems(listName, currentSiteURL, "*,Project_x0020_Leader/Title,Project_x0020_Leader/Id", "Project_x0020_Leader", "")
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
        Plant: item.Plant,
        Department: item.Department,
        Zone: item.Zone,
        Machine: item.Machine,
        Date: item.Date
          ? new Date(item.Date).toLocaleDateString()
          : "",
        Date_x0020_Started: item.Date_x0020_Started
          ? new Date(item.Date_x0020_Started).toLocaleDateString()
          : "",
        Date_x0020_Closed: item.Date_x0020_Closed
          ? new Date(item.Date_x0020_Closed).toLocaleDateString()
          : "",
        KPI: item.KPI,
        Problem_x0020_Description: item.Problem_x0020_Description,
        Frequency: item.Frequency,
        Frequency_x0020_Ratio: item.Frequency_x0020_Ratio,
        MaterialCost: item.MaterialCost,
        MaterialCostRatio: item.MaterialCostRatio,
        Seriousness: item.Seriousness,
        SeriousnessRatio: item.SeriousnessRatio,
        Priority_x0020_Value: item.Priority_x0020_Value,
        Problem_x0020_Classification: item.Problem_x0020_Classification,
        Type_x0020_of_x0020_Kaizen: item.Type_x0020_of_x0020_Kaizen,
        Project_x0020_Leader: item.Project_x0020_Leader?.Title,
        Receiving_x0020_Inspection_x002f: item.Receiving_x0020_Inspection_x002f,
        Primary_x0020_Production_x0020_S: item.Primary_x0020_Production_x0020_S,
        Secondary_x0020_Station: item.Secondary_x0020_Station,
        QC_x002f_QE_x0020_Audit_x0020__x: item.QC_x002f_QE_x0020_Audit_x0020__x,
        Plant_x0020_Rep_x0020__x002f__x0: item.Plant_x0020_Rep_x0020__x002f__x0,
        Customer_x0020_Concern: item.Customer_x0020_Concern,
        OData__x0033_CPR: item.OData__x0033_CPR,
        CPA: item.CPA,
        Yardpurge_x002f_Yardhold: item.Yardpurge_x002f_Yardhold,
        Warranty: item.Warranty,
        NOT_x0020_STARTED: item.NOT_x0020_STARTED ? "Yes" : "No",
        PLAN: item.PLAN,
        DO: item.DO,
        CHECK: item.CHECK,
        ACT: item.ACT,
        COMPLETE: item.COMPLETE,
        Man: item.Man,
        Material: item.Material,
        Method: item.Method,
        _x0034_M_x0020_Machine: item._x0034_M_x0020_Machine,
        Year: item.Year,
      }));

      setData(tableData);

    } catch (e) {
      console.error(e);
    } finally {
      hideLoader();
    }
  };

    const filteredData = React.useMemo(() => {
      let result = data;
  
      // if (filters.year) {
      //   result = result.filter(item =>
      //     String(item.Year) === String(filters.year)
      //   );
      // }
      if (filters.year !== "All") {
  result = result.filter(
    item => String(item.Year) === String(filters.year)
  );
}
      return result.sort(
        (a: any, b: any) =>
        new Date(b.Date || b.Modified).getTime() -
        new Date(a.Date || a.Modified).getTime()
      );
    }, [data, filters]);


  const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
    const name = actionMeta?.name;

    if (!name) return;

    setFilters(prev => ({
      ...prev,
      [name]: selectedOption ? selectedOption.value : null
    }));
  };
  const handleRowClicked = (row: any) => {
    console.log("Row clicked:", row);

    // navigate to edit page
    navigate(`/QA-MatrixForm/${row.Id}`);
  };
  /* ---------------- TABLE COLUMNS ---------------- */

  // const columns = [
  //    {
  //    name: "Edit",
  //    selector: (row: any) => row.Id,
  //    cell: (row: any) => (
  //      <NavLink to={`/QAMatrixForm/${row.Id}`} title="Edit">
  //        <FontAwesomeIcon icon={faEdit} />
  //      </NavLink>
  //    ),
  //    width: "70px"
  //  },

  //   {
  //     name: "Plant",
  //     selector: (row: any) => row.Plant,
  //     sortable: true
  //   },
  //   {
  //     name: "Department",
  //     selector: (row: any) => row.Department,
  //     sortable: true
  //   },
  //   {
  //     name: "Zone",
  //     selector: (row: any) => row.Zone,
  //     sortable: false
  //   },
  //   {
  //     name: "Machine",
  //     selector: (row: any) => row.Machine,
  //     sortable: true
  //   },
  //   {
  //     name: "Date",
  //     selector: (row: any) => row.Date,
  //     sortable: true
  //   },


  //   {
  //     name: "Date Started",
  //     selector: (row: any) => row.Date_x0020_Started,
  //     sortable: true
  //   },
  //   {
  //     name: "Date Closed",
  //     selector: (row: any) => row.Date_x0020_Closed,
  //     sortable: true
  //   },
  //   {
  //     name: "KPI",
  //     selector: (row: any) => row.KPI,
  //     sortable: true
  //   },

  //   {
  //     name: "Problem Description",
  //     selector: (row: any) => row.Problem_x0020_Description,
  //     sortable: true
  //   },
  //   {
  //     name: "Frequency",
  //     selector: (row: any) => row.Frequency,
  //     sortable: true
  //   },
  //   {
  //     name: "Frequency Ratio",
  //     selector: (row: any) => row.Frequency_x0020_Ratio,
  //     sortable: true
  //   },
  //   {
  //     name: "Material Cost",
  //     selector: (row: any) => row.MaterialCost,
  //     sortable: true
  //   },
  //   {
  //     name: "MaterialCost Ratio",
  //     selector: (row: any) => row.MaterialCostRatio,
  //     sortable: true
  //   },
  //   {
  //     name: "Seriousness",
  //     selector: (row: any) => row.Seriousness,
  //     sortable: true
  //   },
  //   {
  //     name: "Seriousness Ratio",
  //     selector: (row: any) => row.SeriousnessRatio,
  //     sortable: true
  //   },
  //   {
  //     name: "Priority Value",
  //     selector: (row: any) => row.Priority_x0020_Value,
  //     sortable: true
  //   },
  //   {
  //     name: "Problem Classification",
  //     selector: (row: any) => row.Problem_x0020_Classification,
  //     sortable: true
  //   },
  //   {
  //     name: "Type of kaizen",
  //     selector: (row: any) => row.Type_x0020_of_x0020_Kaizen,
  //     sortable: true
  //   },
  //   {
  //     name: "Project Leader",
  //     selector: (row: any) => row.Project_x0020_Leader,
  //     sortable: true
  //   },
  //   {
  //     name: "Receiving inspection/Supplier Sort",
  //     selector: (row: any) => row.Receiving_x0020_Inspection_x002f,
  //     sortable: true
  //   },
  //   {
  //     name: "Primary Production Station",
  //     selector: (row: any) => row.Primary_x0020_Production_x0020_S,
  //     sortable: true
  //   },
  //   {
  //     name: "Secondary Station",
  //     selector: (row: any) => row.Secondary_x0020_Station,
  //     sortable: true
  //   },
  //   {
  //     name: "QC/QE Audit-Warehouse Sort",
  //     selector: (row: any) => row.QC_x002f_QE_x0020_Audit_x0020__x,
  //     sortable: true
  //   },
  //   {
  //     name: "Plant Rep / Repair",
  //     selector: (row: any) => row.Plant_x0020_Rep_x0020__x002f__x0,
  //     sortable: true
  //   },
  //   {
  //     name: "Customer Concern",
  //     selector: (row: any) => row.Customer_x0020_Concern,
  //     sortable: true
  //   },
  //   {
  //     name: "3CPR",
  //     selector: (row: any) => row.OData__x0033_CPR,
  //     sortable: true
  //   },
  //   {
  //     name: "CPA",
  //     selector: (row: any) => row.CPA,
  //     sortable: true
  //   },

  //   {
  //     name: "Yardpurge/Yardhold",
  //     selector: (row: any) => row.Yardpurge_x002f_Yardhold,
  //     sortable: true
  //   },
  //   {
  //     name: "Warranty",
  //     selector: (row: any) => row.Warranty,
  //     sortable: true
  //   },
  //   {
  //     name: "NOT STARTED",
  //     selector: (row: any) => row.NOT_x0020_STARTED,
  //     sortable: true
  //   },
  //   {
  //     name: "PLAN",
  //     selector: (row: any) => row.PLAN ? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "DO",
  //     selector: (row: any) => row.DO? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "CHECK",
  //     selector: (row: any) => row.CHECK? "Yes":"No",
  //     sortable: true
  //   },

  //   {
  //     name: "ACT",
  //     selector: (row: any) => row.ACT? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "COMPLETE",
  //     selector: (row: any) => row.COMPLETE? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "4M Man",
  //     selector: (row: any) => row.Man? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "4M Material",
  //     selector: (row: any) => row.Material? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "4M Method",
  //     selector: (row: any) => row.Method? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "4M Machine",
  //     selector: (row: any) => row._x0034_M_x0020_Machine? "Yes":"No",
  //     sortable: true
  //   },
  //   {
  //     name: "Year",
  //     selector: (row: any) => row.Year,
  //     sortable: true
  //   },







  // ];

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
          <NavLink to={`/QA-MatrixForm/${row.Id}`} title="Edit">
            <FontAwesomeIcon icon={faEdit} />
          </NavLink>
        );
      }
    },

    { headerName: "Plant", field: "Plant", sortable: true, filter: "agTextColumnFilter" },
    { headerName: "Department", field: "Department", sortable: true, filter: "agTextColumnFilter",minWidth: 160 },
    { headerName: "Zone", field: "Zone", sortable: true, filter: "agTextColumnFilter" },
    { headerName: "Machine", field: "Machine", sortable: true, filter: "agTextColumnFilter",minWidth: 135 },

    { headerName: "Date", field: "Date", sortable: true, filter: "agDateColumnFilter" },
    { headerName: "Date Started", field: "Date_x0020_Started", sortable: true,minWidth: 149 },
    { headerName: "Date Closed", field: "Date_x0020_Closed", sortable: true,minWidth: 145},

    { headerName: "KPI", field: "KPI", sortable: true },
    { headerName: "Problem Description", field: "Problem_x0020_Description", sortable: true,minWidth: 143 },

    { headerName: "Frequency", field: "Frequency", sortable: true,minWidth: 135 },
    { headerName: "Frequency Ratio", field: "Frequency_x0020_Ratio", sortable: true, minWidth: 135 },

    { headerName: "Material Cost", field: "MaterialCost", sortable: true,minWidth: 124 },
    { headerName: "Material Cost Ratio", field: "MaterialCostRatio", sortable: true,minWidth: 135 },

    { headerName: "Seriousness", field: "Seriousness", sortable: true,minWidth: 160 },
    { headerName: "Seriousness Ratio", field: "SeriousnessRatio", sortable: true,minWidth: 160 },

    { headerName: "Priority Value", field: "Priority_x0020_Value", sortable: true},
    { headerName: "Problem Classification", field: "Problem_x0020_Classification", sortable: true,minWidth: 153},

    { headerName: "Type of Kaizen", field: "Type_x0020_of_x0020_Kaizen", sortable: true },
    { headerName: "Project Leader", field: "Project_x0020_Leader", sortable: true },

    {
      headerName: "Receiving Inspection / Supplier Sort",
      field: "Receiving_x0020_Inspection_x002f",
      sortable: true,
      minWidth: 200
    },

    {
      headerName: "Primary Production Station",
      field: "Primary_x0020_Production_x0020_S",
      sortable: true,
      minWidth: 187
    },

    { headerName: "Secondary Station", field: "Secondary_x0020_Station", sortable: true,minWidth: 137 },

    {
      headerName: "QC / QE Audit - Warehouse Sort",
      field: "QC_x002f_QE_x0020_Audit_x0020__x",
      sortable: true,
      minWidth: 170
    },

    {
      headerName: "Plant Rep / Repair",
      field: "Plant_x0020_Rep_x0020__x002f__x0",
      sortable: true,
      minWidth: 131
    },

    { headerName: "Customer Concern", field: "Customer_x0020_Concern", sortable: true,minWidth: 131 },
    { headerName: "3CPR", field: "OData__x0033_CPR", sortable: true },
    { headerName: "CPA", field: "CPA", sortable: true },

    { headerName: "Yardpurge / Yardhold", field: "Yardpurge_x002f_Yardhold", sortable: true,minWidth: 137 },
    { headerName: "Warranty", field: "Warranty", sortable: true,minWidth: 129 },

    { headerName: "Not Started", field: "NOT_x0020_STARTED", sortable: true },

    {
      headerName: "Plan", field: "PLAN", sortable: true,
      cellRenderer: (row: any) => (
        <div title={row.data.PLAN ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.PLAN ? "Yes" : "No"}
          </span>
        </div>
      ),
    },

    {
      headerName: "Do", field: "DO", sortable: true,
      cellRenderer: (row: any) => (
        <div title={row.data.DO ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.DO ? "Yes" : "No"}
          </span>
        </div>
      ),
    },

    {
      headerName: "Check", field: "CHECK", sortable: true,
      cellRenderer: (row: any) => (
        <div title={row.data.CHECK ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.CHECK ? "Yes" : "No"}
          </span>
        </div>
      ),
    },

    {
      headerName: "Act", field: "ACT", sortable: true,
      cellRenderer: (row: any) => (
        <div title={row.data.ACT ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.ACT ? "Yes" : "No"}
          </span>
        </div>
      ),
    },

    {
      headerName: "Complete", field: "COMPLETE", sortable: true,
      cellRenderer: (row: any) => (
        <div title={row.data.COMPLETE ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.COMPLETE ? "Yes" : "No"}
          </span>
        </div>
      ),
    },

    {
      headerName: "4M Man", field: "Man", sortable: true,
      cellRenderer: (row: any) => (
        <div title={row.data.Man ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.Man ? "Yes" : "No"}
          </span>
        </div>
      )
    },

    {
      headerName: "4M Material", field: "Material", sortable: true,minWidth: 125,
      cellRenderer: (row: any) => (
        <div title={row.data.Material ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.Material ? "Yes" : "No"}
          </span>
        </div>
      )
    },

    {
      headerName: "4M Method", field: "Method", sortable: true,minWidth: 122,
      cellRenderer: (row: any) => (
        <div title={row.data.Method ? "Yes" : "No"}>

          <span className="status-icon-text">{row.data.Method ? "Yes" : "No"}
          </span>
        </div>
      )
    },

    {
      headerName: "4M Machine", field: "_x0034_M_x0020_Machine", sortable: true,minWidth: 125,
      valueFormatter: (p: any) => (p.value ? "Yes" : "No")
    },

    { headerName: "Year", field: "Year", sortable: true }
  ];
  return (

    <div className="container-fluid">

      <div className="light-box border-box-shadow">

        <div className="div-form-title">
          <div className="form-title">QA Matrix</div>
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
            suppressHorizontalScroll={false}
            suppressSizeToFit={true}
            suppressColumnHiding={true}
            suppressAutoSize={true}
            suppressColumnMoveAnimation={true}
            suppressMovableColumns={true}
          />

        </div>

      </div>

    </div>

  );
};

export default QAMatrixView;