import * as React from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Navigate, NavLink } from 'react-router-dom';
import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit } from "@fortawesome/free-solid-svg-icons";
import DateUtilities from "../Utilities/DateUtilities"
import Serachbledropdown from "../Shared/Dropdown";

export interface CheckListStep1Props {
  context: any;
}

export interface CheckListStep1State {
  ActionsData: Array<{
    Id: number;
    Date: string;
    Plant: string;
    Department: string;
    Zone: string;
    Machine: string;
    Audit_x0020_Score: string;
    Year: string;
    Auditor: string;
    Month: string;
  }>;
  loading: boolean;
  pageNumber: number;
  sortBy: number;
  sortOrder: boolean;
          ItemID:number,
  redirect:boolean,
   yearOptions: Array<{ label: string; value: string }>;
  selectedYear: string | undefined;
  
}

export default class CheckListStep1View extends React.Component<CheckListStep1Props,CheckListStep1State> {
  private sp = spfi().using(SPFx(this.props.context));

  constructor(props: CheckListStep1Props) {
    super(props);
    this.state = {
      ActionsData: [],
      loading: false,
      pageNumber: 1,
      sortBy: 1,
      ItemID: 0,
      sortOrder: false,
      redirect:false,
       yearOptions: [],
      selectedYear: "All",
    };
  }

  public componentDidMount() {
    document.title = "Mayco - Safety | CHECK-LIST STEP 1 View";
    this.loadListData();
  }

  private async loadListData() {
    try {
      showLoader();
      const items = await this.sp.web.lists
        .getByTitle("CheckListStep1")
        .items.top(2000)
        .orderBy("Modified", false)();

      const tableData = items.map((item: any) => ({
        Id: item.Id,
        Date: DateUtilities.getDateMMDDYYYY(item.Date), 
        DateForGrid: `<span class='d-none'>${DateUtilities.getDateYYYYMMDDForSorting(item.Date)}</span>${DateUtilities.getDateMMDDYYYY(item.Date)}`,
        Plant: item.Plant,
        Department: item.Department,
        Zone: item.Zone,
        Machine: item.Machine,
        Audit_x0020_Score: item.Audit_x0020_Score,
        Year: item.Year,
        Auditor: item.Auditor, // handles person field
        Month: item.Month,
      }));
          const yearList = tableData
        .map((i) => parseInt(i.Year, 10))
        .filter((y) => !isNaN(y));

      let yearOptions: { label: string; value: string }[] = [
        { label: "All", value: "All" }, // 🔹 Add "All" as first option
      ];

      if (yearList.length > 0) {
        const minYear = Math.min(...yearList);
        const currentYear = new Date().getFullYear();

        for (let y = minYear; y <= currentYear; y++) {
          yearOptions.push({ label: y.toString(), value: y.toString() });
        }
      }

      this.setState({ ActionsData: tableData,yearOptions });
    } catch (e) {
      console.error(e);
    } finally {
      hideLoader();
    }
  }
  
   private  handleRowClicked = (row:any,Id?:any) => {
        let ID = row.Id?row.Id:Id;
        this.setState({ItemID:ID,redirect:true});
      }
 private handleYearChange = (selected: any) => {
  // assuming Serachbledropdown sends { label, value }
  this.setState({ selectedYear: selected?.value || "All" });
};

  public render() {
    const columns = [
               {
                  name: "Edit",
                  //selector: "Id",
                  selector: (row: { Id: any; }, i: any) => row.Id,
                  cell: (record: { Id: any; }) => {
                    return (
                      <React.Fragment>
                        <div style={{ paddingLeft: '10px' }}>
                          <NavLink title="Edit" className="csrLink ms-draggable" to={`/CHECK-LISTSTEP1Form/${record.Id}`}>
                            <FontAwesomeIcon icon={faEdit} ></FontAwesomeIcon>
                          </NavLink>
                        </div>
                      </React.Fragment>
                    );
                  },
                  width:'60px',
                },
      { name: "ID", selector: (row: any) => row.Id, sortable: false,width:'60px' },
        {
        name: "Date",
        selector: (row: any) => row.DateForGrid,
        cell: (row: any) => <div className='' dangerouslySetInnerHTML={{ __html: row.DateForGrid }} />,
        sortable: true
      },
      { name: "Plant", selector: (row: any) => row.Plant, sortable: true },
      { name: "Department", selector: (row: any) => row.Department, sortable: true },
      { name: "Zone", selector: (row: any) => row.Zone, sortable: true },
      { name: "Machine", selector: (row: any) => row.Machine, sortable: true },
      { name: "Audit Score", selector: (row: any) => row.Audit_x0020_Score, sortable: true },
      { name: "Year", selector: (row: any) => row.Year, sortable: true },
      { name: "Auditor", selector: (row: any) => row.Auditor, sortable: true },
      { name: "Month", selector: (row: any) => row.Month, sortable: true },
    ];
    const filteredData =
  this.state.selectedYear && this.state.selectedYear !== "All"
    ? this.state.ActionsData.filter(
        (item) => item.Year === this.state.selectedYear
      )
    : this.state.ActionsData;
       if(this.state.redirect){
                    let url = `/CHECK-LISTSTEP1Form/${this.state.ItemID}`;
                return (<Navigate to={url}/>);
                 }
    return (
      
        <div className="container-fluid">
          <div className="light-box border-box-shadow">
              <div className="div-form-title">
                                <div className="form-title">CHECK-LIST STEP 1 View</div>
                            </div>
                             <div className="mainContent borderLine">
            <div className="row">
               <div className="col-md-3">
                <div className="light-text mt-4">
                   <label htmlFor="">
                    Year Filter 
                  </label>
                <div className="custom-dropdown" id="divRootcauese">
            <Serachbledropdown label={""} Title={"Year Filter"} name={"selectedYear"} id={undefined} className={""} selectedValue={this.state.selectedYear} OptionsList={this.state.yearOptions} OnChange={this.handleYearChange} isRequired={false} disabled={false}></Serachbledropdown>
            </div>
            </div>
            </div>
            <TableGenerator
              columns={columns}
              data={filteredData}
              onRowClick={this.handleRowClicked}
              fileName={"Actions"}
              showPagination={true}
            />
          </div>
          </div>
        </div>
      </div>
    );
  }
}
