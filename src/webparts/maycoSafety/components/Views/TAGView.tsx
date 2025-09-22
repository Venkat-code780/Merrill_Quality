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
import DateUtilities from "../Utilities/DateUtilities";
import Serachbledropdown from "../Shared/Dropdown";

export interface ActionsProps {
  context: any;
}

export interface ActionsState {
  ActionsData: Array<{
    Id: number;
    Plant:string;
    Department:string;
    Zone:string;
    Machine:string;
    Shifts:string;
    Name:string;
    Date:string;
    Completed_x0020_Date:string;
    CompletedBy:string;
    TAG_x0023_:string;
    Year:string;
    Completed_x0020_By:string
    
    
  }>;
  loading: boolean;
  pageNumber: number;
  sortBy: number;
  sortOrder: boolean;
          ItemID:number;
  redirect:boolean;
       yearOptions: Array<{ label: string; value: string }>;
  selectedYear: string | undefined;
}

export default class UCANView extends React.Component<ActionsProps,ActionsState> {
  private sp = spfi().using(SPFx(this.props.context));

  constructor(props: ActionsProps) {
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
    document.title = "Mayco - Safety | Actions";
    this.loadListData();
  }

  private async loadListData() {
    try {
      showLoader();
      const items = await this.sp.web.lists
        .getByTitle("Safety tag")
        .items.select("CompletedBy/Id,CompletedBy/Title,*").expand("CompletedBy").top(2000)
        .orderBy("Modified", false)();

      const tableData = items.map((item: any) => ({
        Id: item.Id,
        Department:item.Department,
        Zone:item.Zone,
        Machine:item.Machine,
        Shifts:item.Shifts,
        Name:item.Name,
        Date:item.Date,
        Completed_x0020_Date:item.Completed_x0020_Date,
        Completed_x0020_By:item.Completed_x0020_By,
        TAG_x0023_:item.TAG_x0023_,
        Year:item.Year,
        CompletedBy:Array.isArray(item.CompletedBy)
                    ? item.CompletedBy.map((u: any) => u.Title).join(", ")
                    : item.CompletedBy?.Title || "",
        Plant:item.Plant
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
                          <NavLink title="Edit" className="csrLink ms-draggable" to={`/TAGForm/${record.Id}`}>
                            <FontAwesomeIcon icon={faEdit} ></FontAwesomeIcon>
                          </NavLink>
                        </div>
                      </React.Fragment>
                    );
                  }
                },
      { name: "ID", selector: (row: any) => row.Id, sortable: false },
      { name: "Plant", selector: (row: any) => row.Plant, sortable: true },
      { name: "Department", selector: (row: any) => row.Department, sortable: true },
      { name: "Zone", selector: (row: any) => row.Zone, sortable: true },
      { name: "Machine", selector: (row: any) => row.Machine, sortable: true },
      { name: "Shifts", selector: (row: any) => row.Shifts, sortable: true },
      { name: "Name", selector: (row: any) => row.Name, sortable: true },
      { name: "Date ", selector: (row: any) =>DateUtilities.getDateMMDDYYYY(row.Date) , sortable: true },
      { name: "Completed Date", selector: (row: any) =>DateUtilities.getDateMMDDYYYY(row.Completed_x0020_Date) , sortable: true },
      { name: "CompletedBy", selector: (row: any) => row.Completed_x0020_By, sortable: true },
      { name: "TAG#", selector: (row: any) =>(row.TAG_x0023_) , sortable: true },
      { name: "Year", selector: (row: any) =>row.Year, sortable: true },
     { name: "Completed By", selector: (row: any) => row.CompletedBy, sortable: true },
    ];
          const filteredData =this.state.selectedYear && this.state.selectedYear !== "All" ? this.state.ActionsData.filter( (item) => item.Year === this.state.selectedYear ): this.state.ActionsData;

       if(this.state.redirect){
                    let url = `/TAGForm/${this.state.ItemID}`;
                return (<Navigate to={url}/>);
                 }
    return (
        <div className="container-fluid">
          <div className="FormContent border-none">
            <div className="title">TAG</div>
                  <div id="content" className="content p-2 pt-2">
             <div className="col-md-3">
                <div className="form-floating">
                <div className="custom-dropdown" id="divRootcauese">
            <Serachbledropdown label={"Year Filter"} Title={"Year Filter"} name={"selectedYear"} id={undefined} className={""} selectedValue={this.state.selectedYear} OptionsList={this.state.yearOptions} OnChange={this.handleYearChange} isRequired={false} disabled={false}></Serachbledropdown>
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
    );
  }
}
