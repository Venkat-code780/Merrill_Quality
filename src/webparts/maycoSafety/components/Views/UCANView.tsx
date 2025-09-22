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

export interface ActionsProps {
  context: any;
}

export interface ActionsState {
  ActionsData: Array<{
    Id: number;
    UCAN_x0020_Type:string;
    Plant:string;
    Department:string;
    Zone:string;
    Machine:string;
    UAType:string;
    Sub_x002d_Type:string;
    Shift:string;
    Reported_x0020_By:string;
    Location_x002f_Persons:string;
    Safety_x0020_Tag:string;
    Date:string;
    Original_x0020_Tag_x0020_No_x002:string;
    Date_x0020_Completed:string;
    Action_x0020_Completed:string;
    Modified:string;
    Year: string;
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
      .getByTitle("UCAN")
      .items.select(
        "Id",
        "UCAN_x0020_Type",
        "Plant",
        "Department",
        "Zone",
        "Machine",
        "UAType/Title",
        "Sub_x002d_Type/Title",
        "Shift",
        "Reported_x0020_By",
        "Location_x002f_Persons",
        "Safety_x0020_Tag",
        "Date",
        "Original_x0020_Tag_x0020_No_x002",
        "Date_x0020_Completed",
        "Action_x0020_Completed",
        "Modified",
        "Year"
      )
      .expand("UAType", "Sub_x002d_Type")
      .orderBy("Modified", false)
      .top(2000)(); 

    const tableData = items.map((item: any) => ({
      Id: item.Id,
      UCAN_x0020_Type: item.UCAN_x0020_Type,
      Plant: item.Plant,
      Department: item.Department,
      Zone: item.Zone,
      Machine: item.Machine,
      UAType: item.UAType?.Title || "",
      Sub_x002d_Type: item.Sub_x002d_Type?.Title || "",
      Shift: item.Shift,
      Reported_x0020_By: item.Reported_x0020_By,
      Location_x002f_Persons: item.Location_x002f_Persons,
      Safety_x0020_Tag: item.Safety_x0020_Tag,
      Date: item.Date,
      Original_x0020_Tag_x0020_No_x002: item.Original_x0020_Tag_x0020_No_x002,
      Date_x0020_Completed: item.Date_x0020_Completed,
      Action_x0020_Completed: item.Action_x0020_Completed,
      Modified: item.Modified,
      Year: item.Year,
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

    this.setState({ ActionsData: tableData, yearOptions });
  } catch (e) {
    console.error(e);
  } finally {
    hideLoader();
  }
}

  // private async loadListData() {
  //   try {
  //     showLoader();
  //     const items = await this.sp.web.lists
  //       .getByTitle("UCAN")
  //       .items.select("UAType/Title,UAType/Id,Sub_x002d_Type/Id,Sub_x002d_Type/Title,*").expand("UAType,Sub_x002d_Type").top(2000)
  //       .orderBy("Modified", false)();

  //     const tableData = items.map((item: any) => ({
  //       Id: item.Id,
  //       UCAN_x0020_Type:item.UCAN_x0020_Type,
  //       Plant:item.Plant,
  //       Department:item.Department,
  //       Zone:item.Zone,
  //       Machine:item.Machine,
  //       UAType:item.UAType?.Title ||"",
  //       Sub_x002d_Type:item.Sub_x002d_Type?.Title ||"",
  //       Shift:item.Shift,
  //       Reported_x0020_By:item.Reported_x0020_By,
  //       Location_x002f_Persons:item.Location_x002f_Persons,
  //       Safety_x0020_Tag:item.Safety_x0020_Tag,
  //       Date:item.Date,
  //       Original_x0020_Tag_x0020_No_x002:item.Original_x0020_Tag_x0020_No_x002,
  //       Date_x0020_Completed:item.Date_x0020_Completed,
  //       Action_x0020_Completed:item.Action_x0020_Completed,
  //       Modified:item.Modified,
  //        Year:item.Year,
  //     }));
  //       const yearList = tableData
  //       .map((i) => parseInt(i.Year, 10))
  //       .filter((y) => !isNaN(y));

  //     let yearOptions: { label: string; value: string }[] = [
  //       { label: "All", value: "All" }, // 🔹 Add "All" as first option
  //     ];

  //     if (yearList.length > 0) {
  //       const minYear = Math.min(...yearList);
  //       const currentYear = new Date().getFullYear();

  //       for (let y = minYear; y <= currentYear; y++) {
  //         yearOptions.push({ label: y.toString(), value: y.toString() });
  //       }
  //     }

  //     this.setState({ ActionsData: tableData,yearOptions });
  //   } catch (e) {
  //     console.error(e);
  //   } finally {
  //     hideLoader();
  //   }
  // }
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
                          <NavLink title="Edit" className="csrLink ms-draggable" to={`/UCANForm/${record.Id}`}>
                            <FontAwesomeIcon icon={faEdit} ></FontAwesomeIcon>
                          </NavLink>
                        </div>
                      </React.Fragment>
                    );
                  }
                },
      { name: "ID", selector: (row: any) => row.Id, sortable: false },
      { name: "Near miss, Unsafe condition", selector: (row: any) => row.UCAN_x0020_Type, sortable: true },
      { name: "Plant", selector: (row: any) => row.Plant, sortable: true },
      { name: "Department", selector: (row: any) => row.Department, sortable: true },
      { name: "Zone", selector: (row: any) => row.Zone, sortable: true },
      { name: "Machine", selector: (row: any) => row.Machine, sortable: true },
      { name: "UA Type", selector: (row: any) => row.UAType, sortable: true },
      { name: "Sub-Type", selector: (row: any) => row.Sub_x002d_Type, sortable: true },
      { name: "Shift", selector: (row: any) => row.Shift, sortable: true },
      { name: "Reported By", selector: (row: any) => row.Reported_x0020_By, sortable: true },
      { name: "Location/Persons", selector: (row: any) => row.Location_x002f_Persons, sortable: true },
      { name: "Safety Tag", selector: (row: any) => row.Safety_x0020_Tag ? "Yes":"No", sortable: true },
      { name: "Date ", selector: (row: any) =>DateUtilities.getDateMMDDYYYY(row.Date) , sortable: true },
      { name: "Original Tag No.", selector: (row: any) => row.Original_x0020_Tag_x0020_No_x002, sortable: true },
      { name: "Date Completed", selector: (row: any) =>DateUtilities.getDateMMDDYYYY(row.Date_x0020_Completed) , sortable: true },
      { name: "Action Completed", selector: (row: any) => row.Action_x0020_Completed, sortable: true },
      { name: "Modified", selector: (row: any) =>DateUtilities.getFriendlyDate(row.Modified) , sortable: true },
      { name: "Year", selector: (row: any) =>row.Year, sortable: true },


      
    ];
              const filteredData =this.state.selectedYear && this.state.selectedYear !== "All" ? this.state.ActionsData.filter( (item) => item.Year === this.state.selectedYear ): this.state.ActionsData;

       if(this.state.redirect){
                    let url = `/UCANForm/${this.state.ItemID}`;
                return (<Navigate to={url}/>);
                 }
    return (
        <div className="container-fluid">
          <div className="FormContent border-none">
            <div className="title">UCAN</div>
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
