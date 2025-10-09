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

export interface SEWOProps {
  context: any;
}

export interface SEWOState {
  ActionsData: Array<{
    Id: number;
    Year: string;
    InjuredName: string;
    InjuryType: string;
    Injury_x0020_Date_x0020_Time: string;
    Plant: string;
    Department: string;
    Zone: string;
    Machine: string;
    AccidentType: string;
    AccidentCause: string;
  }>;
  loading: boolean;
  pageNumber: number;
  sortBy: number;
  sortOrder: boolean;
  ItemID: number;
  redirect: boolean;
  yearOptions: Array<{ label: string; value: string }>;
  selectedYear: string | undefined;
}

export default class SEWOView extends React.Component<SEWOProps, SEWOState> {
  private sp = spfi().using(SPFx(this.props.context));

  constructor(props: SEWOProps) {
    super(props);
    this.state = {
      ActionsData: [],
      loading: false,
      pageNumber: 1,
      sortBy: 1,
      ItemID: 0,
      sortOrder: false,
      redirect: false,
      yearOptions: [],
      selectedYear: "All",
    };
  }

  public componentDidMount() {
    document.title = "Mayco - Safety | SEWO View";
    this.loadListData();
  }

  private async loadListData() {
    try {
      showLoader();
      const items = await this.sp.web.lists
        .getByTitle("SEWO")
        .items.select("InjuryType/Title,InjuryType/Id,AccidentType/Title,AccidentType/Id,AccidentCause/Title,AccidentCause/Id,*").expand("InjuryType,AccidentType,AccidentCause").top(2000)
        .orderBy("Modified", false)();

      const tableData = items.map((item: any) => ({
        Id: item.Id,
        Year: item.Year,
        InjuredName: item.InjuredName,
        InjuryType: item.InjuryType?.Title || "",
        Injury_x0020_Date_x0020_Time: DateUtilities.getDateMMDDYYYY(item.Injury_x0020_Date_x0020_Time), 
        Injury_x0020_Date_x0020_TimeForGrid: `<span class='d-none'>${DateUtilities.getDateYYYYMMDDForSorting(item.Injury_x0020_Date_x0020_Time)}</span>${DateUtilities.getDateMMDDYYYY(item.Injury_x0020_Date_x0020_Time)}`, 
        Plant: item.Plant,
        Department: item.Department,
        Zone: item.Zone,
        Machine: item.Machine,
        AccidentType: item.AccidentType?.Title || "",
        AccidentCause: item.AccidentCause?.Title || ""
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
  private handleRowClicked = (row: any, Id?: any) => {
    let ID = row.Id ? row.Id : Id;
    this.setState({ ItemID: ID, redirect: true });
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
                <NavLink title="Edit" className="csrLink ms-draggable" to={`/SEWOForm/${record.Id}`}>
                  <FontAwesomeIcon icon={faEdit} ></FontAwesomeIcon>
                </NavLink>
              </div>
            </React.Fragment>
          );
        },
        width:'60px'
      },
      { name: "ID", selector: (row: any) => row.Id, sortable: false },
      { name: "Year", selector: (row: any) => row.Year, sortable: true },
      { name: "Name of injured", selector: (row: any) => row.InjuredName, sortable: true },
      { name: "Injury Type", selector: (row: any) => row.InjuryType, sortable: true },
      {
        name: "Injury Date Time",
        selector: (row: any) => row.Injury_x0020_Date_x0020_TimeForGrid,
        cell: (row: any) => <div className='' dangerouslySetInnerHTML={{ __html: row.Injury_x0020_Date_x0020_TimeForGrid }} />,
        sortable: true
      }, { name: "Plant", selector: (row: any) => row.Plant, sortable: true },
      { name: "Department", selector: (row: any) => row.Department, sortable: true },
      { name: "Zone", selector: (row: any) => row.Zone, sortable: true },
      { name: "Machine", selector: (row: any) => row.Machine, sortable: true },
      { name: "Accident Type", selector: (row: any) => row.AccidentType, sortable: true },
      { name: "Accident Cause", selector: (row: any) => row.AccidentCause, sortable: true },
    ];
    const filteredData = this.state.selectedYear && this.state.selectedYear !== "All" ? this.state.ActionsData.filter((item) => item.Year === this.state.selectedYear) : this.state.ActionsData;

    if (this.state.redirect) {
      let url = `/SEWOForm/${this.state.ItemID}`;
      return (<Navigate to={url} />);
    }
    return (
      <div className="container-fluid">
        <div className="light-box border-box-shadow">
             <div className="div-form-title">
                                <div className="form-title">SEWO</div>
                            </div>
                            <div className="mainContent px-4 borderLine">

          <div id="content" className="content p-2 pt-2">
            <div className="col-md-3">
              <div className="light-text">
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
