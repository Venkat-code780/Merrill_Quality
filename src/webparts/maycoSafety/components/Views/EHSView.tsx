import * as React from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Navigate, NavLink } from 'react-router-dom';
import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit, faTrashCan } from "@fortawesome/free-solid-svg-icons";
import DateUtilities from "../Utilities/DateUtilities"
import ModalApprovePopUp from "../Shared/ModalApprovePopUp";
import Serachbledropdown from "../Shared/Dropdown";
import { showToast } from "../Shared/Toaster";

export interface EHSProps {
  context: any;
}

export interface EHSState {
  ActionsData: Array<{
    Id: number;
    Date: string;
    Year: number;
    Shifts: number;
    AuditorName: string,
    Plant: string;
    Department: string;
    Zone: string;
    Machine: string;
    WorkCell: string;
    ToolNo: string;

    Supervisor: string;
    Comments: string;
    unsafeactCount: number;
    unsafeconditionCount: number;


  }>;
  loading: boolean;
  pageNumber: number;
  sortBy: number;
  sortOrder: boolean;
  ItemID: number,
  redirect: boolean
  showDeleteModal: boolean;
  deleteItemId: number | null;
  yearOptions: Array<{ label: string; value: string }>;
  selectedYear: any;
}

export default class EHSView extends React.Component<EHSProps, EHSState> {
  private sp = spfi().using(SPFx(this.props.context));

  constructor(props: EHSProps) {
    super(props);
    this.state = {
      ActionsData: [],
      loading: false,
      pageNumber: 1,
      sortBy: 1,
      ItemID: 0,
      sortOrder: false,
      redirect: false,
      showDeleteModal: false,
      deleteItemId: null,
      yearOptions: [],
      selectedYear: "All",
    };
  }

  public componentDidMount() {
    document.title = "Mayco - Safety | EHS View";
    this.loadListData();
  }

  private async loadListData() {
    try {
      showLoader();
      const items = await this.sp.web.lists
        .getByTitle("EHS")
        .items.top(3000)
        .orderBy("Modified", false)();

      const tableData = items.map((item: any) => ({
        Id: Number(item.Id),
        Date: DateUtilities.getDateMMDDYYYY(item.Date),
        DateForGrid: `<span class='d-none'>${DateUtilities.getDateYYYYMMDDForSorting(item.Date)}</span>${DateUtilities.getDateMMDDYYYY(item.Date)}`,

        Year: [null, undefined, ''].includes(item.Year) ? item.Year : Number(item.Year),
        Shifts: [null, undefined, ''].includes(item.Shifts) ? item.Shifts : Number(item.Shifts),
        AuditorName: item.AuditorName,
        Plant: item.Plant,
        Department: item.Department,
        Zone: item.Zone,
        Machine: item.Machine,
        WorkCell: item.WorkCell,
        ToolNo: item.ToolNo,
        Supervisor: item.Supervisor,
        Comments: item.Comments,
        unsafeactCount: [null, undefined, ''].includes(item.unsafeactCount) ? item.unsafeactCount : Number(item.unsafeactCount),
        unsafeconditionCount: [null, undefined, ''].includes(item.unsafeconditionCount) ? item.unsafeconditionCount : Number(item.unsafeconditionCount),


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

        for (let y = currentYear; y >= minYear; y--) {
          yearOptions.push({ label: y.toString(), value: y.toString() });
        }
      }
      else{
        let currYear = new Date().getFullYear();
        let lastYear = new Date().getFullYear()-1;
        yearOptions.push({ label: currYear.toString(), value: currYear.toString() });
        yearOptions.push({ label: lastYear.toString(), value: lastYear.toString() });
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
  private openDeleteModal = (id: number) => {
    this.setState({ showDeleteModal: true, deleteItemId: id });
    setTimeout(() => { document.getElementById("btnOK")?.focus() }, 300);
  };

  private handleConfirmDelete = async (id?: number | null) => {
    try {
      if (id != null) {
        showLoader();
        this.setState({ showDeleteModal: false });
        const parentId = id;
        // Delete children
        const childItems = await this.sp.web.lists
          .getByTitle("EHSLine")
          .items.filter(`EHSID eq ${parentId}`)();

        // for (const child of childItems) {
        //   await this.sp.web.lists.getByTitle("EHSLine").items.getById(child.Id).delete();
        // }
        await Promise.all(
          childItems.map(child =>
            this.sp.web.lists.getByTitle("EHSLine").items.getById(child.Id).delete()
          )
        );
        // Delete parent
        await this.sp.web.lists.getByTitle("EHS").items.getById(parentId).delete().then(res => {
          showToast("success", `EHS - ${parentId}  deleted Successfully`);
           this.setState(prev => ({
          ActionsData: prev.ActionsData.filter(item => item.Id !== parentId),
          deleteItemId: null,
        }));
        });
      }
    } catch (err) {
      console.error("Delete failed:", err);
    } finally {
      hideLoader();
    }
  };



  private handleYearChange = (selected: any, actionMeta?: any) => {
    // assuming Serachbledropdown sends { label, value }
    let value = actionMeta.action == 'clear' ? '' : selected.value;
    this.setState({ selectedYear: value });
  };

  private handleCancelDelete = () => {
    this.setState({ showDeleteModal: false, deleteItemId: null });
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
                <NavLink title="Edit" className="csrLink ms-draggable" to={`/EHSForm/${record.Id}`}>
                  <FontAwesomeIcon icon={faEdit} ></FontAwesomeIcon>
                </NavLink>
              </div>
            </React.Fragment>
          );
        },
        width: '60px',
      },
      {
        name: "Delete",
        selector: (row: { Id: any }) => row.Id,
        cell: (record: { Id: number }) => (
          <div style={{ paddingLeft: "10px", cursor: "pointer" }}>
            <span className="spn-del" title="Delete" onClick={() => this.openDeleteModal(record.Id)}>
              <FontAwesomeIcon
                icon={faTrashCan}
              />
            </span>
          </div>
        ),
        width: "70px",
      }, ,
      { name: "ID", selector: (row: any) => row.Id, sortable: true, width: '60px' },
      {
        name: "Date",
        selector: (row: any) => row.DateForGrid,
        cell: (row: any) => <div className='' dangerouslySetInnerHTML={{ __html: row.DateForGrid }} onClick={(event) => this.handleRowClicked(event, row.Id)} />,
        sortable: true
      },
      { name: "Year", selector: (row: any) => row.Year, sortable: true },
      { name: "Shifts", selector: (row: any) => row.Shifts, sortable: true },
      { name: "Auditor's Name", selector: (row: any) => row.AuditorName, sortable: true },
      { name: "Plant", selector: (row: any) => row.Plant, sortable: true },
      { name: "Department", selector: (row: any) => row.Department, sortable: true },
      { name: "Zone", selector: (row: any) => row.Zone, sortable: true },
      { name: "Machine", selector: (row: any) => row.Machine, sortable: true },
      { name: "Work Cell", selector: (row: any) => row.WorkCell, sortable: true },
      { name: "Tool No.", selector: (row: any) => row.ToolNo, sortable: true },
      { name: "Supervisor", selector: (row: any) => row.Supervisor, sortable: true },
      { name: "Comments", selector: (row: any) => row.Comments, sortable: true },
      { name: "Unsafe act count", selector: (row: any) => row.unsafeactCount, sortable: true },
      { name: "Unsafe Condition count", selector: (row: any) => row.unsafeconditionCount, sortable: true },

    ];
    const filteredData = this.state.selectedYear && this.state.selectedYear !== "All" ? this.state.ActionsData.filter((item) => item.Year == this.state.selectedYear) : this.state.ActionsData;

    if (this.state.redirect) {
      let url = `/EHSForm/${this.state.ItemID}`;
      return (<Navigate to={url} />);
    }
    return (
      <div className="container-fluid">
        <div className="light-box border-box-shadow">
          <div className="div-form-title">
            <div className="form-title">EHS View</div>
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
              <ModalApprovePopUp
                message="Are you sure you want to Delete?"
                modalHeader="modal-header-reject"
                title="Delete"
                isVisible={this.state.showDeleteModal}
                isSuccess={false}
                errorMessage=""
                comments={() => { }}
                commentsValue=""
                onConfirm={() => this.handleConfirmDelete(this.state.deleteItemId)}
                onCancel={this.handleCancelDelete}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
