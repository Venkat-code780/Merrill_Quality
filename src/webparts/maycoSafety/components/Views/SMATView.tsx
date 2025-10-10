import * as React from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Navigate, NavLink } from 'react-router-dom';
import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit,faTrashCan } from "@fortawesome/free-solid-svg-icons";
import DateUtilities from "../Utilities/DateUtilities"
import ModalApprovePopUp from "../Shared/ModalApprovePopUp";
import Serachbledropdown from "../Shared/Dropdown";

export interface SMATProps {
  context: any;
}

export interface SMATState {
  ActionsData: Array<{
    Id: number;
    WCCDate:string;
    Year: string;
    ShiftType:string;
    AuditorName:string,
    Plant:string;
    Department:string;
    Zone:string;
    Machine:string;
    WorkCell:string;
    ToolNumber:string;
    
    Supervisor:string;
    Comments:string;
    unsafeactCount:string;
    unsafeconditionCount:string;
    ActionCompleted:string;
    CompletedDate:string;
    
  }>;
  loading: boolean;
  pageNumber: number;
  sortBy: number;
  sortOrder: boolean;
          ItemID:number,
  redirect:boolean
  showDeleteModal: boolean;
  deleteItemId: number | null;
        yearOptions: Array<{ label: string; value: string }>;
  selectedYear: string | undefined;
}

export default class SMATView extends React.Component<SMATProps,SMATState> {
  private sp = spfi().using(SPFx(this.props.context));

  constructor(props: SMATProps) {
    super(props);
    this.state = {
      ActionsData: [],
      loading: false,
      pageNumber: 1,
      sortBy: 1,
      ItemID: 0,
      sortOrder: false,
      redirect:false,
        showDeleteModal: false,
      deleteItemId: null,
         yearOptions: [],
      selectedYear: "All",
    };
  }

  public componentDidMount() {
    document.title = "Mayco - Safety | SMAT View";
    this.loadListData();
  }

  private async loadListData() {
    try {
      showLoader();
      const items = await this.sp.web.lists
        .getByTitle("WCC")
        .items.top(3000)
        .orderBy("Modified", false)();

      const tableData = items.map((item: any) => ({
        Id: item.Id,
          WCCDate: DateUtilities.getDateMMDDYYYY(item.WCCDate), 
        WCCDateForGrid: `<span class='d-none'>${DateUtilities.getDateYYYYMMDDForSorting(item.WCCDate)}</span>${DateUtilities.getDateMMDDYYYY(item.WCCDate)}`, 
        Year:item.Year,
        ShiftType:item.ShiftType,
        AuditorName:item.AuditorName,
        Plant:item.Plant,
        Department:item.Department,
        Zone:item.Zone,
        Machine:item.Machine,
        WorkCell:item.WorkCell,
        ToolNumber:item.ToolNumber,
        Supervisor:item.Supervisor,
        Comments:item.Comments,
        unsafeactCount:item.unsafeactCount,
        unsafeconditionCount:item.unsafeconditionCount,
        ActionCompleted:item.ActionCompleted,
        CompletedDate:item.CompletedDate
      
      }));
       const yearList = tableData
        .map((i) => parseInt(i.Year, 10))
        .filter((y) => !isNaN(y));

      let yearOptions: { label: string; value: string }[] = [
        { label: "All", value: "All" }, //  Add "All" as first option
      ];

      if (yearList.length > 0) {
        const minYear = Math.min(...yearList);
        const currentYear = new Date().getFullYear();

        for (let y = minYear; y <= currentYear; y++) {
          yearOptions.push({ label: y.toString(), value: y.toString() });
        }
      }

      this.setState({ ActionsData: tableData,yearOptions});
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
    private openDeleteModal = (id: number) => {
    this.setState({ showDeleteModal: true, deleteItemId: id });
  };
                private handleYearChange = (selected: any) => {
  // assuming Serachbledropdown sends { label, value }
  this.setState({ selectedYear: selected?.value || "All" });
};

private handleConfirmDelete = async (id?: number | null) => {
  try {
    if (id != null) {
      showLoader();
      this.setState({ showDeleteModal: false});
      const parentId = id;

      // Delete children
      const childItems = await this.sp.web.lists
        .getByTitle("WccLine")
        .items.filter(`WCCID eq ${parentId}`)();

      for (const child of childItems) {
        await this.sp.web.lists.getByTitle("WccLine").items.getById(child.Id).delete();
      }

      // Delete parent
      await this.sp.web.lists.getByTitle("WCC").items.getById(parentId).delete();

      // Update UI
      this.setState(prev => ({
        ActionsData: prev.ActionsData.filter(item => item.Id !== parentId),
        deleteItemId: null,
      }));
    }
  } catch (err) {
    console.error("Delete failed:", err);
  } finally {
    hideLoader();
  }
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
                          <NavLink title="Edit" className="csrLink ms-draggable" to={`/SMATForm/${record.Id}`}>
                            <FontAwesomeIcon icon={faEdit} ></FontAwesomeIcon>
                          </NavLink>
                        </div>
                      </React.Fragment>
                    );
                  },
                  width:'60px',
                },
                    {
        name: "Delete",
        selector: (row: { Id: any }) => row.Id,
        cell: (record: { Id: number }) => (
          <div style={{ paddingLeft: "10px", cursor: "pointer" }}>
            <FontAwesomeIcon
              icon={faTrashCan}
              onClick={() => this.openDeleteModal(record.Id)} // CHANGED HERE
            />
          </div>
        ),
        width: "70px",
      },,
      { name: "ID", selector: (row: any) => row.Id, sortable: false,width:'60px' },
        {
        name: "Date",
        selector: (row: any) => row.WCCDateForGrid,
        cell: (row: any) => <div className='' dangerouslySetInnerHTML={{ __html: row.WCCDateForGrid }} />,
        sortable: true
      },
      { name: "Year", selector: (row: any) =>row.Year, sortable: true },
      { name: "Shifts", selector: (row: any) => row.ShiftType, sortable: true },
      { name: "Auditor's Name", selector: (row: any) => row.AuditorName, sortable: true },
      { name: "Plant", selector: (row: any) => row.Plant, sortable: true },
      { name: "Department", selector: (row: any) => row.Department, sortable: true },
      { name: "Zone", selector: (row: any) => row.Zone, sortable: true },
      { name: "Machine", selector: (row: any) => row.Machine, sortable: true },
      { name: "Work Cell", selector: (row: any) => row.WorkCell, sortable: true },
      { name: "Tool No.", selector: (row: any) => row.ToolNumber, sortable: true },
      { name: "Supervisor", selector: (row: any) => row.Supervisor, sortable: true },
      { name: "Comments", selector: (row: any) => row.Comments, sortable: true },
      { name: "Unsafe act count", selector: (row: any) => row.unsafeactCount, sortable: true },
      { name: "Unsafe Condition count", selector: (row: any) => row.unsafeconditionCount, sortable: true },
      { name: "Action Completed", selector: (row: any) => row.ActionCompleted, sortable: true },
      { name: "Supervisor Action Completed", selector: (row: any) =>DateUtilities.getDateMMDDYYYY(row.CompletedDate) , sortable: true },
    ];
                  const filteredData =this.state.selectedYear && this.state.selectedYear !== "All" ? this.state.ActionsData.filter( (item) => item.Year === this.state.selectedYear ): this.state.ActionsData;

       if(this.state.redirect){
                    let url = `/SMATForm/${this.state.ItemID}`;
                return (<Navigate to={url}/>);
                 }
    return (
        <div className="container-fluid">
          <div className="light-box border-box-shadow">
             <div className="div-form-title">
                                <div className="form-title">SMAT</div>
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
             <ModalApprovePopUp
          message="Are you sure you want to Delete?"
          modalHeader="modal-header-reject"
          title="Delete"
          isVisible={this.state.showDeleteModal}
          isSuccess={false}
          errorMessage=""
          comments={() => {}}
          commentsValue=""
          onConfirm={()=>this.handleConfirmDelete(this.state.deleteItemId)}
          onCancel={this.handleCancelDelete}
        />
          </div>
          </div>
        </div>
      </div>
    );
  }
}
