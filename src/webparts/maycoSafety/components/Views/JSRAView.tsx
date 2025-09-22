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

export interface ActionsProps {
  context: any;
}

export interface ActionsState {
  ActionsData: Array<{
    Id: number;
    Date:string;
    Plant:string;
    Department:string;
    Zone:string;
    Machine:string;
    Work_x0020_Cell:string;
    Supervisor:string;
    Shift:string;
    Task_x0020_Analyzed:string;
     Tool_x0020_Number:string;
     Hot_x0020_Work_x0020_Permit_x002:string;
     Hot_x0020_Work_x0020_Permit_x0020:string;
     Electrical_x0020_Permit_x0020_Re:string;
     Electrical_x0020_Permit_x0020_Ac:string;
     Confined_x0020_Space_x0020_Permi:string;
     Confined_x0020_Space_x0020_Permi0:string;
     Supervisor_x0020_Name:string;
     Supervisor_x0020_Date:string;
     Year:string;
        
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

export default class SMATView extends React.Component<ActionsProps,ActionsState> {
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
        showDeleteModal: false,
      deleteItemId: null,
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
        .getByTitle("JSRA")
        .items.top(3000)
        .orderBy("Modified", false)();

      const tableData = items.map((item: any) => ({
        Id: item.Id,
        Date:item.Date,
        Plant:item.Plant,
        Department:item.Department,
        Zone:item.Zone,
        Machine:item.Machine,
        Work_x0020_Cell:item.Work_x0020_Cell,
        Supervisor:item.Supervisor,
        Shift:item.Shift,
        Task_x0020_Analyzed:item.Task_x0020_Analyzed,
        Hot_x0020_Work_x0020_Permit_x002:item.Hot_x0020_Work_x0020_Permit_x002,
        Hot_x0020_Work_x0020_Permit_x0020:item.Hot_x0020_Work_x0020_Permit_x0020,
        Electrical_x0020_Permit_x0020_Re:item.Electrical_x0020_Permit_x0020_Re,
        Electrical_x0020_Permit_x0020_Ac:item.Electrical_x0020_Permit_x0020_Ac,
        Confined_x0020_Space_x0020_Permi:item.Confined_x0020_Space_x0020_Permi,
        Confined_x0020_Space_x0020_Permi0:item.Confined_x0020_Space_x0020_Permi0,
        Supervisor_x0020_Name:item.Supervisor_x0020_Name,
        Supervisor_x0020_Date:item.Supervisor_x0020_Date,
        Year:item.Year,
        Tool_x0020_Number:item.Tool_x0020_Number
   
      
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
    private openDeleteModal = (id: number) => {
    this.setState({ showDeleteModal: true, deleteItemId: id });
  };


 private handleConfirmDelete = async (id?: number | null) => {
    try {
      if (id != null) {
        showLoader();
        const parentId = id;
  
        // Delete children
        const childItems = await this.sp.web.lists
          .getByTitle("JSRA Line")
          .items.filter(`JSRA_x0020_ID eq ${parentId}`)();
  
        for (const child of childItems) {
          await this.sp.web.lists.getByTitle("JSRA Line").items.getById(child.Id).delete();
        }
  
        // Delete parent
        await this.sp.web.lists.getByTitle("JSRA").items.getById(parentId).delete();
  
        // Update UI
        this.setState(prev => ({
          ActionsData: prev.ActionsData.filter(item => item.Id !== parentId),
          showDeleteModal: false,
          deleteItemId: null,
        }));
      }
    } catch (err) {
      console.error("Delete failed:", err);
    } finally {
      hideLoader();
    }
  };


          private handleYearChange = (selected: any) => {
  // assuming Serachbledropdown sends { label, value }
  this.setState({ selectedYear: selected?.value || "All" });
};


  //  private handleConfirmDelete = async () => {
  //   try {
  //     if (this.state.deleteItemId) {
  //       showLoader();
  //       await this.sp.web.lists.getByTitle("WCC").items.getById(this.state.deleteItemId).delete();
  //       this.setState({
  //         ActionsData: this.state.ActionsData.filter(item => item.Id !== this.state.deleteItemId),
  //         showDeleteModal: false,
  //         deleteItemId: null,
  //       });
  //     }
  //   } catch (err) {
  //     console.error("Delete failed:", err);
  //   } finally {
  //     hideLoader();
  //   }
  // };
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
                          <NavLink title="Edit" className="csrLink ms-draggable" to={`/JSRAForm/${record.Id}`}>
                            <FontAwesomeIcon icon={faEdit} ></FontAwesomeIcon>
                          </NavLink>
                        </div>
                      </React.Fragment>
                    );
                  }
                },
                    {
        name: "Delete",
        selector: (row: { Id: any }) => row.Id,
        cell: (record: { Id: number }) => (
          <div style={{ paddingLeft: "10px", cursor: "pointer" }}>
            <FontAwesomeIcon
              icon={faTrashCan}
              onClick={() => this.openDeleteModal(record.Id)} // 🔥 CHANGED HERE
            />
          </div>
        ),
      },,
      { name: "ID", selector: (row: any) => row.Id, sortable: false },
      { name: "Date ", selector: (row: any) =>DateUtilities.getDateMMDDYYYY(row.Date) , sortable: true },
      { name: "Plant", selector: (row: any) => row.Plant, sortable: true },
      { name: "Department", selector: (row: any) => row.Department, sortable: true },
      { name: "Zone", selector: (row: any) => row.Zone, sortable: true },
      { name: "Machine", selector: (row: any) => row.Machine, sortable: true },
      { name: "Work Cell", selector: (row: any) => row.Work_x0020_Cell, sortable: true },
      { name: "Supervisor", selector: (row: any) => row.Supervisor, sortable: true },
      { name: "Shift", selector: (row: any) => row.Shift, sortable: true },
      { name: "Task Analyzed", selector: (row: any) => row.Task_x0020_Analyzed, sortable: true },
      { name: "Tool Number.", selector: (row: any) => row.Tool_x0020_Number, sortable: true },
      { name: "Hot Work Permit Required", selector: (row: any) => row.Hot_x0020_Work_x0020_Permit_x002 ? "Yes":"No", sortable: true },
      { name: "Hot Work Permit Acquired", selector: (row: any) => row.Hot_x0020_Work_x0020_Permit_x0020 ? "Yes":"No", sortable: true },
      { name: "Electrical Permit Required", selector: (row: any) => row.Electrical_x0020_Permit_x0020_Re? "Yes":"No", sortable: true },
      { name: "Electrical Permit Acquired", selector: (row: any) => row.Electrical_x0020_Permit_x0020_Ac? "Yes":"No", sortable: true },
      { name: "Confined Space Permit Required", selector: (row: any) => row.Confined_x0020_Space_x0020_Permi? "Yes":"No", sortable: true },
      { name: "Confined Space Permit Acquired", selector: (row: any) => row.Confined_x0020_Space_x0020_Permi0? "Yes":"No", sortable: true },
      { name: "Supervisor Name", selector: (row: any) => row.Supervisor_x0020_Name, sortable: true },
      { name: "Supervisor Date", selector: (row: any) =>DateUtilities.getDateMMDDYYYY(row.Supervisor_x0020_Date) , sortable: true },
      { name: "Year", selector: (row: any) =>row.Year, sortable: true },
      
    ];
      const filteredData =this.state.selectedYear && this.state.selectedYear !== "All" ? this.state.ActionsData.filter( (item) => item.Year === this.state.selectedYear ): this.state.ActionsData;

       if(this.state.redirect){
                    let url = `/JSRAForm/${this.state.ItemID}`;
                return (<Navigate to={url}/>);
                 }
    return (
        <div className="container-fluid">
          <div className="FormContent border-none">
            <div className="title">JSRA</div>
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
             <ModalApprovePopUp
          message="Are you sure you want to delete this record?"
          modalHeader="modal-header"
          title="Delete"
          isVisible={this.state.showDeleteModal}
          isSuccess={true}
          errorMessage=""
          comments={() => {}}
          commentsValue=""
          onConfirm={()=>this.handleConfirmDelete(this.state.deleteItemId)}
          onCancel={this.handleCancelDelete}
        />
          </div>
        </div>
      </div>
    );
  }
}
