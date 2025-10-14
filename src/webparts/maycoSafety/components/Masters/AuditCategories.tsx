import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit, faPlus } from "@fortawesome/free-solid-svg-icons";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { ActionStatus, ControlType } from "../Constants/Contants";
import formValidation from "../Utilities/FormValidator";
// import InputCheckBox from "../Shared/InputCheckBox";
import { showToast } from "../Shared/Toaster";
import { Navigate, NavLink } from "react-router-dom";
import InputCheckBox from "../Shared/InputCheckBox";

export interface AuditcategoryProps {
    match:any;
    spContext:any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    currentUser : any,
}

export interface AuditcategoryState {
    ActionsData: Array<Object>;
    loading: boolean;
    pageNumber: number;
    sortBy: number;
    sortOrder: boolean;
    searchText: string;
    isFormOpen: boolean;
    ItemId: number;
    formData: {
        Title: string;
        Is_x0020_Active:boolean;
    
    },
    redirect: boolean,
    isEdit: boolean,
    displayMessage:string,
    isUnauthorized: Boolean,
    RootCauses:any,
    RootCausesid:number,
    
}

 export default class AuditCategories extends React.Component<AuditcategoryProps, AuditcategoryState> {

    private ActionsList = "Audit_Categories";
    private txtCategory;
    

    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: AuditcategoryProps){
        super(props);

        var lsTableProps = localStorage.getItem('PrvData');
        let TablePropsJson =  lsTableProps != 'null' && lsTableProps != undefined && lsTableProps != null ? JSON.parse(lsTableProps):null;

        this.state = {
            ActionsData: [],
            loading: true,
            pageNumber: TablePropsJson != null ? TablePropsJson.pageNumber : 1,
            sortBy: TablePropsJson != null ? TablePropsJson.sortBy: 1,
            sortOrder: TablePropsJson != null ? TablePropsJson.sortOrder == 'asc'? true: false : false,
            searchText: TablePropsJson != null ? TablePropsJson.searchText : "",
            isFormOpen: false,
            ItemId: 0,
            formData: {
                Title: '',
                Is_x0020_Active:true
           
               
            },
            redirect: false,
            isEdit: false,
            displayMessage:'',
            isUnauthorized: false,
            RootCauses:[],
               RootCausesid:0,
        };

        this.txtCategory = React.createRef<HTMLInputElement>();
     


    }

    public componentDidMount(){
        highlightCurrentNav("liAuditCategories");
        document.title = "Mayco - Safety | Audit Categories";
        this.loadListData();
    }

    public componentDidUpdate(){
        if( this.state.redirect) {
            this.loadListData();
        }
    }

    private async loadListData(){
        try{
            showLoader();
            this.setState({ redirect: false});
             let lsTableProps = {'PageNumber':1,"sortOrder":false,"sortBy":1,'SearchKey':null};
            localStorage.setItem('PrvData', JSON.stringify(lsTableProps));

           let  [AuditCategories]=await Promise.all([
            this.sp.web.lists.getByTitle('Audit_Categories').items.select('Id', 'Title','Is_x0020_Active').top(2000).orderBy("Modified", false)(),
           ])
           let tableData: { Id: any; Title: any;}[]=[];
           AuditCategories.forEach(Src=>{
             let tableObj = {
                                Id: Src.Id,
                                Title: Src.Title,
                                Is_x0020_Active:Src.Is_x0020_Active
                            
                            
                            }
                tableData.push(tableObj);
           })

           
        this.setState({ ActionsData: tableData});
        }
        catch(e){
            this.onError();
            console.log(e);
        }
        finally{
            hideLoader();
        }
    }

    private async editItem( Id: number ){
        try{
             var formData = {...this.state.formData};
            formData.Title = '';
            //formData.IsActive = false;
            showLoader();
            this.setState({ isFormOpen: true, ItemId: Id, formData});
            await this.sp.web.lists.getByTitle(this.ActionsList).items.getById(Id)().then( (item:any) => {
                if( item.Error ){
                    hideLoader();
                    console.log(item.Error);
                }
                else{
                    formData.Title = item.Title;
                    formData.Is_x0020_Active=item.Is_x0020_Active
                  //formData.IsActive = item.IsActive;
                    hideLoader();
                    this.setState({ formData,   
                     
                    },()=>{
                        this.txtCategory.current?.focus();
                    });
                }
            })
        }
        catch(e){
            this.onError();
            hideLoader();
            console.log(e);
        }
    }
    
    private addNew = () => {
        this.setState({ isFormOpen: true, ItemId: 0 },()=>{
            this.txtCategory.current?.focus();
        });
    }

private async checkDuplicate() {
    try {
        showLoader();
        const formData = { ...this.state.formData };

        let isValid = true;

        // Escape single quotes in Title
        const escapedTitle = formData.Title.replace(/'/g, "''");

        // Build OData filter for all three fields
        // Note: Adjust property names according to your SharePoint list fields
        let filterQuery = `Title eq '${escapedTitle}'`;

        if (this.state.ItemId > 0) {
            // Exclude the current item (for update scenario)
            filterQuery += ` and Id ne ${this.state.ItemId}`;
        }

        const results = await this.sp.web.lists
            .getByTitle(this.ActionsList)
            .items.filter(filterQuery)();

        if (results && results.length > 0) {
            isValid = false;
            showToast("error", "Record already exists");
        }

        hideLoader();
        return isValid;
    } catch (e) {
        this.onError();
        hideLoader();
        console.error(e);
        return false;
    }
}

   private handleChangeDynamic = (event: any) => {
        const formData:any = {...this.state.formData};
        const name = event.target.name;
        let value = event.target.type == 'checkbox' ? event.target.checked : event.target.value;
        formData[name] = value;
        this.setState({formData});
    }







    private handleSubmit =async (event:any) =>{
        showLoader();
        try{
            event.preventDefault();
            var data = {
                Status: { val: (this.state.formData.Title.trim()), required: true, Name: 'Category', Type: ControlType.string, Focusid: this.txtCategory },

                 
            }
            let isValid = formValidation.FormValidation( data );

            if( isValid.status ){ 
                let validDuplicate = await this.checkDuplicate();
                
                if( validDuplicate ){
                    this.InsertOrUpdateDate();
                }

            }else{
                showToast( "error", isValid.message  );
                hideLoader();
            }
        }
        catch(e){
            this.onError();
            hideLoader();
            console.log(e);
        }
    }

    private InsertOrUpdateDate() {
        try{
            let itemId = this.state.ItemId;
            let formData = {...this.state.formData};
            formData.Title = formData.Title ? formData.Title.trim() : "";
            if( itemId > 0 ){
                this.sp.web.lists.getByTitle(this.ActionsList).items.getById(this.state.ItemId).update( formData ).then( (res) => {
                    let msg = "Audit Categories updated successfully";
                    this.setState({displayMessage: msg, redirect:true});
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else{
                this.sp.web.lists.getByTitle(this.ActionsList).items.add(formData).then( (res) => {
                    let msg = "Audit Categories submitted successfully";
                    this.setState({displayMessage: msg, redirect:true});
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
        }
        catch(e){
            this.onError();
            hideLoader();
            console.log(e);
        }
    }

    private onSuccess = ( ) =>{
        this.closeForm();
        showToast( "success", this.state.displayMessage );
        hideLoader();
    }

    private onError = () =>{
        showToast( "error", ActionStatus.Error );
        hideLoader();
    }

    private closeForm= () =>{
        var formData = {...this.state.formData};
        formData.Title = '';
        formData.Is_x0020_Active = true;
        this.setState({ isFormOpen: false, formData });
    }

    private onPageChange =(pageIndex:any)=>{
        this.setState({pageNumber: pageIndex});  
    }

   

  private handleCheckbox = (event: React.ChangeEvent<HTMLInputElement>) => {
  const formData = { ...this.state.formData };
  const name = event.target.name as keyof typeof formData;
  const checked = event.target.checked;

  formData[name] = checked as never;
  this.setState({ formData });

  console.log("Checkbox Changed:", name, checked);
};



    private handleRowClicked = (row:any,Id?: any) => {
        let ID = row.Id? row.Id:Id;
        this.editItem(ID);
    }

    public render(){
        const columns = [
            {
                name: "Edit",
                selector: (row: { Id: any; }, i: any) => row.Id,
                export: false,
                
                cell: (record: { Id: any; }) => {
                    return (
                        <React.Fragment>
                            <div>
                                <NavLink id="btnEdit" className="csrLink ms-draggable" title="Edit" onClick={() => this.editItem(record.Id)} to={""}>
                                    <FontAwesomeIcon icon={faEdit}></FontAwesomeIcon>
                                </NavLink>
                            </div>
                        </React.Fragment>
                    );
                },
                width: '60px',
                sortable: false
            },
            {
                name: "Category",
                selector: (row: { Title: any; }, i: any) => row.Title,
                sortable: true,
                cell: (record: { Title:  any; }) => {
                    return (
                        record.Title
                    );
                },
            },
          {
               name: "Is Active",
            selector: (row: { Is_x0020_Active: boolean; }) => row.Is_x0020_Active,
               sortable: true,
             cell: (record: { Is_x0020_Active: boolean; }) => (
               <span>{record.Is_x0020_Active ? 'Yes' : 'No'}</span>
    ),
}
        ];

        if(this.state.isUnauthorized){
            return <Navigate to="/UnAuthorized" />
        }
        else{
            
            return(
                <React.Fragment>
                    <div className="container-fluid">
                         <div className="light-box border-box-shadow">
                            
                                 <div className="div-form-title">
                                    <div className="form-title">Audit Categories</div>
                                     {this.state.isFormOpen && <span className="span-mandatory-text"> <span className="text-danger">* </span> are mandatory fields</span>}
                                    </div>
                                    <div className="mainContent px-4 borderLine">
                                        <div>
                                        { !this.state.isFormOpen && 
                                        <div className="text-end me-1" id="">
                                            <button type="button" id="btnNew" className="NewButton" title="New" onClick={this.addNew}>
                                                <FontAwesomeIcon icon={faPlus}></FontAwesomeIcon> New</button>
                                        </div> }
                                        { this.state.isFormOpen && 
                                            <div className="">
                                                <div className="form-border-box p-2 mx-1 mt-2">
                                                    <div className="row">
                                                        <div className="col-md-3">
                                                            <div className="light-text">
                                                                <input className="form-control" required={true} type="text" name="Title" title={this.state.formData.Title} value={ this.state.formData.Title} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.txtCategory} maxLength={250}/>
                                                                <label>Category<span className="mandatoryhastrick">*</span></label>
                                                            </div>
                                                        </div>
                                                        <div className="col-md-3">
                                                            <div className="form-floating">
                                                               <InputCheckBox label={"Is Active"} name={"Is_x0020_Active"} checked={this.state.formData.Is_x0020_Active} onChange={this.handleCheckbox} isdisable={false} isRequired={false}></InputCheckBox>
                                                            </div>
                                                        </div>
                                                       
                                                        <div className="col-md-3 py-2 text-center" id="">
                                                            <button type="button" id="btnSubmit" className="btn btn-primary mx-2" title={this.state.ItemId ? 'Update' : 'Submit'} onClick={this.handleSubmit}>{this.state.ItemId? 'Update':'Submit'}</button>
                                                            <button type="button" id="btnCancel" className="btn btn-secondary" title="Cancel" onClick={this.closeForm}>Cancel</button>
                                                        </div>
                                                    </div>
                                                </div>
                                               
                                            </div>
                                        }
                                        </div>
                                    
                                   
                                    <TableGenerator columns={columns} data={this.state.ActionsData} onChange={this.onPageChange}  prvPageNumber={this.state.pageNumber} prvDirection={this.state.sortOrder} fileName={"Actions"} onRowClick={this.handleRowClicked} showPagination={true}></TableGenerator>
                                  </div>                                   
                         </div>
                    </div>
                   
                </React.Fragment>
            )
        }
    }
}
