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
import { Navigate } from "react-router-dom";

export interface ActionsProps {
    match:any;
    spContext:any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    currentUser : any,
}

export interface ActionsState {
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
        RootCause:Number;
        SecondaryRootCause: Number;
    },
    redirect: boolean,
    isEdit: boolean,
    displayMessage:string,
    isUnauthorized: Boolean,
    RootCauses:any,
    SecondaryRootCauses:any
}

export default class MicroRootCauses extends React.Component<ActionsProps, ActionsState> {

    private ActionsList = "Actions";
    private txtLeadSourceName;
    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: ActionsProps){
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
                RootCause: 0,
                SecondaryRootCause:0
            },
            redirect: false,
            isEdit: false,
            displayMessage:'',
            isUnauthorized: false,
            RootCauses:[],
            SecondaryRootCauses:[]
        };

        this.txtLeadSourceName = React.createRef<HTMLInputElement>();
    }

    public componentDidMount(){
        highlightCurrentNav("liActions");
        document.title = "Mayco - Safety | Actions";
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

           let  [Actions,RootCauses,SecondaryRootCauses]=await Promise.all([
            this.sp.web.lists.getByTitle(this.ActionsList).items.top(2000).select('Title,RootCause/Title,RootCause/Id,SecondaryRootCause/Title,SecondaryRootCause/Id,*').expand('RootCause,SecondaryRootCause').orderBy("Modified", false)(),
            this.sp.web.lists.getByTitle('RootCauses').items.top(2000).orderBy("Title", true)(),
            this.sp.web.lists.getByTitle('SecondaryRootCauses').items.top(2000).orderBy("Title", true)(),
           ])
           let tableData: { Id: any; Title: any; RootCauseId: any; SecondaryRootCauseId: any; RootCauseTitle: any; SecondaryRootCauseTitle: any; }[]=[];
           Actions.forEach(Act=>{
             let tableObj = {
                                Id: Act.Id,
                                Title: Act.Title,
                                RootCauseId:Act.RootCause.Id,
                                SecondaryRootCauseId:Act.SecondaryRootCause.Id,
                                RootCauseTitle:Act.RootCause.Title,
                                SecondaryRootCauseTitle:Act.SecondaryRootCause.Title,
                            }
                tableData.push(tableObj);
           })
        this.setState({ ActionsData: tableData,RootCauses, SecondaryRootCauses});
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
                    //formData.IsActive = item.IsActive;
                    hideLoader();
                    this.setState({ formData });
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
        this.setState({ isFormOpen: true, ItemId: 0 });
    }

    private async checkDuplicate(){
        try{
            showLoader();
            var formData = {...this.state.formData};
            let isValid = true;
            let escapedTitle = formData.Title.replace(/'/g, "''"); 
            let filterQuery = "Title eq '"+ escapedTitle +"'";

            if( this.state.ItemId > 0 ){
                filterQuery += " and Id ne "+this.state.ItemId+"";
            }

            await this.sp.web.lists.getByTitle(this.ActionsList).items.filter(filterQuery)().then( (res:any) =>{
                if( !res.Error && res.length > 0){
                    isValid = false;
                    var message = "Action already exists";
                    showToast( "error", message );
                    hideLoader();
                }
                else{
                    hideLoader();
                }
            })
            return isValid;
        }
        catch(e){
            this.onError();
            hideLoader();
            console.log(e);
        }
    }
    private handleSubmit =async (event:any) =>{
        showLoader();
        try{
            event.preventDefault();
            var data = {
                leadSource: { val: (this.state.formData.Title.trim()), required: true, Name: "'Lead Source'", Type: ControlType.string, Focusid: this.txtLeadSourceName }
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

            if( itemId > 0 ){
                this.sp.web.lists.getByTitle(this.ActionsList).items.getById(this.state.ItemId).update( formData ).then( (res) => {
                    let msg = "Action updated successfully";
                    this.setState({displayMessage: msg, redirect:true});
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else{
                this.sp.web.lists.getByTitle(this.ActionsList).items.add(formData).then( (res) => {
                    let msg = "Action submitted successfully";
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
        //formData.IsActive = true;
        this.setState({ isFormOpen: false, formData });
    }

    private onPageChange =(pageIndex:any)=>{
        this.setState({pageNumber: pageIndex});  
    }

    private sortOrder =(event:any,sortDirection:any)=>{
        this.setState({sortBy: event.id,sortOrder:sortDirection});     
    }

    private handleChangeDynamic = (event: any) => {
        const formData:any = {...this.state.formData};
        const name = event.target.name;
        let value = event.target.type == 'checkbox' ? event.target.checked : event.target.value;
        formData[name] = value;
        this.setState({formData});
    }

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
                width: '100px',
                cell: (record: { Id: any; }) => {
                    return (
                        <React.Fragment>
                            <div>
                                <button type="button" id="btnEdit" className="btn" title="Edit" onClick={ () =>this.editItem(record.Id)}>
                                    <FontAwesomeIcon icon={faEdit}></FontAwesomeIcon>
                                </button>
                            </div>
                        </React.Fragment>
                    );
                },
                sortable: false
            },
            {
                name: "Action",
                selector: (row: { Title: any; }, i: any) => row.Title,
                sortable: true,
                cell: (record: { Title:  any; }) => {
                    return (
                        record.Title
                    );
                },
            },
            {
                name: "Root Cause",
                selector: (row: { RootCauseTitle: any; }, i: any) => row.RootCauseTitle,
                sortable: true,
                cell: (record: { RootCauseTitle:  any; }) => {
                    return (
                        record.RootCauseTitle
                    );
                },
            },
            {
                name: "Secondary Root Cause",
                selector: (row: { SecondaryRootCauseTitle: any; }, i: any) => row.SecondaryRootCauseTitle,
                sortable: true,
                cell: (record: { SecondaryRootCauseTitle:  any; }) => {
                    return (
                        record.SecondaryRootCauseTitle
                    );
                },
            }
        ];

        if(this.state.isUnauthorized){
            return <Navigate to="/UnAuthorized" />
        }
        else{
            
            return(
                <React.Fragment>
                        <div id="content" className="content p-2 pt-2">
                            <div className="container-fluid">
                                <div className="FormContent border-none">
                                    <div className="title">Lead Source</div>
                                    <div className="" id="">
                                        { !this.state.isFormOpen && 
                                        <div className="text-end" id="">
                                            <button type="button" id="btnNew" className="SubmitButtons btn btn-new fw-bold" title="New" onClick={this.addNew}>
                                                <FontAwesomeIcon icon={faPlus}></FontAwesomeIcon> New</button>
                                        </div> }
                                        { this.state.isFormOpen && 
                                            <div className="" id="divNew">
                                                <div className="border-top mt-3 py-3">
                                                    <div className="row">
                                                        <div className="col-md-3">
                                                            <div className="form-floating">
                                                                <input className="form-control" required={true} placeholder="Lead Source" type="text" name="Title" title="LeadSource" value={ this.state.formData.Title} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.txtLeadSourceName} maxLength={250}/>
                                                                <label>Lead Source Name <span className="mandatoryhastrick">*</span></label>
                                                            </div>
                                                        </div>
                                                        {/* <InputCheckBox 
                                                            label="Is Active"
                                                            name="IsActive" 
                                                            checked={this.state.formData.IsActive} 
                                                            onChange={this.handleChangeDynamic} 
                                                            isdisable={false} 
                                                            isRequired={false}   
                                                            id="chckIsActiveLeadSource"                                                 
                                                        /> */}
                                                        <div className="col-md-3 btnDiv" id="">
                                                            <button type="button" id="btnSubmit" className="SubmitButtons btn" title="Submit" onClick={this.handleSubmit}>Submit</button>
                                                            <button type="button" id="btnCancel" className="CancelButtons btn btn-secondary" title="Cancel" onClick={this.closeForm}>Cancel</button>
                                                        </div>
                                                    </div>
                                                </div>
                                                <span id="spanErrorMessage" style={{display:"none", color:"red"}}></span>
                                            </div>
                                        }
                                    </div>
                                    <TableGenerator columns={columns} data={this.state.ActionsData} onChange={this.onPageChange} onSortChange={this.sortOrder} prvPageNumber={this.state.pageNumber} prvDirection={this.state.sortOrder} prvSort={this.state.sortBy} fileName={"Actions"} onRowClick={this.handleRowClicked} showPagination={true}></TableGenerator>
                                </div>
                            </div>
                        </div>
                </React.Fragment>
            )
        }
    }
}
