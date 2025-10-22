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
import SearchableDropdown from "../Shared/Dropdown";

export interface MicrorootcausesProps {
    match:any;
    spContext:any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    currentUser : any,
}

export interface MicrorootcausesState {
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
        RootCauseId:number;
        SecondaryRootCauseId: number;
    },
    redirect: boolean,
    isEdit: boolean,
    displayMessage:string,
    isUnauthorized: Boolean,
    RootCauses:any,
    SecondaryRootCauses:any
    RootCausesid:number,
    SecondaryRootCauseid:number,
    FilteredSecondaryrootCauses:any

    
}

 export default class Microrootcauses extends React.Component<MicrorootcausesProps, MicrorootcausesState> {

    private ActionsList = "MicroRootCauses";
    private MicroRootCause;


    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: MicrorootcausesProps){
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
                RootCauseId: 0,
                SecondaryRootCauseId:0
            },
            redirect: false,
            isEdit: false,
            displayMessage:'',
            isUnauthorized: false,
            RootCauses:[],
            SecondaryRootCauses:[],
                 FilteredSecondaryrootCauses:[],
               RootCausesid:0,
              SecondaryRootCauseid:0
        };

        this.MicroRootCause = React.createRef<HTMLInputElement>();


    }

    public componentDidMount(){
        highlightCurrentNav("liMicroRootCauses");
        document.title = "Mayco - Safety | Micro Root Causes";
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

           let  [MicroRootCauses,RootCauses,SecondaryRootCauses]=await Promise.all([
            this.sp.web.lists.getByTitle(this.ActionsList).items.top(2000).select('Title,RootCause/Title,RootCause/Id,SecondaryRootCause/Title,SecondaryRootCause/Id,*').expand('RootCause,SecondaryRootCause').orderBy("Modified", false)(),
            this.sp.web.lists.getByTitle('RootCauses').items.top(2000).orderBy("Title", true)(),
            this.sp.web.lists.getByTitle('SecondaryRootCauses').items.top(2000).orderBy("Title", true)(),
           ])
           let tableData: { Id: any; Title: any; RootCauseId: any; SecondaryRootCauseId: any; RootCauseTitle: any; SecondaryRootCauseTitle: any; }[]=[];
           MicroRootCauses.forEach(Act=>{
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
           let RootCaueseOptions = RootCauses.map((item: any) => ({
             label: item.Title,   
             value: item.Id       
             }));
                let SecondaryRootCauseOptions = SecondaryRootCauses.map((item: any) => ({
                            Id: item.Id,
                          Title: item.Title,
                         RootCauseId: item.RootCauseId ?? item.RootCause?.Id,  // ensure relation
                   }));

           
        this.setState({ ActionsData: tableData,RootCauses:RootCaueseOptions,SecondaryRootCauses:SecondaryRootCauseOptions,FilteredSecondaryrootCauses:[]});
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
                    formData.RootCauseId=item.RootCauseId;
                    formData.SecondaryRootCauseId=item.SecondaryRootCauseId;
                const filteredSecondary = this.state.SecondaryRootCauses
                     .filter((src: any) => src.RootCauseId === item.RootCauseId)
                     .map((src: any) => ({
                     label: src.Title,
                      value: src.Id,
                     }));
                    //formData.IsActive = item.IsActive;
                    hideLoader();
                    this.setState({ formData,
                        RootCausesid: item.RootCauseId,
                        FilteredSecondaryrootCauses: filteredSecondary 
                    },()=>{
                   this.MicroRootCause.current?.focus();
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
            this.MicroRootCause.current?.focus();});
    }

private async checkDuplicate() {
    try {
        showLoader();
        const formData = { ...this.state.formData };

        let isValid = true;

        // Escape single quotes in Title
                 if (formData.Title) formData.Title = formData.Title.trim();


        // Build OData filter for all three fields
        // Note: Adjust property names according to your SharePoint list fields
        let filterQuery = `Title eq '${this.state.formData.Title}' and RootCauseId eq ${formData.RootCauseId} and SecondaryRootCauseId eq ${formData.SecondaryRootCauseId}`;

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









    private handleSubmit =async (event:any) =>{
        showLoader();
        try{
            event.preventDefault();
            var data = {
                Action: { val: (this.state.formData.Title.trim()), required: true, Name: 'Micro Root Cause', Type: ControlType.string, Focusid: this.MicroRootCause },
                RootCause: { val: (this.state.formData.RootCauseId), required: true, Name: 'Root Cause', Type: ControlType.reactSelect, Focusid:'divRootcause' },
                SecondRootCause: { val: (this.state.formData.SecondaryRootCauseId), required: true, Name: 'Secondary Root Cause', Type: ControlType.reactSelect, Focusid:'divSecondaryRootcause'},

                 
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
             if (formData.Title) formData.Title = formData.Title.trim();
            if( itemId > 0 ){
                this.sp.web.lists.getByTitle(this.ActionsList).items.getById(this.state.ItemId).update( formData ).then( (res) => {
                    let msg = "Micro Root Cause updated successfully";
                    this.setState({displayMessage: msg, redirect:true});
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else{
                this.sp.web.lists.getByTitle(this.ActionsList).items.add(formData).then( (res) => {
                    let msg = "Micro Root Cause submitted successfully";
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

 private closeForm = () => {
  this.setState({
    isFormOpen: false,
    formData: {
      Title: '',
      RootCauseId: 0,
      SecondaryRootCauseId: 0,
      // IsActive: true  // uncomment if needed
    },
    FilteredSecondaryrootCauses: [] // reset dropdown options if needed
  });
};


    private onPageChange =(pageIndex:any)=>{
        this.setState({pageNumber: pageIndex});  
    }

    private handleChangeDynamic = (event: any) => {
        const formData:any = {...this.state.formData};
        const name = event.target.name;
        let value = event.target.type == 'checkbox' ? event.target.checked : event.target.value;
        formData[name] = value;
        this.setState({formData});
    }
 


private handleChangeClient = (selected: any) => {
  const selectedRootCauseId = selected?.value ?? 0;
      document.getElementById("divRootcause")?.classList.remove("searchMandatory");


  const filteredSecondaryCauses = this.state.SecondaryRootCauses
    .filter((sub: any) => sub.RootCauseId === selectedRootCauseId) //  filter correctly
    .map((sub: any) => ({
      label: sub.Title,
      value: sub.Id,
    }));

  this.setState((prevState) => ({
    formData: {
      ...prevState.formData,
      RootCauseId: selectedRootCauseId,       //  correct field
      SecondaryRootCauseId: 0                 // reset when root cause changes
    },
    FilteredSecondaryrootCauses: filteredSecondaryCauses
  }));
};

private handleSecondaryRootCauseChange = (selected: any) => {
  const selectedSecondaryId = selected?.value ?? 0;
        document.getElementById("divSecondaryRootcause")?.classList.remove("searchMandatory");


  this.setState((prevState) => ({
    formData: {
      ...prevState.formData,
      SecondaryRootCauseId: selectedSecondaryId   //  correct field
    }
  }));
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
                sortable: false,
                 width: '60px',
            },
            {
                name: "Micro Root Cause",
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
                        
                            <div className="container-fluid">
                                <div className="light-box border-box-shadow">
                                    
                            <div className="div-form-title">
                                <div className="form-title">Micro Root Causes</div>
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
                                                                <input className="form-control" required={true} type="text" name="Title" title={this.state.formData.Title} value={ this.state.formData.Title} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.MicroRootCause} maxLength={250}/>
                                                                <label>Micro Root Cause <span className="mandatoryhastrick">*</span></label>
                                                            </div>
                                                        </div>
                                                        <div className="col-md-3">
                                                            <div className="light-text">
                                                                      <label htmlFor="ddlRootCause">
                                                                        Root Cause<span className="mandatoryhastrick">*</span>
                                                                     </label>
                                                                     <div className="custom-dropdown" id="divRootcause" title={(this.state.RootCauses.find((i: { label: string; value: any }) => i.value == this.state.formData.RootCauseId) as { label: string; value: any } | undefined)?.label}>
                                                                    <SearchableDropdown label={""} Title={"Root Cause"} name={"RootCauseId"} id={"ddlRootCause"} className={"RootCauseId"} selectedValue={this.state.formData.RootCauseId} OptionsList={this.state.RootCauses} OnChange={this.handleChangeClient} isRequired={true} disabled={false} placeholderText="" noOptionsMessage="No Root Cause available"></SearchableDropdown>
                                                                 </div>
                                                              </div>
                                                        </div>
                                                    
                                                        <div className="col-md-3">
                                                            <div className="light-text">
                                                                    <label htmlFor="ddlSecondaryRootCause">
                                                                        Secondary Root Cause<span className="mandatoryhastrick">*</span>
                                                                     </label>
                                                                   <div className="custom-dropdown" id="divSecondaryRootcause" title={(this.state.FilteredSecondaryrootCauses.find((i: { label: string; value: any }) => i.value == this.state.formData.SecondaryRootCauseId) as { label: string; value: any } | undefined)?.label}>
                                                                <SearchableDropdown label={""} Title={"Secondary Root Cause"} name={"SecondaryRootCauseId"} id={"ddlSecondaryRootCause"} className={"SecondaryRootCauseId"} selectedValue={this.state.formData.SecondaryRootCauseId} OptionsList={this.state.FilteredSecondaryrootCauses} OnChange={this.handleSecondaryRootCauseChange} isRequired={true} disabled={false} placeholderText="" noOptionsMessage="No Secondary Root Cause available"></SearchableDropdown>
                                                             </div>
                                                        
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
                                    <TableGenerator columns={columns} data={this.state.ActionsData} onChange={this.onPageChange} prvPageNumber={this.state.pageNumber} prvDirection={this.state.sortOrder} fileName={"Actions"} onRowClick={this.handleRowClicked} showPagination={true}></TableGenerator>
                                </div>
                                </div>
                            </div>
                     
                </React.Fragment>
            )
        }
    }
}
