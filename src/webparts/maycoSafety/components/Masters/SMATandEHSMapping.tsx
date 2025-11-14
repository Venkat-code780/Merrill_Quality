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

export interface SMATandEHSMappingProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    currentUser: any,
}

export interface SMATandEHSMappingState {
    ActionsData: Array<Object>;
    loading: boolean;
    pageNumber: number;
    sortBy: number;
    sortOrder: boolean;
    searchText: string;
    isFormOpen: boolean;
    ItemId: number;
    formData: {
        Audit_categoriesId: number;
        Audit_SubCategory: string;
        Form_x0020_Type: string;
        Is_x0020_Active: boolean | null
    },
    redirect: boolean,
    isEdit: boolean,
    displayMessage: string,
    isUnauthorized: Boolean,
    AuditCategory: any,
    AuditCategoryid: number,

}

export default class SMATandEHSMapping extends React.Component<SMATandEHSMappingProps, SMATandEHSMappingState> {

    private ActionsList = "WCC/EHS mapping screen";
    private Audit_SubCategory;

    private Form_x0020_Type;
    private Is_x0020_Active;

    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: SMATandEHSMappingProps) {
        super(props);

        var lsTableProps = localStorage.getItem('PrvData');
        let TablePropsJson = lsTableProps != 'null' && lsTableProps != undefined && lsTableProps != null ? JSON.parse(lsTableProps) : null;

        this.state = {
            ActionsData: [],
            loading: true,
            pageNumber: TablePropsJson != null ? TablePropsJson.pageNumber : 1,
            sortBy: TablePropsJson != null ? TablePropsJson.sortBy : 1,
            sortOrder: TablePropsJson != null ? TablePropsJson.sortOrder == 'asc' ? true : false : false,
            searchText: TablePropsJson != null ? TablePropsJson.searchText : "",
            isFormOpen: false,
            ItemId: 0,
            formData: {
                Audit_categoriesId: 0,
                Audit_SubCategory: '',
                Form_x0020_Type: '',
                Is_x0020_Active: null
            },
            redirect: false,
            isEdit: false,
            displayMessage: '',
            isUnauthorized: false,
            AuditCategory: [],
            AuditCategoryid: 0,
        };

        this.Audit_SubCategory = React.createRef<HTMLInputElement>();
        this.Form_x0020_Type = React.createRef<HTMLSelectElement>();
        this.Is_x0020_Active = React.createRef<HTMLSelectElement>();


    }

    public componentDidMount() {
        highlightCurrentNav("liSMATandEHSMapping");
        document.title = "Mayco - Safety | SMAT and EHS Mapping";
        this.loadListData();
    }

    public componentDidUpdate() {
        if (this.state.redirect) {
            this.loadListData();
        }
    }

    private async loadListData() {
        try {
            showLoader();
            this.setState({ redirect: false });
            let lsTableProps = { 'PageNumber': 1, "sortOrder": false, "sortBy": 1, 'SearchKey': null };
            localStorage.setItem('PrvData', JSON.stringify(lsTableProps));

            let [AuditCategories, SmatandEhs] = await Promise.all([
                this.sp.web.lists.getByTitle('Audit_Categories').items.top(2000).orderBy("Title", true)(),
                this.sp.web.lists.getByTitle('WCC/EHS mapping screen').items.top(2000).select('Audit_categories/Title,Audit_categories/Id,*').expand('Audit_categories').orderBy("Modified", false)(),
            ])
            let tableData: { Id: any; Audit_categoriesId: any; Audit_categoriesTitle: any; Audit_SubCategory: any; Form_x0020_Type: any; Is_x0020_Active: any }[] = [];
            SmatandEhs.forEach(Src => {
                let tableObj = {
                    Id: Src.Id,
                    Audit_categoriesId: Src.Audit_categories.Id,
                    Audit_categoriesTitle: Src.Audit_categories.Title,
                    Audit_SubCategory: Src.Audit_SubCategory,
                    Form_x0020_Type: Src.Form_x0020_Type,
                    Is_x0020_Active: Src.Is_x0020_Active


                }
                tableData.push(tableObj);
            })
            let AuditCategoryOptions = AuditCategories.map((item: any) => ({
                label: item.Title,
                value: item.Id
            }));


            this.setState({ ActionsData: tableData, AuditCategory: AuditCategoryOptions, });
        }
        catch (e) {
            this.onError();
            console.log(e);
        }
        finally {
            hideLoader();
        }
    }

    private async editItem(Id: number) {
        try {
            var formData = { ...this.state.formData };
            //formData.IsActive = false;
            showLoader();
            this.setState({ isFormOpen: true, ItemId: Id, formData });
            await this.sp.web.lists.getByTitle(this.ActionsList).items.getById(Id)().then((item: any) => {
                if (item.Error) {
                    hideLoader();
                    console.log(item.Error);
                }
                else {
                    formData.Audit_categoriesId = item.Audit_categoriesId,
                        formData.Audit_SubCategory = item.Audit_SubCategory,
                        formData.Form_x0020_Type = item.Form_x0020_Type,
                        formData.Is_x0020_Active = item.Is_x0020_Active
                    //formData.IsActive = item.IsActive;
                    hideLoader();
                    this.setState({
                        formData,
                        AuditCategoryid: item.RootCauseId,

                    });
                    document.getElementById("divAuditCategory")?.getElementsByTagName('input')[0].focus();
                }
            })
        }
        catch (e) {
            this.onError();
            hideLoader();
            console.log(e);
        }
    }

    private addNew = () => {
        this.setState({ isFormOpen: true, ItemId: 0 });
        setTimeout(() => { document.getElementById("divAuditCategory")?.getElementsByTagName('input')[0].focus() }, 300);
    }


    private async checkDuplicate() {
        try {
            showLoader();
            const formData = { ...this.state.formData };

            let isValid = true;
            if (formData.Audit_SubCategory) formData.Audit_SubCategory = formData.Audit_SubCategory.trim();

            // const escapedTitle = formData.Audit_SubCategory.replace(/'/g, "''");


            let filterQuery = `Audit_SubCategory eq '${formData.Audit_SubCategory}' and Audit_categoriesId eq '${formData.Audit_categoriesId}' and Form_x0020_Type eq '${formData.Form_x0020_Type}'`;

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

    private handleSubmit = async (event: any) => {
        showLoader();
        try {
            event.preventDefault();
            var data = {
                AuditCategory: { val: (this.state.formData.Audit_categoriesId), required: true, Name: 'Audit Category', Type: ControlType.reactSelect, Focusid: 'divAuditCategory' },
                AuditSubcategory: { val: (this.state.formData.Audit_SubCategory.trim()), required: true, Name: 'Audit Sub-Category', Type: ControlType.string, Focusid: this.Audit_SubCategory },
                FormType: { val: (this.state.formData.Form_x0020_Type), required: true, Name: 'FormType', Type: ControlType.string, Focusid: this.Form_x0020_Type },
                IsActive: { val: (this.state.formData.Is_x0020_Active), required: true, Name: 'Is Active', Type: ControlType.string, Focusid: this.Is_x0020_Active },

            }
            let isValid = formValidation.FormValidation(data);

            if (isValid.status) {
                let validDuplicate = await this.checkDuplicate();

                if (validDuplicate) {
                    this.InsertOrUpdateDate();
                }

            } else {
                showToast("error", isValid.message);
                hideLoader();
            }
        }
        catch (e) {
            this.onError();
            hideLoader();
            console.log(e);
        }
    }

    private InsertOrUpdateDate() {
        try {
            let itemId = this.state.ItemId;
            let formData = { ...this.state.formData };
            if (formData.Audit_SubCategory) formData.Audit_SubCategory = formData.Audit_SubCategory.trim();
            if (itemId > 0) {
                this.sp.web.lists.getByTitle(this.ActionsList).items.getById(this.state.ItemId).update(formData).then((res) => {
                    let msg = "SMAT and EHS Mapping updated successfully";
                    this.setState({ displayMessage: msg, redirect: true });
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                this.sp.web.lists.getByTitle(this.ActionsList).items.add(formData).then((res) => {
                    let msg = "SMAT and EHS Mapping submitted successfully";
                    this.setState({ displayMessage: msg, redirect: true });
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
        }
        catch (e) {
            this.onError();
            hideLoader();
            console.log(e);
        }
    }

    private onSuccess = () => {
        this.closeForm();
        showToast("success", this.state.displayMessage);
        hideLoader();
    }

    private onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    }
    private handleAcive = (e: React.ChangeEvent<HTMLSelectElement>) => {
        const { name, value } = e.target;

        let boolValue: boolean | null = null;
        if (value === "true") boolValue = true;
        else if (value === "false") boolValue = false;

        this.setState((prevState: any) => ({
            formData: {
                ...prevState.formData,
                [name]: boolValue   //  now it's true/false/null
            }
        }));
    };

    private closeForm = () => {
        var formData = { ...this.state.formData };
        formData.Audit_SubCategory = '';
        formData.Audit_categoriesId = 0;
        formData.Form_x0020_Type = '',
            formData.Is_x0020_Active = null
        //formData.IsActive = true;
        this.setState({ isFormOpen: false, formData });
    }

    private onPageChange = (pageIndex: any) => {
        this.setState({ pageNumber: pageIndex });
    }
    private handleChangeClient = (selected: any) => {
        document.getElementById("divAuditCategory")?.classList.remove("searchMandatory");

        this.setState((prevState: Readonly<SMATandEHSMappingState>) => ({
            formData: {
                ...prevState.formData,
                Audit_categoriesId: !selected
                    ? null //  when cleared
                    : selected.value ??  // if { label, value }
                    selected.key ??    // if { text, key }
                    selected.Id ??     // if SharePoint object
                    selected           // if raw number
            }
        }));
    };

    private handleChangeDynamic = (event: any) => {
        const formData: any = { ...this.state.formData };
        const name = event.target.name;
        let value = event.target.type == 'checkbox' ? event.target.checked : event.target.value;
        formData[name] = value;
        this.setState({ formData });
    }
    private handleSecondary = (event: any) => {
        const formData = { ...this.state.formData };
        const name = event.target.name as keyof typeof formData;
        const value = event.target.value;

        formData[name as keyof typeof formData] = value as never;
        this.setState({ formData });
    }

    private handleRowClicked = (row: any, Id?: any) => {
        let ID = row.Id ? row.Id : Id;
        this.editItem(ID);
    }

    public render() {
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
                name: "Audit Category",
                selector: (row: { Audit_categoriesTitle: any; }, i: any) => row.Audit_categoriesTitle,
                sortable: true,
                cell: (record: { Audit_categoriesTitle: any; }) => {
                    return (
                        record.Audit_categoriesTitle
                    );
                },
            },
            {
                name: "Audit Sub-Category",
                selector: (row: { Audit_SubCategory: any; }, i: any) => row.Audit_SubCategory,
                sortable: true,
                cell: (record: { Audit_SubCategory: any; }) => {
                    return (
                        record.Audit_SubCategory
                    );
                },
            },
            {
                name: "Form Type",
                selector: (row: { Form_x0020_Type: any; }, i: any) => row.Form_x0020_Type,
                sortable: true,
                cell: (record: { Form_x0020_Type: any; }) => {
                    return (
                        record.Form_x0020_Type
                    );
                },
            },
               {
                name: "Is Active",
                selector: (row: { Is_x0020_Active: any; }, i: any) => row.Is_x0020_Active,
                sortable: true,
                cell: (record: { Is_x0020_Active: any; }) => {
                    return (
                        // record.Is_x0020_Active

                        record.Is_x0020_Active === true ? 'Yes' : record.Is_x0020_Active === false ? 'No' : ''
                    );
                },
            },
        ];

        if (this.state.isUnauthorized) {
            return <Navigate to="/UnAuthorized" />
        }
        else {

            return (
                <React.Fragment>

                    <div className="container-fluid">
                        <div className="light-box border-box-shadow">
                            <div className="div-form-title">
                                <div className="form-title">SMAT and EHS Mapping</div>
                                {this.state.isFormOpen && <span className="span-mandatory-text"> <span className="text-danger">* </span> are mandatory fields</span>}
                            </div>
                            <div className="p-2 mx-1 ViewTable">
                                {!this.state.isFormOpen &&
                                    <div className="text-end me-1" id="">
                                        <button type="button" id="btnNew" className="NewButton" title="New" onClick={this.addNew}>
                                            <FontAwesomeIcon icon={faPlus}></FontAwesomeIcon> New</button>
                                    </div>}
                                {this.state.isFormOpen &&
                                    <div className="">
                                        <div className="form-border-box p-2 mx-1 my-2">
                                            <div className="row">

                                                <div className="col-md-3">
                                                    <div className="custom-dropdown" id="divAuditCategory" title={(this.state.AuditCategory.find((i: { label: string; value: any }) => i.value == this.state.formData.Audit_categoriesId) as { label: string; value: any } | undefined)?.label}>
                                                        <SearchableDropdown label={"Audit Category"} Title={"Audit Category"} name={"Audit_categoriesId"} id={"ddAuditCategory"} className={""} selectedValue={this.state.formData.Audit_categoriesId} OptionsList={this.state.AuditCategory} OnChange={this.handleChangeClient} isRequired={true} disabled={false} placeholderText="" noOptionsMessage="No Audit Category available"></SearchableDropdown>
                                                    </div>
                                                </div>
                                                <div className="col-md-3">
                                                    <div className="light-text">
                                                        <input className="form-control" required={true} type="text" name="Audit_SubCategory" title={this.state.formData.Audit_SubCategory} value={this.state.formData.Audit_SubCategory} onChange={this.handleChangeDynamic} id="txtSubcategory" autoComplete="off" ref={this.Audit_SubCategory} maxLength={250} />
                                                        <label>Audit Sub-Category <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>


                                                <div className="col-md-3">
                                                    <div className="light-text">
                                                        <select className="form-select pt-3 ps-2" name="Form_x0020_Type" title={this.state.formData.Form_x0020_Type} value={this.state.formData.Form_x0020_Type} onChange={this.handleSecondary} ref={this.Form_x0020_Type} id="ddlFormType" >
                                                            <option value={""}>--Select One--</option>
                                                            <option value={'WCC'}>WCC</option>
                                                            <option value={'EHS'}>EHS</option>

                                                        </select>
                                                        <label>Form Type<span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-3">
                                                    <div className="light-text">
                                                        <select className="form-select pt-3 ps-2" name="Is_x0020_Active" title={this.state.formData.Is_x0020_Active === true ? 'Yes' : this.state.formData.Is_x0020_Active === false ? 'No' : ''} value={this.state.formData.Is_x0020_Active === null ? "" : String(this.state.formData.Is_x0020_Active)} onChange={this.handleAcive} ref={this.Is_x0020_Active} id="ddlFormActive" >
                                                            <option value={0}>--Select One--</option>
                                                            <option value={'true'}>Yes</option>
                                                            <option value={'false'}>No</option>
                                                        </select>
                                                        <label>Is Active<span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                            </div>


                                            <div className="col-sm-12 text-center py-3" id="">
                                                <button type="button" id="btnSubmit" className="btn btn-primary mx-2" title={this.state.ItemId ? 'Update' : 'Submit'} onClick={this.handleSubmit}>{this.state.ItemId ? 'Update' : 'Submit'}</button>
                                                <button type="button" id="btnCancel" className="btn btn-secondary" title="Cancel" onClick={this.closeForm}>Cancel</button>
                                            </div>
                                        </div>

                                    </div>
                                }
                                <TableGenerator columns={columns} data={this.state.ActionsData} onChange={this.onPageChange} prvPageNumber={this.state.pageNumber} prvDirection={this.state.sortOrder} fileName={"Actions"} className="sp-Datatable-hh" onRowClick={this.handleRowClicked} showPagination={true}></TableGenerator>
                            </div>
                        </div>
                    </div>

                </React.Fragment>
            )
        }
    }
}
