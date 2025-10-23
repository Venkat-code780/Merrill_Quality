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
import Dropdown from "../Shared/Dropdown";

export interface JSRAdetailsProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    currentUser: any,
}

export interface JSRAdetailsState {
    ActionsData: Array<Object>;
    loading: boolean;
    pageNumber: number;
    sortBy: number;
    sortOrder: boolean;
    searchText: string;
    isFormOpen: boolean;
    ItemId: number;
    formData: {
        CategoryId: Number;
        Sub_x0020_CategoryId: Number;
        Details: string;
        Mitigation_x002f_Controls_x0020_: string;
        Mitigation_x002f_Controls_x0020_0: string;
        Mitigation_x002f_Controls_x0020_1: string;
        Probability_x0020_1: string;
        Probability_x0020_2: string;
        Probability_x0020_3: string;


    },
    redirect: boolean,
    isEdit: boolean,
    displayMessage: string,
    isUnauthorized: Boolean,
    JSRACategories: any,
    JSRASubCategories: any
    FilteredSubCategories: any;
}

export default class JSRADetails extends React.Component<JSRAdetailsProps, JSRAdetailsState> {

    private ActionsList = "JSRA Details";

    private Details;
    private Mitigation_x002f_Controls_x0020_;
    private Mitigation_x002f_Controls_x0020_0;
    private Mitigation_x002f_Controls_x0020_1;
    private Probability_x0020_1;
    private Probability_x0020_2;
    private Probability_x0020_3;


    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: JSRAdetailsProps) {
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
                CategoryId: 0,
                Sub_x0020_CategoryId: 0,
                Details: '',
                Mitigation_x002f_Controls_x0020_: '',
                Mitigation_x002f_Controls_x0020_0: '',
                Mitigation_x002f_Controls_x0020_1: '',
                Probability_x0020_1: '',
                Probability_x0020_2: '',
                Probability_x0020_3: '',
            },
            redirect: false,
            isEdit: false,
            displayMessage: '',
            isUnauthorized: false,
            JSRACategories: [],
            JSRASubCategories: [],
            FilteredSubCategories: []

        };


        this.Details = React.createRef<HTMLTextAreaElement>();
        this.Mitigation_x002f_Controls_x0020_ = React.createRef<HTMLInputElement>();
        this.Mitigation_x002f_Controls_x0020_0 = React.createRef<HTMLInputElement>();
        this.Mitigation_x002f_Controls_x0020_1 = React.createRef<HTMLInputElement>();
        this.Probability_x0020_1 = React.createRef<HTMLInputElement>();
        this.Probability_x0020_2 = React.createRef<HTMLInputElement>();
        this.Probability_x0020_3 = React.createRef<HTMLInputElement>();
    }

    public componentDidMount() {
        highlightCurrentNav("liJSRADetails");
        document.title = "Mayco - Safety | JSRA Details";
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

            let [JSRADetails, JSRACategories, JSRASubCategories] = await Promise.all([
                this.sp.web.lists.getByTitle(this.ActionsList).items.top(2000).select('Sub_x0020_Category/Title,Sub_x0020_Category/Id,Category/Title,Category/Id,*').expand('Sub_x0020_Category,Category').orderBy("Modified", false)(),
                this.sp.web.lists.getByTitle('JSRACategories').items.select('Id,Title').top(2000).orderBy("Title", true)(),
                this.sp.web.lists.getByTitle('JSRASubCategories').items.top(2000).select('id,Title,Category/Title,Category/Id').expand('Category').orderBy("Modified", false)(),
            ])
            let tableData: { Id: any; Sub_x0020_CategoryId: any; Sub_x0020_CategoryTitle: any; CategoryId: any; CategoryTitle: any; Probability_x0020_1: any; Probability_x0020_2: any; Probability_x0020_3: any[]; Mitigation_x002f_Controls_x0020_: any[]; Mitigation_x002f_Controls_x0020_0: any[]; Mitigation_x002f_Controls_x0020_1: any[] }[] = [];
            JSRADetails.forEach(Act => {
                let tableObj = {
                    Id: Act.Id,
                    Sub_x0020_CategoryId: Act.Sub_x0020_Category.Id,
                    Sub_x0020_CategoryTitle: Act.Sub_x0020_Category.Title,
                    CategoryId: Act.Category.Id,
                    CategoryTitle: Act.Category.Title,
                    Probability_x0020_1: Act.Probability_x0020_1,
                    Probability_x0020_2: Act.Probability_x0020_2,
                    Probability_x0020_3: Act.Probability_x0020_3,
                    Mitigation_x002f_Controls_x0020_: Act.Mitigation_x002f_Controls_x0020_,
                    Mitigation_x002f_Controls_x0020_0: Act.Mitigation_x002f_Controls_x0020_0,
                    Mitigation_x002f_Controls_x0020_1: Act.Mitigation_x002f_Controls_x0020_1

                }
                tableData.push(tableObj);
            })
            let categoryOptions = JSRACategories.map((item: any) => ({
                label: item.Title,
                value: item.Id
            }));
            let subCategories = JSRASubCategories.map((item: any) => ({
                Id: item.Id,
                Title: item.Title,
                CategoryId: item.CategoryId ?? item.Category?.Id,  // ensure relation
            }));

            this.setState({ ActionsData: tableData, JSRACategories: categoryOptions, JSRASubCategories: subCategories, FilteredSubCategories: [] });
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
            formData.Details = '';
            //formData.IsActive = false;
            showLoader();
            this.setState({ isFormOpen: true, ItemId: Id, formData });
            await this.sp.web.lists.getByTitle(this.ActionsList).items.getById(Id)().then((item: any) => {
                if (item.Error) {
                    hideLoader();
                    console.log(item.Error);
                }
                else {
                    formData.CategoryId = item.CategoryId;
                    formData.Sub_x0020_CategoryId = item.Sub_x0020_CategoryId;
                    formData.Probability_x0020_1 = item.Probability_x0020_1;
                    formData.Probability_x0020_2 = item.Probability_x0020_2;
                    formData.Probability_x0020_3 = item.Probability_x0020_3;
                    formData.Mitigation_x002f_Controls_x0020_ = item.Mitigation_x002f_Controls_x0020_;
                    formData.Mitigation_x002f_Controls_x0020_0 = item.Mitigation_x002f_Controls_x0020_0;
                    formData.Mitigation_x002f_Controls_x0020_1 = item.Mitigation_x002f_Controls_x0020_1;
                    formData.Details = item.Details;

                    const filteredSubCategories = this.state.JSRASubCategories
                        .filter((sub: any) => sub.CategoryId === formData.CategoryId)
                        .map((sub: any) => ({
                            label: sub.Title,
                            value: sub.Id,
                        }));
                    hideLoader();
                    this.setState({ formData, FilteredSubCategories: filteredSubCategories });
                    document.getElementById("divCategory")?.getElementsByTagName('input')[0].focus();
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
        setTimeout(() => { document.getElementById("divCategory")?.getElementsByTagName('input')[0].focus() }, 300);

    }

    private handleSubmit = async (event: any) => {
        showLoader();
        try {
            event.preventDefault();
            var data = {
                Category: { val: this.state.formData.CategoryId, required: true, Name: 'Category', Type: ControlType.reactSelect, Focusid: "divCategory" },
                SubCategory: { val: this.state.formData.Sub_x0020_CategoryId, required: true, Name: 'Sub Category', Type: ControlType.reactSelect, Focusid: "divSubCategory" },
                Probability1: { val: (this.state.formData.Probability_x0020_1.trim()), required: true, Name: 'Probability 1', Type: ControlType.string, Focusid: this.Probability_x0020_1 },
                Probability2: { val: (this.state.formData.Probability_x0020_2.trim()), required: true, Name: 'Probability 2', Type: ControlType.string, Focusid: this.Probability_x0020_2 },
                Probability3: { val: (this.state.formData.Probability_x0020_3.trim()), required: true, Name: 'Probability 3', Type: ControlType.string, Focusid: this.Probability_x0020_3 },
                MigrationControls1: { val: (this.state.formData.Mitigation_x002f_Controls_x0020_.trim()), required: true, Name: 'Mitigation/Controls 1', Type: ControlType.string, Focusid: this.Mitigation_x002f_Controls_x0020_ },
                MigrationControls2: { val: (this.state.formData.Mitigation_x002f_Controls_x0020_0.trim()), required: true, Name: 'Mitigation/Controls 2', Type: ControlType.string, Focusid: this.Mitigation_x002f_Controls_x0020_0 },
                MigrationControls3: { val: (this.state.formData.Mitigation_x002f_Controls_x0020_1.trim()), required: true, Name: 'Mitigation/Controls 3', Type: ControlType.string, Focusid: this.Mitigation_x002f_Controls_x0020_1 },
                Details: { val: (this.state.formData.Details.trim()), required: true, Name: 'Details', Type: ControlType.string, Focusid: this.Details },

            }
            let isValid = formValidation.FormValidation(data);

            if (isValid.status) {
                // let validDuplicate = await this.checkDuplicate();
                this.InsertOrUpdateDate();
                // if( validDuplicate ){

                // }

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

            if (itemId > 0) {
                this.sp.web.lists.getByTitle(this.ActionsList).items.getById(this.state.ItemId).update(formData).then((res) => {
                    let msg = "JSRA Details updated successfully";
                    this.setState({ displayMessage: msg, redirect: true });
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                this.sp.web.lists.getByTitle(this.ActionsList).items.add(formData).then((res) => {
                    let msg = "JSRA Details submitted successfully";
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

    private closeForm = () => {
        var formData = { ...this.state.formData };
        formData.Details = '';
        formData.CategoryId = 0;
        formData.Sub_x0020_CategoryId = 0;
        formData.Mitigation_x002f_Controls_x0020_ = '';
        formData.Mitigation_x002f_Controls_x0020_0 = '';
        formData.Mitigation_x002f_Controls_x0020_1 = '';
        formData.Probability_x0020_1 = '';
        formData.Probability_x0020_2 = '';
        formData.Probability_x0020_3 = '';
        //formData.IsActive = true;
        this.setState({ isFormOpen: false, formData, FilteredSubCategories: [] });
    }

    private onPageChange = (pageIndex: any) => {
        this.setState({ pageNumber: pageIndex });
    }
    private handleChangeDynamic = (event: any) => {
        const formData: any = { ...this.state.formData };
        const name = event.target.name;
        let value = event.target.type == 'checkbox' ? event.target.checked : event.target.value;
        formData[name] = value;
        this.setState({ formData });
    }
    private handleChangeClient = (selected: any) => {

        const selectedCategoryId = !selected
            ? null
            : selected.value ??
            selected.key ??
            selected.Id ??
            selected;
        document.getElementById("divCategory")?.classList.remove("searchMandatory");

        // Filter SubCategories based on Category
        const filteredSubCategories = this.state.JSRASubCategories
            .filter((sub: any) => sub.CategoryId === selectedCategoryId)
            .map((sub: any) => ({
                label: sub.Title,
                value: sub.Id,
            }));

        this.setState((prevState: Readonly<JSRAdetailsState>) => ({
            formData: {
                ...prevState.formData,
                CategoryId: selectedCategoryId,
                Sub_x0020_CategoryId: 0, // reset subcategory when category changes
            },
            FilteredSubCategories: filteredSubCategories
        }));
    };


    private handleSubCategoryChange = (selected: any) => {
        document.getElementById("divSubCategory")?.classList.remove("searchMandatory");

        this.setState((prevState: Readonly<JSRAdetailsState>) => ({
            formData: {
                ...prevState.formData,
                Sub_x0020_CategoryId: !selected
                    ? null
                    : selected.value ??
                    selected.key ??
                    selected.Id ??
                    selected
            }
        }));
    };
    private handleChange = (event: any) => {
        const formData: any = { ...this.state.formData };

        const name = event.target.name;
        let inputValue = (event.target.type == "text" || event.target.type == "textarea") ? event.target.value : event.target.checked ? event.target.value : '';

        formData[name] = inputValue;
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
                name: "Category",
                selector: (row: { CategoryTitle: any; }, i: any) => row.CategoryTitle,
                sortable: true,
                cell: (record: { CategoryTitle: any; }) => {
                    return (
                        record.CategoryTitle
                    );
                },
            },
            {
                name: "Sub Category",
                selector: (row: { Sub_x0020_CategoryTitle: any; }, i: any) => row.Sub_x0020_CategoryTitle,
                sortable: true,
                cell: (record: { Sub_x0020_CategoryTitle: any; }) => {
                    return (
                        record.Sub_x0020_CategoryTitle
                    );
                },
            },
            {
                name: "Probability 1",
                selector: (row: { Probability_x0020_1: any; }, i: any) => row.Probability_x0020_1,
                sortable: true,
                cell: (record: { Probability_x0020_1: any; }) => {
                    return (
                        record.Probability_x0020_1
                    );
                },
            },
            {
                name: "Probability 2",
                selector: (row: { Probability_x0020_2: any; }, i: any) => row.Probability_x0020_2,
                sortable: true,
                cell: (record: { Probability_x0020_2: any; }) => {
                    return (
                        record.Probability_x0020_2
                    );
                },
            },
            {
                name: "Probability 3",
                selector: (row: { Probability_x0020_3: any; }, i: any) => row.Probability_x0020_3,
                sortable: true,
                cell: (record: { Probability_x0020_3: any; }) => {
                    return (
                        record.Probability_x0020_3
                    );
                },
            },
            {
                name: "Mitigation/Controls 1",
                selector: (row: { Mitigation_x002f_Controls_x0020_: any; }, i: any) => row.Mitigation_x002f_Controls_x0020_,
                sortable: true,
                cell: (record: { Mitigation_x002f_Controls_x0020_: any; }) => {
                    return (
                        record.Mitigation_x002f_Controls_x0020_
                    );
                },
            },
            {
                name: "Mitigation/Controls 2",
                selector: (row: { Mitigation_x002f_Controls_x0020_0: any; }, i: any) => row.Mitigation_x002f_Controls_x0020_0,
                sortable: true,
                cell: (record: { Mitigation_x002f_Controls_x0020_0: any; }) => {
                    return (
                        record.Mitigation_x002f_Controls_x0020_0
                    );
                },
            },
            {
                name: "Mitigation/Controls 3",
                selector: (row: { Mitigation_x002f_Controls_x0020_1: any; }, i: any) => row.Mitigation_x002f_Controls_x0020_1,
                sortable: true,
                cell: (record: { Mitigation_x002f_Controls_x0020_1: any; }) => {
                    return (
                        record.Mitigation_x002f_Controls_x0020_1
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
                                <div className="form-title">JSRA Details</div>
                                {this.state.isFormOpen && <span className="span-mandatory-text"> <span className="text-danger">* </span> are mandatory fields</span>}
                            </div>
                            <div className="p-2 mx-1">
                                {!this.state.isFormOpen &&
                                    <div className="text-end me-1" id="">
                                        <button type="button" id="btnNew" className="NewButton" title="New" onClick={this.addNew}>
                                            <FontAwesomeIcon icon={faPlus}></FontAwesomeIcon> New</button>
                                    </div>}
                                {this.state.isFormOpen &&
                                    <div className="">
                                        <div className="form-border-box p-2 mx-1 my-2">
                                            <div className="row pt-2 px-2">
                                                <div className="col-md-3">
                                                    <div className="custom-dropdown" id="divCategory" title={(this.state.JSRACategories.find((i: { label: string; value: any }) => i.value == this.state.formData.CategoryId) as { label: string; value: any } | undefined)?.label}>
                                                        <Dropdown label={"Category"} Title={"CategoryId"} name={"CategoryId"} id={"CategoryIdddl"} className={"UAType0Id"} selectedValue={this.state.formData.CategoryId} OptionsList={this.state.JSRACategories} OnChange={this.handleChangeClient} isRequired={true} disabled={false} placeholderText="" noOptionsMessage="No Category available"></Dropdown>
                                                    </div>
                                                </div>
                                                <div className="col-md-3">
                                                    <div className="custom-dropdown" id="divSubCategory" title={(this.state.FilteredSubCategories.find((i: { label: string; value: any }) => i.value == this.state.formData.Sub_x0020_CategoryId) as { label: string; value: any } | undefined)?.label}>
                                                        <Dropdown label={"Sub-Category"} Title={"Sub-CategoryId"} name={"Sub_x0020_CategoryId"} id={"Sub_x0020_CategoryIdddl"} className={"Sub_x0020_CategoryId"} selectedValue={this.state.formData.Sub_x0020_CategoryId} OptionsList={this.state.FilteredSubCategories} OnChange={this.handleSubCategoryChange} isRequired={true} disabled={false} placeholderText="" noOptionsMessage="No Sub Category available"></Dropdown>
                                                    </div>
                                                </div>
                                                <div className="col-md-3">
                                                    <div className="light-text">
                                                        <input className="form-control" required={true} type="text" name="Probability_x0020_1" title={this.state.formData.Probability_x0020_1} value={this.state.formData.Probability_x0020_1} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.Probability_x0020_1} maxLength={250} />
                                                        <label>Probability 1 <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-3">
                                                    <div className="light-text">
                                                        <input className="form-control" required={true} type="text" name="Probability_x0020_2" title={this.state.formData.Probability_x0020_2} value={this.state.formData.Probability_x0020_2} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.Probability_x0020_2} maxLength={250} />
                                                        <label>Probability 2 <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-3 mt-2">
                                                    <div className="light-text ">
                                                        <input className="form-control" required={true} type="text" name="Probability_x0020_3" title={this.state.formData.Probability_x0020_3} value={this.state.formData.Probability_x0020_3} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.Probability_x0020_3} maxLength={250} />
                                                        <label>Probability 3 <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-3 mt-2">
                                                    <div className="light-text">
                                                        <input className="form-control" required={true} type="text" name="Mitigation_x002f_Controls_x0020_" title={this.state.formData.Mitigation_x002f_Controls_x0020_} value={this.state.formData.Mitigation_x002f_Controls_x0020_} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.Mitigation_x002f_Controls_x0020_} maxLength={250} />
                                                        <label>Mitigation/Controls 1 <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-3 mt-2">
                                                    <div className="light-text">
                                                        <input className="form-control" required={true} type="text" name="Mitigation_x002f_Controls_x0020_0" title={this.state.formData.Mitigation_x002f_Controls_x0020_0} value={this.state.formData.Mitigation_x002f_Controls_x0020_0} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.Mitigation_x002f_Controls_x0020_0} maxLength={250} />
                                                        <label>Mitigation/Controls 2 <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-3 mt-2">
                                                    <div className="light-text">
                                                        <input className="form-control" required={true} type="text" name="Mitigation_x002f_Controls_x0020_1" title={this.state.formData.Mitigation_x002f_Controls_x0020_1} value={this.state.formData.Mitigation_x002f_Controls_x0020_1} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.Mitigation_x002f_Controls_x0020_1} maxLength={250} />
                                                        <label>Mitigation/Controls 3 <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-12 mt-2">
                                                    <div className="light-text">
                                                        <textarea className="form-control bs-textarea" rows={3} id="txtDetails" name="Details" ref={this.Details} value={this.state.formData.Details} onChange={this.handleChange} disabled={false} title={this.state.formData.Details} ></textarea>
                                                        <label>Details<span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>

                                                <div className="col-sm-12 text-center py-3" id="">
                                                    <button type="button" id="btnSubmit" className="btn btn-primary mx-2" title={this.state.ItemId ? 'Update' : 'Submit'} onClick={this.handleSubmit}>{this.state.ItemId ? 'Update' : 'Submit'}</button>
                                                    <button type="button" id="btnCancel" className="btn btn-secondary" title="Cancel" onClick={this.closeForm}>Cancel</button>
                                                </div>
                                            </div>
                                        </div>

                                    </div>
                                }
                                <TableGenerator columns={columns} data={this.state.ActionsData} onChange={this.onPageChange} prvPageNumber={this.state.pageNumber} prvDirection={this.state.sortOrder} fileName={"Actions"} onRowClick={this.handleRowClicked} showPagination={true}></TableGenerator>
                            </div>
                        </div>
                    </div>

                </React.Fragment>
            )
        }
    }
}