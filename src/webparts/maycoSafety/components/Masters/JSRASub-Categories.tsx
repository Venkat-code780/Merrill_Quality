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
export interface JSRASubCategoriesProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    currentUser: any,
}

export interface JSRASubCategoriesState {
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
        CategoryId: Number;
    },
    redirect: boolean,
    isEdit: boolean,
    displayMessage: string,
    isUnauthorized: Boolean,
    JSRACategory: any,

}

export default class JSRASubCategories extends React.Component<JSRASubCategoriesProps, JSRASubCategoriesState> {

    private ActionsList = "JSRASubCategories";
    private txtJSRASubCategory;
    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: JSRASubCategoriesProps) {
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
                Title: '',
                CategoryId: 0,
            },
            redirect: false,
            isEdit: false,
            displayMessage: '',
            isUnauthorized: false,
            JSRACategory: [],

        };

        this.txtJSRASubCategory = React.createRef<HTMLInputElement>();

    }

    public componentDidMount() {
        highlightCurrentNav("liJSRASub-Categories");
        document.title = "Mayco - Safety | JSRA Sub-Categories";
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

            let [JSRASubCategories, JSRACategory] = await Promise.all([
                this.sp.web.lists.getByTitle(this.ActionsList).items.top(2000).select('Title,Category/Title,Category/Id,*').expand('Category').orderBy("Modified", false)(),
                this.sp.web.lists.getByTitle('JSRACategories').items.select("Id,Title").top(2000).orderBy("Title", true)(),

            ])
            let tableData: { Id: any; Title: any; CategoryId: any; CategoryTitle: any; }[] = [];
            JSRASubCategories.forEach(Act => {
                let tableObj = {
                    Id: Act.Id,
                    Title: Act.Title,
                    CategoryId: Act.Category.Id,
                    CategoryTitle: Act.Category.Title,

                }
                tableData.push(tableObj);
            })
            let categoryOptions = JSRACategory.map((item: any) => ({
                label: item.Title,
                value: item.Id
            }));
            this.setState({ ActionsData: tableData, JSRACategory: categoryOptions });
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
            formData.Title = '';
            //formData.IsActive = false;
            showLoader();
            this.setState({ isFormOpen: true, ItemId: Id, formData });
            await this.sp.web.lists.getByTitle(this.ActionsList).items.getById(Id)().then((item: any) => {
                if (item.Error) {
                    hideLoader();
                    console.log(item.Error);
                }
                else {
                    formData.Title = item.Title;
                    formData.CategoryId = item.CategoryId;
                    //formData.IsActive = item.IsActive;
                    hideLoader();
                    this.setState({ formData });
                    document.getElementById("divCategory")?.getElementsByTagName('input')[0].focus()
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

    private async checkDuplicate() {
        try {
            showLoader();
            const formData = { ...this.state.formData };

            let isValid = true;

            // Escape single quotes in Title
            if (formData.Title) formData.Title = formData.Title.trim();
            // Build OData filter for all three fields
            // Note: Adjust property names according to your SharePoint list fields
            let filterQuery = `Title eq '${formData.Title}' and CategoryId eq ${formData.CategoryId}`;

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
                Category: { val: (this.state.formData.CategoryId), required: true, Name: 'JSRA Category', Type: ControlType.reactSelect, Focusid: "divCategory" },
                SubCategory: { val: (this.state.formData.Title.trim()), required: true, Name: 'JSRA Sub-Category', Type: ControlType.string, Focusid: this.txtJSRASubCategory },

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
            if (formData.Title) formData.Title = formData.Title.trim();

            if (itemId > 0) {
                this.sp.web.lists.getByTitle(this.ActionsList).items.getById(this.state.ItemId).update(formData).then((res) => {
                    let msg = "JSRA Sub-Category updated successfully";
                    this.setState({ displayMessage: msg, redirect: true });
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                this.sp.web.lists.getByTitle(this.ActionsList).items.add(formData).then((res) => {
                    let msg = "JSRA Sub-Category submitted successfully";
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
    private handleChangeClient = (selected: any) => {
        document.getElementById("divCategory")?.classList.remove("searchMandatory");

        this.setState((prevState: Readonly<JSRASubCategoriesState>) => ({
            formData: {
                ...prevState.formData,
                CategoryId: !selected
                    ? null
                    : selected.value ??  // if { label, value }
                    selected.key ??    // if { text, key }
                    selected.Id ??     // if SharePoint object
                    selected           // if raw number
            }
        }));
    };



    private closeForm = () => {
        var formData = { ...this.state.formData };
        formData.Title = '';
        formData.CategoryId = 0;
        //formData.IsActive = true;
        this.setState({ isFormOpen: false, formData });
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
                width: '60px',
                sortable: false
            },
            {
                name: "JSRA Sub-Category",
                selector: (row: { Title: any; }, i: any) => row.Title,
                sortable: true,
                cell: (record: { Title: any; }) => {
                    return (
                        record.Title
                    );
                },
            },
            {
                name: "JSRA Category",
                selector: (row: { CategoryTitle: any; }, i: any) => row.CategoryTitle,
                sortable: true,
                cell: (record: { CategoryTitle: any; }) => {
                    return (
                        record.CategoryTitle
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
                                <div className="form-title">JSRA Sub-Categories</div>
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
                                            <div className="row">
                                                <div className="col-md-3">
                                                    <div className="custom-dropdown" id="divCategory" title={(this.state.JSRACategory.find((i: { label: string; value: any }) => i.value == this.state.formData.CategoryId) as { label: string; value: any } | undefined)?.label}>
                                                        <Dropdown label={"JSRA Category"} Title={"JSRA Category"} name={"JSRA Category"} id={"CategoryDropdown"} className={"Category"} selectedValue={this.state.formData.CategoryId} OptionsList={this.state.JSRACategory} OnChange={this.handleChangeClient} isRequired={true} disabled={false} placeholderText="" noOptionsMessage="No JSRA Category available"></Dropdown>
                                                    </div>
                                                </div>
                                                <div className="col-md-3">
                                                    <div className="light-text">
                                                        <label>JSRA Sub-Category <span className="mandatoryhastrick">*</span></label>
                                                        <input className="form-control" required={true} type="text" name="Title" title={this.state.formData.Title} value={this.state.formData.Title} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.txtJSRASubCategory} maxLength={250} />
                                                    </div>
                                                </div>


                                                <div className="col-md-3 py-2 text-center" id="">
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
