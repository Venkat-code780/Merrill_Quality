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
export interface SecondaryRootCausesProps {
    match: any;
    spContext: any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
    currentUser: any,
}

export interface SecondaryRootCausesState {
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
        RootCauseId: number;
    },
    redirect: boolean,
    isEdit: boolean,
    displayMessage: string,
    isUnauthorized: Boolean,
    RootCauses: any,
    RootCausesid: number,

}

export default class SecondaryRootCauses extends React.Component<SecondaryRootCausesProps, SecondaryRootCausesState> {

    private ActionsList = "SecondaryRootCauses";
    private txtsecondaryrootcause;

    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: SecondaryRootCausesProps) {
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
                RootCauseId: 0,

            },
            redirect: false,
            isEdit: false,
            displayMessage: '',
            isUnauthorized: false,
            RootCauses: [],
            RootCausesid: 0,
        };

        this.txtsecondaryrootcause = React.createRef<HTMLInputElement>();


    }

    public componentDidMount() {
        highlightCurrentNav("liSecondaryRootCauses");
        document.title = "Mayco - Safety | Secondary Root Causes";
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

            let [RootCauses, SecondaryRootCauses] = await Promise.all([
                this.sp.web.lists.getByTitle('RootCauses').items.top(2000).orderBy("Title", true)(),
                this.sp.web.lists.getByTitle('SecondaryRootCauses').items.top(2000).select('Title,RootCause/Title,RootCause/Id,*').expand('RootCause').orderBy("Modified", false)(),
            ])
            let tableData: { Id: any; Title: any; RootCauseId: any; RootCauseTitle: any; }[] = [];
            SecondaryRootCauses.forEach(Src => {
                let tableObj = {
                    Id: Src.Id,
                    Title: Src.Title,
                    RootCauseId: Src.RootCause.Id,
                    RootCauseTitle: Src.RootCause.Title,

                }
                tableData.push(tableObj);
            })
            let RootCausesOptions = RootCauses.map((item: any) => ({
                label: item.Title,
                value: item.Id
            }));

            this.setState({ ActionsData: tableData, RootCauses: RootCausesOptions, });
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
                    formData.RootCauseId = item.RootCauseId;
                    //formData.IsActive = item.IsActive;
                    hideLoader();
                    this.setState({
                        formData,
                        RootCausesid: item.RootCauseId,

                    }, () => {
                        this.txtsecondaryrootcause.current?.focus();
                    });
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
        this.setState({ isFormOpen: true, ItemId: 0 }, () => {
            this.txtsecondaryrootcause.current?.focus();
        });
    }
    private async checkDuplicate() {
        try {
            showLoader();
            const formData = { ...this.state.formData };
            let isValid = true;
            if (formData.Title) formData.Title = formData.Title.trim();
            let filterQuery = `Title eq '${this.state.formData.Title}' and RootCauseId eq ${formData.RootCauseId}`;

            if (this.state.ItemId > 0) {
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
                Action: { val: (this.state.formData.Title.trim()), required: true, Name: 'Secondary Root Cause', Type: ControlType.string, Focusid: this.txtsecondaryrootcause },
                RootCause: { val: (this.state.formData.RootCauseId), required: true, Name: 'Root Cause', Type: ControlType.reactSelect, Focusid: 'divRootCause' },


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
                    let msg = "Secondary Root Cause updated successfully";
                    this.setState({ displayMessage: msg, redirect: true });
                    this.onSuccess();
                }, (error) => {
                    console.log(error);
                    this.onError();
                })
            }
            else {
                this.sp.web.lists.getByTitle(this.ActionsList).items.add(formData).then((res) => {
                    let msg = "Secondary Root Cause submitted successfully";
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
        formData.Title = '';
        formData.RootCauseId = 0;
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
    private handleChangeClient = (selected: any) => {

        document.getElementById("divRootCause")?.classList.remove("searchMandatory");

        this.setState((prevState: Readonly<SecondaryRootCausesState>) => ({
            formData: {
                ...prevState.formData,
                RootCauseId: !selected
                    ? null //  when cleared
                    : selected.value ??  // if { label, value }
                    selected.key ??    // if { text, key }
                    selected.Id ??     // if SharePoint object
                    selected           // if raw number
            }
        }));
    };

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
                name: "Secondary Root Cause",
                selector: (row: { Title: any; }, i: any) => row.Title,
                sortable: true,
                cell: (record: { Title: any; }) => {
                    return (
                        record.Title
                    );
                },
            },
            {
                name: "Root Cause",
                selector: (row: { RootCauseTitle: any; }, i: any) => row.RootCauseTitle,
                sortable: true,
                cell: (record: { RootCauseTitle: any; }) => {
                    return (
                        record.RootCauseTitle
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
                                <div className="form-title">Secondary Root Causes</div>
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
                                                    <div className="light-text">
                                                        <input className="form-control" required={true} type="text" name="Title" title={this.state.formData.Title} value={this.state.formData.Title} onChange={this.handleChangeDynamic} id="txtLeadSourceName" autoComplete="off" ref={this.txtsecondaryrootcause} maxLength={250} />
                                                        <label>Secondary Root Cause <span className="mandatoryhastrick">*</span></label>
                                                    </div>
                                                </div>
                                                <div className="col-md-3">
                                                    <div className="custom-dropdown" id="divRootCause" title={(this.state.RootCauses.find((i: { label: string; value: any }) => i.value == this.state.formData.RootCauseId) as { label: string; value: any } | undefined)?.label}>
                                                        <SearchableDropdown label={""} Title={"Root Cause"} name={"RootCauseId"} id={"ddRootCause"} className={""} selectedValue={this.state.formData.RootCauseId} OptionsList={this.state.RootCauses} OnChange={this.handleChangeClient} isRequired={true} disabled={false} placeholderText="" noOptionsMessage="No Root Cause available"></SearchableDropdown>
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
