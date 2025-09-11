import * as React from "react";
import { hideLoader, showLoader } from "../Shared/Loader";
import { Navigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http";
// import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "../CSS/EHSForm.css";
import { ActionStatus } from "../Constants/Contants";
import { showToast } from "../Shared/Toaster";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
// import SearchableDropdown from "../Shared/Dropdown";
// import FileUpload from "../Shared/FileUpload";
// import InputCheckBox from "../Shared/InputCheckBox";
// import { format } from "date-fns";
// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faChevronDown, faChevronUp } from "@fortawesome/free-solid-svg-icons";

export interface EHSFormProps{
    match:any;
    spContext:any;
    spHttpClient: SPHttpClient;
    context: any;
    history: any;
}

export interface EHSFormState{
}

 export default class EHSForm extends React.Component<EHSFormProps,EHSFormState>{

    private siteURL: string;
    // private sp = spfi().using(SPFx(this.props.context));
    public state = {
        formData:{
            
        },
        Homeredirect:false,
    }

    constructor(props: EHSFormProps){
        super(props);
        this.siteURL = this.props.spContext.siteAbsoluteUrl;
        console.log(this.siteURL);      //SynergyDev
    }
    
    public componentDidMount(): void {
        highlightCurrentNav("liEHSForm");
        document.title = "Mayco - Safety | EHS";
        this.getOnLoadData();
    }

    private getOnLoadData = async () => {
        try {
            showLoader();

           
            hideLoader();
        } catch (e) {
            console.log(e);
            this.onError();
        }
    }

    private onError = () => {
        showToast("error", ActionStatus.Error );
        hideLoader();
    }
    public render() {
        if (this.state.Homeredirect) {
            let url = "/Home";
            return (<Navigate to={url} />)
        }
        else{
            return(
                <React.Fragment>
                    <div className="container-fluid">
                        <div className="light-box border-box-shadow mb-4 brd-f1f1f1">
                            <div className="m-0 px-2">
                                <div className="col-12"><h2 className="mb-0 mt-2 text-center"> EHS Form</h2></div>
                            </div>
                            <div className="row">
                                <label className="text-end px-4"> <span className="text-danger">* </span> are mandatory fields</label>
                            </div>
                        </div>
                    </div>
                </React.Fragment>
            )
        }
    }

}