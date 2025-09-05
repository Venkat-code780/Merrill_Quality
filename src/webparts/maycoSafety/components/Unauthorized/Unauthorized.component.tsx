import * as React from "react";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faHandPaper } from "@fortawesome/free-solid-svg-icons";
import "../CSS/Unauthorized.css";
import { hideLoader } from "../Shared/Loader";

export interface UnAuthorizedProps {
    spContext: any;
}

export interface UnAuthorizedState {

}

class UnAuthorized extends React.Component<UnAuthorizedProps,UnAuthorizedState>{

    private siteURL =  this.props.spContext.webAbsoluteUrl;

    public componentDidMount(): void {
        hideLoader();
    }
    // public componentDidMount(): void {
    //     var sideNav:any = document.getElementsByClassName("outer-sidebar");
    //     if( sideNav ){
    //         sideNav[0].style.display = "none";
    //     }
    // }

    public render() {

        return(
            <div className="outer-Unauthorized">
                <div className="inner-Unauthorized">
                    <div className="row align-items-center">
                        <div className="col-md-4 text-right">
                            <div className="hand">
                                <FontAwesomeIcon icon={faHandPaper}/>
                            </div>
                        </div>
                        <div className="col-md-8">
                            <h2 style={{ color: '#df5556'}}>Access Denied</h2>
                            <p style={{fontSize: "21px"}}>You dont have Access to this Page.</p>
                            <p><a href={this.siteURL}>Click Here</a> to navigate to Dashboard</p>
                        </div>
                    </div>
                </div>
            </div>
        )
    }
}

export default UnAuthorized;