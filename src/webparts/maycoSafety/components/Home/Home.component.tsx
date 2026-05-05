import * as React from 'react';
import {useEffect} from 'react';
import { SPHttpClient } from "@microsoft/sp-http";
import "../CSS/Home.css";
import {hideLoader } from "../Shared/Loader";

export interface HomeProps{
    match:any;
    spContext:any;
    spHttpClient: SPHttpClient;
    context: any;
}
// export interface HomeState{
//     redirectToHome: boolean,
// }
const Home =(Props:HomeProps)=>{
 useEffect(() => {
    hideLoader();
  },[]);
                return(
                    <React.Fragment>
                          <div className="background">
                            <div className="overlay-color"></div>
                            <div className="overlay-box">
                                <img src={`${Props.spContext.webAbsoluteUrl}/PublishingImages/SitePages/Home/qa_dashboard.jpg`} alt="qa_dashboard.jpg" className="logo-image" />
                            </div>
                        </div>
                    </React.Fragment>
                )

}

export default Home;