import * as React from 'react';
import type { IMaycoSafetyProps } from './IMaycoSafetyProps';
import { HashRouter } from 'react-router-dom';
import RoutesItems from './Navigation/Routesitems';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { hideLoader, showLoader } from './Shared/Loader';
import { ActionStatus } from './Constants/Contants';
import { showToast } from './Shared/Toaster';
import { ToastContainer } from 'react-toastify';
import { forEach } from 'lodash';
import NavBar from './Navigation/NavBar.components';

export default class MaycoSafety extends React.Component<IMaycoSafetyProps> {

  public state = {
    isAuthorized: false,
    currentUserGroups: [],
    siteURL: '',
    webAbsoluteURL: '',
    currPlantTitle:'',
    isSuperAdmin:false
  }
  private siteURL = this.props.spContext.siteAbsoluteUrl;

  public componentDidMount() {
    this.getUserRoles();
    this.removeExtraClasses();
  }

  private getUserRoles = async () => {
    try {
      showLoader();
      let currentUserGroupsList: any[] = [];
      let superAdminGrp = "WCM Super Admin";
      let isSuperAdmin = false;
      let siteURL = this.props.spContext.siteAbsoluteUrl;
      let webAbsoluteURL = this.props.spContext.webAbsoluteUrl;

      var webUrlSplit = webAbsoluteURL.split("/");
      let currPlantTitle = webUrlSplit[webUrlSplit.length-2];

      const spGroupsQuery = this.siteURL + "/_api/web/currentuser/groups";
      await this.props.spHttpClient.get( spGroupsQuery, SPHttpClient.configurations.v1 ).then( (res: SPHttpClientResponse) => {
        if(res.ok){
          res.json().then((resp)=>{
            let items = resp.value;
             for( let group of items ){
              if( group.Title == superAdminGrp ){ isSuperAdmin = true; }

              currentUserGroupsList.push(group.Title);
            }

            this.setState({
              isAuthorized: true,
              currentUserGroups: currentUserGroupsList,
              siteURL, webAbsoluteURL, currPlantTitle, isSuperAdmin
            });
          });  
        }
        else{
          console.log("Something went wrong while fetching user groups");
        }
      });
    } catch (e) {
      console.log(e);
      this.onError();
    }
  }

    private removeExtraClasses(){
  var workbenchElement = document.getElementById("workbenchPageContent");
  let wbClass = workbenchElement?.classList.value;
  workbenchElement?.classList.remove(wbClass? wbClass : "");
  // this.removeAll();
  workbenchElement?.addEventListener("click", this.removeAll);

}

private removeAll = () => {
  var workbenchElement = document.getElementById("workbenchPageContent");
  workbenchElement?.removeEventListener("click",this.removeAll);

  var canvasComponent1 = document.getElementsByClassName("CanvasZoneContainer");
  forEach( canvasComponent1, element => {
    let eleClass = element.classList.value;
    let eleClassArr = eleClass.split(" ");

    eleClassArr.forEach((elem: string) => {
      element.classList.remove(elem.trim());
    });
  })
  var canvasComponent1 = document.getElementsByClassName("CanvasZone");
  forEach( canvasComponent1, element => {
    let eleClass = element.classList.value;
   let eleClassArr = eleClass.split(" ");

    eleClassArr.forEach((elem: string) => {
      element.classList.remove(elem.trim());
    });
  })
  var canvasComponent1 = document.getElementsByClassName("CanvasSection");
  forEach( canvasComponent1, element => {
    let eleClass = element.classList.value;
   let eleClassArr = eleClass.split(" ");

    eleClassArr.forEach((elem: string) => {
      element.classList.remove(elem.trim());
    });
  })
}

  private onError = () => {
    showToast("error", ActionStatus.Error );
    hideLoader();
  }
  
  public render(){
    return(
      <React.Fragment>
        <ToastContainer />
        <HashRouter>
          <div className='menu-hide wrapper d-flex align-items-stretch' id="sideMenuNav">
            { this.state.isAuthorized ? <NavBar {...this.props} {...this.state} />: null}
            { this.state.isAuthorized ? <RoutesItems {...this.props} {...this.state} />: null}
          </div>
        </HashRouter>
      </React.Fragment>
    )
  }
}

