import * as React from 'react';
import type { IMaycoSafetyProps } from './IMaycoSafetyProps';
import { HashRouter } from 'react-router-dom';
import RoutesItems from './Navigation/Routesitems';
import { SPHttpClient } from '@microsoft/sp-http';
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
    currPlantTitle: '',
    isWCM: false,
    FormAccessConfiguration: {}
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
      let siteURL = this.props.spContext.siteAbsoluteUrl;
      let webAbsoluteURL = this.props.spContext.webAbsoluteUrl;

      var webUrlSplit = webAbsoluteURL.split("/");
      let currPlantTitle = webUrlSplit[webUrlSplit.length - 2];
      let currEnvironment = webUrlSplit[webUrlSplit.length - 4];
      console.log("Environment: " + currEnvironment);
      let isWCM = false;


      const spGroupsQuery = this.siteURL + "/_api/web/currentuser/groups";
      const formAccessQuery = `${siteURL}/mayco/_api/web/lists/getbytitle('SafetyFormAccessConfiguration')/items?$expand=AccessGroup,Plant&$select=*,AccessGroup/Id,AccessGroup/Title,Plant/Id,Plant/Title`;
      let [res, formAccess] = await Promise.all([this.props.spHttpClient.get(spGroupsQuery, SPHttpClient.configurations.v1), this.props.spHttpClient.get(formAccessQuery, SPHttpClient.configurations.v1)]);
      if (res.ok && formAccess.ok) {
        await res.json().then((resp: any) => {
          let items = resp.value;
          for (let group of items) {
            currentUserGroupsList.push(group.Title);
          }
          if (currEnvironment.toLowerCase() == "wcm") { isWCM = true };
        });

        let formAccessGroups: any = {};
        await formAccess.json().then((access: any) => {
          let accessConfig = access.value;
          console.log(accessConfig);
          accessConfig.forEach((item: any) => {
            if (item.AccessGroup && Array.isArray(item.AccessGroup) && item.Form && item.Plant.Title.toLowerCase() == currPlantTitle ) {
              let groupTitles = item.AccessGroup.map((grp: any) => grp.Title);
              formAccessGroups[item.Form] = groupTitles;
            }
          });
          console.log(formAccessGroups);
        })

        this.setState({
          isAuthorized: true,
          currentUserGroups: currentUserGroupsList,
          siteURL, webAbsoluteURL, currPlantTitle, isWCM,
          FormAccessConfiguration: formAccessGroups
        });
      }
      else {
        if (!res.ok) {
          throw new Error("Something went wrong while fetching user groups");
        }
        if (!formAccess.ok) {
          throw new Error("Error retrieving SafetyFormAccessConfiguration list.");
        }
        // console.log("Something went wrong while fetching user groups");
      }
      // });
    } catch (e) {
      console.log(e);
      this.onError();
    }
  }

  private removeExtraClasses() {
    var workbenchElement = document.getElementById("workbenchPageContent");
    let wbClass = workbenchElement?.classList.value;
    workbenchElement?.classList.remove(wbClass ? wbClass : "");
    // this.removeAll();
    workbenchElement?.addEventListener("click", this.removeAll);

  }

  private removeAll = () => {
    var workbenchElement = document.getElementById("workbenchPageContent");
    workbenchElement?.removeEventListener("click", this.removeAll);

    var canvasComponent1 = document.getElementsByClassName("CanvasZoneContainer");
    forEach(canvasComponent1, element => {
      let eleClass = element.classList.value;
      let eleClassArr = eleClass.split(" ");

      eleClassArr.forEach((elem: string) => {
        element.classList.remove(elem.trim());
      });
    })
    var canvasComponent1 = document.getElementsByClassName("CanvasZone");
    forEach(canvasComponent1, element => {
      let eleClass = element.classList.value;
      let eleClassArr = eleClass.split(" ");

      eleClassArr.forEach((elem: string) => {
        element.classList.remove(elem.trim());
      });
    })
    var canvasComponent1 = document.getElementsByClassName("CanvasSection");
    forEach(canvasComponent1, element => {
      let eleClass = element.classList.value;
      let eleClassArr = eleClass.split(" ");

      eleClassArr.forEach((elem: string) => {
        element.classList.remove(elem.trim());
      });
    })
  }

  private onError = () => {
    showToast("error", ActionStatus.Error);
    hideLoader();
  }

  public render() {
    return (
      <React.Fragment>
        <ToastContainer />
        <HashRouter>
          <div className='menu-hide wrapper d-flex align-items-stretch' id="sideMenuNav">
            {this.state.isAuthorized ? <NavBar {...this.props} {...this.state} /> : null}
            {this.state.isAuthorized ? <RoutesItems {...this.props} {...this.state} /> : null}
          </div>
        </HashRouter>
      </React.Fragment>
    )
  }
}

