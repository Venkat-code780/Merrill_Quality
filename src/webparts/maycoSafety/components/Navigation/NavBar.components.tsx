import { faAddressCard, faBars, faBullhorn, faChevronDown, faChevronUp, faHome } from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import * as React from "react";
import { NavLink } from "react-router-dom";
import "../CSS/left-nav.css";

export interface NavBarProps {
    currentUserGroups: any;
    isAuthorized: boolean;
    isSupplierTeam: boolean,
    isDTETeam: boolean,
    isProcurementTeam: boolean,
}

export interface NavBarState {
    currentUserLinks: string[];
    showSidebar: boolean;
    isMastersOpen: boolean;
}

class NavBar extends React.Component<NavBarProps, NavBarState> {

    public state : NavBarState = { 
        currentUserLinks : [],
        showSidebar: true,
        isMastersOpen: true
    };
    public onNavItemClick(event: any) {
        let navLinks = document.querySelectorAll('.nav-click');
        if( navLinks.length > 0 ){
            navLinks.forEach( item => {
                item.className = '';
            });
        }
        event.currentTarget.className = 'nav-click';
    }

    private toggleSidebar = () => {
        let prevShowSidebar = this.state.showSidebar;
        this.setState({ showSidebar: !prevShowSidebar})
    }

    private toggleMastersMenu = () => {
        this.setState({ isMastersOpen: !this.state.isMastersOpen });
    }

    public render() {
        let MasterTitles=['SEWO','SMAT/EHS','JSRA','Unsafe Act'];
        return(
            <React.Fragment>
                
                  
                  <div className="brd-left-nav" id='Left-Nav-Bar'>
                    <span className='click-nav-icon'>
                        <FontAwesomeIcon icon={faBars} onClick={this.toggleSidebar}></FontAwesomeIcon>
                    </span>
                {/* //   <div className="outer-sidebar"> */}
                    { ( this.state.showSidebar) &&
                     <div className="sidebar">
                        <div>
                            <ul className="list-unstyled ul-leftnav components mb-5">
                                <li id="liHome" onClick={(event) => this.onNavItemClick(event)}>
                                    <div className="sidebar-title a-home"> 
                                        <NavLink to={"/Home"}><span><FontAwesomeIcon icon={faHome}/> Home</span></NavLink>
                                    </div>
                                </li>
                                 {
                                    <li className="liMasters mb-2">
                                        <div className="sidebar-title" onClick={this.toggleMastersMenu}>
                                            Masters 
                                            <span className="icon-down">
                                                <FontAwesomeIcon icon={ this.state.isMastersOpen ? faChevronUp : faChevronDown }></FontAwesomeIcon>
                                            </span>
                                        </div>
                                        {MasterTitles.map(title=>
                                            (
                                            <div className="master-sidebar-title" onClick={this.toggleMastersMenu}>
                                            {title }
                                            <span className="icon-down">
                                                <FontAwesomeIcon icon={ this.state.isMastersOpen ? faChevronUp : faChevronDown }></FontAwesomeIcon>
                                            </span>
                                           </div>
                                            )
                                        )}
                                        
                                        { this.state.isMastersOpen && (<ul className="ul-leftnav">

                                            <li id="liActions" onClick={(event) => this.onNavItemClick(event)}>
                                                <NavLink to={"/Actions"}><span><FontAwesomeIcon icon={faAddressCard}/> Actions</span></NavLink>
                                            </li>
                                            <li id="liSecondaryRootCauses" onClick={(event) => this.onNavItemClick(event)}>
                                                <NavLink to={"/SecondaryRootCauses"}><span><FontAwesomeIcon icon={faBullhorn}/> Secondary Root Causes</span></NavLink>
                                            </li>
                                        </ul>)}
                                    </li>}
                            </ul>
                        </div>
                    </div>}
                </div>
            </React.Fragment>
        )
    }
}

export default NavBar;