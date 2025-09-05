import { faAddressCard, faBars, faBullhorn, faChevronDown, faChevronUp, faHome } from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import * as React from "react";
import { NavLink } from "react-router-dom";

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

    public componentDidMount() {
    }
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
        return(
            <React.Fragment>
                
                  
                  <div className="brd-left-nav" id='Left-Nav-Bar'>
                    <span className='click-nav-icon'>
                        <FontAwesomeIcon icon={faBars} onClick={this.toggleSidebar}></FontAwesomeIcon>
                    </span>
                {/* //   <div className="outer-sidebar"> */}
                    { ( this.state.showSidebar && this.props.isAuthorized) &&
                     <div className="sidebar">
                        <div>
                            <ul className="list-unstyled ul-leftnav components mb-5">
                                <li id="liHome" onClick={(event) => this.onNavItemClick(event)}>
                                    <div className="sidebar-title a-home"> 
                                        <NavLink to={"/Home"}><span><FontAwesomeIcon icon={faHome}/> Home</span></NavLink>
                                    </div>
                                </li>
                                 { (this.props.isSupplierTeam ||  this.props.isProcurementTeam) && 
                                    <li className="liMasters mb-2">
                                        <div className="sidebar-title" onClick={this.toggleMastersMenu}>
                                            Masters 
                                            <span className="icon-down">
                                                <FontAwesomeIcon icon={ this.state.isMastersOpen ? faChevronUp : faChevronDown }></FontAwesomeIcon>
                                            </span>
                                        </div>
                                        { this.state.isMastersOpen && (<ul className="ul-leftnav">

                                            <li id="liVendorProfile" onClick={(event) => this.onNavItemClick(event)}>
                                                <NavLink to={"/VendorProfile"}><span><FontAwesomeIcon icon={faAddressCard}/> Vendor Profile</span></NavLink>
                                            </li>
                                            <li id="liLeadSource" onClick={(event) => this.onNavItemClick(event)}>
                                                <NavLink to={"/LeadSource"}><span><FontAwesomeIcon icon={faBullhorn}/> Lead Source</span></NavLink>
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