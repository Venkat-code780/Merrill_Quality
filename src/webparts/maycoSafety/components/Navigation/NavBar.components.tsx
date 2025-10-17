import { faBars, faChevronDown, faChevronUp, faHome, faCogs, faFileAlt, faEye } from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import * as React from "react";
import { NavigateFunction, NavLink, Params} from "react-router-dom";
import "../CSS/left-nav.css";
import { withRouter } from "./withRouter";


export interface NavBarProps {
    currentUserGroups: any;
    isAuthorized: boolean;
    router:{
    location: Location;
    navigate: NavigateFunction;
    params: Params;
    }
    
}

export interface NavBarState {
    currentUserLinks: string[];
    showSidebar: boolean;
    openSideBars: { [key: string]: boolean };
    openMasters: { [key: string]: boolean };
     activeRoute: string | null;
}

class NavBar extends React.Component<NavBarProps, NavBarState> {

    public state: NavBarState = {
        currentUserLinks: [],
        showSidebar: false,
        openSideBars: {},
        openMasters: {},
        activeRoute: null
    };

    componentDidMount() {
        this.setAccordionFromPath();
    }

    public componentDidUpdate(prevProps: NavBarProps) {
  if (prevProps.router.location.pathname !== this.props.router.location.pathname) {
    this.setAccordionFromPath();
  }
}
   


     private setAccordionFromPath = () => {
     let path = this.props.router.location.pathname.toLowerCase();
          const openSideBars: { [key: string]: boolean } = {
              Masters: false,
               Forms: false,
                Views: false,
               Home: false
};
    // const openSideBars: any = {};
    const openMasters: any = {};
    let activeRoute: string | null = null;
        
     const basePath = '/' + path.split('/')[1]; // normalize dynamic segments
  activeRoute = basePath;
    // Handle Masters
    if (
      /^\/(actions|secondaryrootcauses|microrootcauses|injurytypes|status|auditcategories|smatandehsmapping|jsracategories|jsrasub-categories|ppetypes|jsradetails|uatypes|uasub-types)/.test(path)
    ) {
      openSideBars["Masters"] = true;

      // Detect which master section to open
      if (/actions|secondaryrootcauses|microrootcauses|injurytypes|status/.test(path))
        openMasters["SEWO"] = true;
      else if (/auditcategories|smatandehsmapping/.test(path))
        openMasters["SMAT/EHS"] = true;
      else if (/jsracategories|jsrasub-categories|ppetypes|jsradetails/.test(path))
        openMasters["JSRA"] = true;
      else if (/uatypes|uasub-types/.test(path))
        openMasters["Unsafe Act"] = true;
    }

    // Handle Forms
    else if (
      /^\/(sewoform|ucanform|smatform|ehsform|jsraform|tagform|check-liststep1form|check-liststep2form|check-liststep3form)/.test(
        path
      )
    ) {
      openSideBars["Forms"] = true;
    }

    // Handle Views
    else if (
      /^\/(sewoview|ucanview|smatview|ehsview|jsraview|tagview|check-liststep1view|check-liststep2view|check-liststep3view)/.test(
        path
      )
    ) {
      openSideBars["Views"] = true;
    }

    // Home
    else if (/^\/home/.test(path)) {
      openSideBars["Home"] = true;
    }

    this.setState({ openSideBars, openMasters, activeRoute})
    }
    public onNavItemClick(event: any) {
        let navLinks = document.querySelectorAll('.nav-click');
        if (navLinks.length > 0) {
            navLinks.forEach(item => {
                item.className = '';
            });
        }
        event.currentTarget.className = 'nav-click';
    }

    private toggleSidebar = () => {
        let prevShowSidebar = this.state.showSidebar;
        this.setState({ showSidebar: !prevShowSidebar })
        if (!prevShowSidebar) {
            document.getElementById('sideMenuNav')?.classList.add('active-navbar');
        } else {
            document.getElementById('sideMenuNav')?.classList.remove('active-navbar');
        }
    }

    private toggleSideBarItem = (event: any, title: string) => {
        event.preventDefault();
        let navLinks = document.querySelectorAll('.sidebar-title');
        if (navLinks.length > 0) {
            navLinks.forEach(item => {
                item.className = 'sidebar-title';
            });
        }
        event.currentTarget.className = 'sidebar-title left-nav-active';

        let newOpenSideBars: { [key: string]: boolean } = {};
        let prevOpenSideBars = { ...this.state.openSideBars };
        // Iterate over the previous open state openSideBars
        for (const key in prevOpenSideBars) {
            if (prevOpenSideBars.hasOwnProperty(key)) {
                if (key !== title) {
                    newOpenSideBars[key] = false;
                } else {
                    newOpenSideBars[key] = !prevOpenSideBars[title];
                }
            }
        }
        if (!prevOpenSideBars.hasOwnProperty(title)) {
            newOpenSideBars[title] = true; // Open it if it was not present before
        }
        this.setState({ openSideBars: newOpenSideBars });
        // this.setState((prevState: any) => ({
        //     openSideBars: {
        //         ...prevState.openSideBars,
        //         [title]: !prevState.openSideBars[title]
        //     }
        // }));
    }
    private toggleMasterItem = (title: string) => {
        let newopenMasters: { [key: string]: boolean } = {};
        let prevopenMasters = { ...this.state.openMasters };
        // Iterate over the previous open state openMasters
        for (const key in prevopenMasters) {
            if (prevopenMasters.hasOwnProperty(key)) {
                if (key !== title) {
                    newopenMasters[key] = false;
                } else {
                    newopenMasters[key] = !prevopenMasters[title];
                }
            }
        }
        if (!prevopenMasters.hasOwnProperty(title)) {
            newopenMasters[title] = true; // Open it if it was not present before
        }
        this.setState({ openMasters: newopenMasters });
        // this.setState((prevState: any) => ({
        //     openMasters: {
        //         ...prevState.openMasters,
        //         [title]: !prevState.openMasters[title]
        //     }
        // }));
    };

    public render() {
        const MasterTitles = ['SEWO', 'SMAT/EHS', 'JSRA', 'Unsafe Act'] as const;
        type MasterTitle = typeof MasterTitles[number];
        const MasterTitlesSubs: Record<MasterTitle, string[]> = {
            'SEWO': ['Actions', 'Secondary Root Causes', 'Micro Root Causes', 'Injury Types', 'Status'],
            'SMAT/EHS': ['Audit Categories', 'SMAT and EHS Mapping'],
            'JSRA': ['JSRA Categories', 'JSRA Sub - Categories', 'PPE Types', 'JSRA Details'],
            'Unsafe Act': ['UA Types', 'UA Sub - Types']
        };
        const FormAndViewTitles = ['SEWO', 'UCAN', 'SMAT', 'EHS', 'JSRA', 'TAG', 'CHECK - LIST STEP 1', 'CHECK - LIST STEP 2', 'CHECK - LIST STEP 3'];

        
        return (
            <React.Fragment>


                <div className="brd-left-nav" id='Left-Nav-Bar'>
                    <span className={`click-nav-icon`}>
                        <FontAwesomeIcon icon={faBars} onClick={this.toggleSidebar}></FontAwesomeIcon>
                    </span>
                    {/* //   <div className="outer-sidebar"> */}
                    {(this.state.showSidebar) &&
                        <div className="sidebar">
                            <div>
                                <ul className="list-unstyled ul-leftnav components mb-5">
                                    <li id="liHome" className="mb-1">
                                        <div className={`sidebar-title ${this.state.openSideBars["Home"] ? 'left-nav-active' : ''}`} onClick={(e) => this.toggleSideBarItem(e, 'Home')}>
                                            <NavLink to={"/Home"}><span><FontAwesomeIcon icon={faHome} /> Home</span></NavLink>
                                        </div>
                                    </li>
                                    {/* Masters Section */}
                                    <li className="liMasters mb-1">
                                   

                                        <div className={`sidebar-title ${
                                         this.state.openSideBars['Masters'] ||
                                          MasterTitles.some(title =>
                                          MasterTitlesSubs[title].some(sub =>
                                          this.state.activeRoute === `/${sub.replace(/\s+/g, '').toLowerCase()}`
                                              )
                                              ) ? 'left-nav-active': ''}`} onClick={(e) => this.toggleSideBarItem(e, 'Masters')}>
                                            <span><FontAwesomeIcon icon={faCogs}></FontAwesomeIcon>
                                                Masters
                                            </span>
                                            <span className="icon-down">
                                                <FontAwesomeIcon icon={this.state.openSideBars['Masters'] ? faChevronUp : faChevronDown} />
                                            </span>
                                        </div>

                                        {/* Loop Masters if open */}
                                        {this.state.openSideBars['Masters'] && MasterTitles.map((title: MasterTitle) => (
                                            <div key={title}>
                                                {/* Master Title */}
                                                <div className="master-sidebar-title" onClick={() => this.toggleMasterItem(title)}>
                                                    {title}
                                                    <span className="icon-down">
                                                        <FontAwesomeIcon icon={this.state.openMasters[title] ? faChevronUp : faChevronDown} />
                                                    </span>
                                                </div>

                                                {/* Sub-items */}
                                                {this.state.openMasters[title] && (
                                                    <ul className="ul-leftnav">
                                                        {MasterTitlesSubs[title].map((subItem: string) => (<li id={`li${subItem.replace(/\s+/g, '')}`} key={subItem} className={this.state.activeRoute === `/${subItem.replace(/\s+/g, '').toLowerCase()}` ? 'nav-click' : ''} >
                                                        <NavLink to={`/${subItem.replace(/\s+/g, '')}`}>
                                                        <span>{subItem}</span>
                                                          </NavLink>
                                                           </li>

                                                            // <li id={`li${subItem.replace(/\s+/g, '')}`} key={subItem} onClick={(event) => this.onNavItemClick(event)}>
                                                            //     <NavLink to={`/${subItem.replace(/\s+/g, '')}`}>
                                                            //         <span>{subItem}</span>
                                                            //     </NavLink>
                                                            // </li>
                                                        ))}
                                                    </ul>
                                                )}
                                            </div>
                                        ))}
                                    </li>
                                    {/* Forms Section */}
                                    <li className="liForms mb-1">

                                        <div className={`sidebar-title ${this.state.openSideBars['Forms'] ||MasterTitles.some(title => FormAndViewTitles.some(sub =>this.state.activeRoute === `/${sub.replace(/\s+/g, '').toLowerCase()}`) )? 'left-nav-active': ''}`} onClick={(e) => this.toggleSideBarItem(e, 'Forms')}>
                                            <span><FontAwesomeIcon icon={faFileAlt}></FontAwesomeIcon>
                                                Forms
                                            </span>
                                            <span className="icon-down">
                                                <FontAwesomeIcon icon={this.state.openSideBars['Forms'] ? faChevronUp : faChevronDown} />
                                            </span>
                                        </div>
                                        {this.state.openSideBars['Forms'] && (
                                            <ul className="ul-leftnav">
                                                {FormAndViewTitles.map((subItem: string) => (
                                                    <li id={`li${subItem.replace(/\s+/g, '')}Form`} key={subItem + 'Form'} className={this.state.activeRoute === `/${subItem.replace(/\s+/g, '').toLowerCase()}form` ? 'nav-click' : ''}>
                                            <NavLink to={`/${subItem.replace(/\s+/g, '')}Form`}>
                                            <span>{subItem}</span>
                                            </NavLink>
                                            </li>

                                                    // <li id={`li${subItem.replace(/\s+/g, '')}Form`} key={subItem + 'Form'} onClick={(event) => this.onNavItemClick(event)}>
                                                    //     <NavLink to={`/${subItem.replace(/\s+/g, '')}Form`}>
                                                    //         <span>{subItem}</span>
                                                    //     </NavLink>
                                                    // </li>
                                                ))}
                                            </ul>
                                        )}
                                    </li>
                                    {/* Views Section */}
                                    <li className="liViews mb-1">
                                        <div className={`sidebar-title  ${this.state.openSideBars['Views'] ? 'left-nav-active' : ''}`} onClick={(e) => this.toggleSideBarItem(e, 'Views')}>
                                            <span><FontAwesomeIcon icon={faEye}></FontAwesomeIcon>
                                                Views
                                            </span>
                                            <span className="icon-down">
                                                <FontAwesomeIcon icon={this.state.openSideBars['Views'] ? faChevronUp : faChevronDown} />
                                            </span>
                                        </div>
                                        {this.state.openSideBars['Views'] && (
                                            <ul className="ul-leftnav">
                                                {FormAndViewTitles.map((subItem: string) => (<li id={`li${subItem.replace(/\s+/g, '')}View`} key={subItem + 'View'} className={this.state.activeRoute === `/${subItem.replace(/\s+/g, '').toLowerCase()}view` ? 'nav-click' : ''}>
                                                <NavLink to={`/${subItem.replace(/\s+/g, '')}View`}>
                                                <span>{subItem}</span>
                                                </NavLink>
                                                   </li>

                                                    // <li id={`li${subItem.replace(/\s+/g, '')}View`} key={subItem + 'View'} onClick={(event) => this.onNavItemClick(event)}>
                                                    //     <NavLink to={`/${subItem.replace(/\s+/g, '')}View`}>
                                                    //         <span>{subItem}</span>
                                                    //     </NavLink>
                                                    // </li>
                                                ))}
                                            </ul>
                                        )}
                                    </li>

                                </ul>
                            </div>
                        </div>}
                </div>
            </React.Fragment >
        )
    }
}

export default withRouter(NavBar) ;
