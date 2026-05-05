import { faBars, faChevronDown, faChevronUp, faHome, faCogs, faFileAlt, faEye } from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import * as React from "react";
import { NavigateFunction, NavLink, Params } from "react-router-dom";
import "../CSS/left-nav.css";
import { withRouter } from "./withRouter";

export interface NavBarProps {
    currentUserGroups: any;
    isAuthorized: boolean;
    router: {
        location: Location;
        navigate: NavigateFunction;
        params: Params;
    }
}

export interface NavBarState {
    currentUserLinks: string[];
    showSidebar: boolean;
    openSideBars: { [key: string]: boolean };
    activeRoute: string | null;
}

class NavBar extends React.Component<NavBarProps, NavBarState> {

    public state: NavBarState = {
        currentUserLinks: [],
        showSidebar: true,
        openSideBars: {},
        activeRoute: null
    };

    componentDidMount() {
        document.getElementById('sideMenuNav')?.classList.add('active-navbar');
        this.setAccordionFromPath();
    }

    componentDidUpdate(prevProps: NavBarProps) {
        if (prevProps.router.location.pathname !== this.props.router.location.pathname) {
            this.setAccordionFromPath();
        }
    }

    private setAccordionFromPath = () => {

        let path = this.props.router.location.pathname.toLowerCase();

        const openSideBars: any = {
            Masters: false,
            Forms: false,
            Views: false,
            Home: false,
            Reports: false
        };

        const basePath = '/' + path.split('/')[1];
        let activeRoute = basePath;
        const masterRoutes = [
            "/auditcategories",
            "/kpi",
            "/lpaauditorlevels",
            "/lpaauditors",
            "/lpacategories",
            "/lpa"
        ];
        // Masters
        // if (
        //     /^\/(auditcategories|kpi|lpaauditorlevel|lpaauditors|lpacategories|lpa)/.test(path)
        // ) {
        //     openSideBars["Masters"] = true;
        // }

        // // Forms
        // else if (/form/.test(path)) {
        //     openSideBars["Forms"] = true;
        // }

        // // Views
        // else if (/view/.test(path)) {
        //     openSideBars["Views"] = true;
        // }

        // // Home
        // else if (/home/.test(path) || path === "/") {
        //     openSideBars["Home"] = true;
        // }
        // Home FIRST
        if (path === "/" || path === "/home") {
            openSideBars["Home"] = true;
        }

        // ✅ Forms FIRST
        else if (path.includes("form")) {
            openSideBars["Forms"] = true;
        }

        // ✅ Views SECOND
        else if (path.includes("view")) {
            openSideBars["Views"] = true;
        }
        else if (path.startsWith("/lpa-report")) {   // ✅ Reports here
            openSideBars["Reports"] = true;
        }

        // ✅ Masters LAST (fallback only)
        else if (masterRoutes.some(route => path.startsWith(route))) {
            openSideBars["Masters"] = true;
        }


        this.setState({ openSideBars, activeRoute });
    }

    private toggleSidebar = () => {

        let prevShowSidebar = this.state.showSidebar;

        this.setState({ showSidebar: !prevShowSidebar });

        if (!prevShowSidebar) {
            document.getElementById('sideMenuNav')?.classList.add('active-navbar');
        } else {
            document.getElementById('sideMenuNav')?.classList.remove('active-navbar');
        }
    }

    private toggleSideBarItem = (event: any, title: string) => {

        event.preventDefault();

        let prevOpenSideBars = { ...this.state.openSideBars };

        let newOpenSideBars: any = {};

        for (const key in prevOpenSideBars) {
            newOpenSideBars[key] = key === title ? !prevOpenSideBars[key] : false;
        }

        if (!prevOpenSideBars.hasOwnProperty(title)) {
            newOpenSideBars[title] = true;
        }

        this.setState({ openSideBars: newOpenSideBars });
    }

    public render() {

        const MasterLinks = [

            "Audit Categories",
            "KPI",
            "LPA Auditor Levels",
            "LPA Auditors",
            "LPA Categories",
            "LPA"

        ];

        const FormAndViewTitles = [
            "LPA",
            "QA-Matrix"

        ];
        const ReportTitles = [
            "LPA-Report"
        ];

        return (

            <div className="brd-left-nav" id='Left-Nav-Bar'>

                <span className="click-nav-icon" onClick={this.toggleSidebar}>
                    <FontAwesomeIcon icon={faBars} />
                </span>

                {this.state.showSidebar && (

                    <div className="sidebar">

                        <ul className="list-unstyled ul-leftnav components mb-5">

                            {/* Home */}

                            <li id="liHome">

                                <div
                                    className={`sidebar-title ${this.state.openSideBars["Home"] ? 'left-nav-active' : ''}`}
                                    onClick={(e) => this.toggleSideBarItem(e, 'Home')}
                                >
                                    <NavLink to="/Home">
                                        <FontAwesomeIcon icon={faHome} /> Home
                                    </NavLink>
                                </div>

                            </li>


                            {/* Masters */}

                            <li>

                                <div
                                    className={`sidebar-title ${this.state.openSideBars['Masters'] ? 'left-nav-active' : ''}`}
                                    onClick={(e) => this.toggleSideBarItem(e, 'Masters')}
                                >

                                    <span>
                                        <FontAwesomeIcon icon={faCogs} /> Masters
                                    </span>

                                    <span className="icon-down">
                                        <FontAwesomeIcon icon={this.state.openSideBars['Masters'] ? faChevronUp : faChevronDown} />
                                    </span>

                                </div>

                                {this.state.openSideBars['Masters'] && (

                                    <ul className="ul-leftnav">

                                        {MasterLinks.map((item) => {

                                            const route = "/" + item.replace(/\s+/g, "");

                                            return (

                                                <li
                                                    key={item}
                                                    className={this.state.activeRoute === route.toLowerCase() ? "nav-click" : ""}
                                                >

                                                    <NavLink to={route}>
                                                        {item}
                                                    </NavLink>

                                                </li>

                                            )

                                        })}

                                    </ul>

                                )}

                            </li>


                            {/* Forms */}

                            <li>

                                <div
                                    className={`sidebar-title ${this.state.openSideBars['Forms'] ? 'left-nav-active' : ''}`}
                                    onClick={(e) => this.toggleSideBarItem(e, 'Forms')}
                                >

                                    <span>
                                        <FontAwesomeIcon icon={faFileAlt} /> Forms
                                    </span>

                                    <span className="icon-down">
                                        <FontAwesomeIcon icon={this.state.openSideBars['Forms'] ? faChevronUp : faChevronDown} />
                                    </span>

                                </div>

                                {this.state.openSideBars['Forms'] && (

                                    <ul className="ul-leftnav">

                                        {FormAndViewTitles.map((item) => {

                                            const route = "/" + item + "Form";

                                            return (

                                                // <li key={route}>
                                                //     <NavLink to={route}>{item}</NavLink>
                                                // </li>
                                                <li
                                                    key={route}
                                                    className={this.state.activeRoute === route.toLowerCase() ? "nav-click" : ""}
                                                >
                                                    <NavLink to={route}>
                                                        {item}
                                                    </NavLink>
                                                </li>

                                            )

                                        })}

                                    </ul>

                                )}

                            </li>


                            {/* Views */}

                            <li>

                                <div
                                    className={`sidebar-title ${this.state.openSideBars['Views'] ? 'left-nav-active' : ''}`}
                                    onClick={(e) => this.toggleSideBarItem(e, 'Views')}
                                >

                                    <span>
                                        <FontAwesomeIcon icon={faEye} /> Views
                                    </span>

                                    <span className="icon-down">
                                        <FontAwesomeIcon icon={this.state.openSideBars['Views'] ? faChevronUp : faChevronDown} />
                                    </span>

                                </div>

                                {this.state.openSideBars['Views'] && (

                                    <ul className="ul-leftnav">

                                        {FormAndViewTitles.map((item) => {

                                            const route = "/" + item + "View";

                                            return (

                                                // <li key={route}>
                                                //     <NavLink to={route}>{item}</NavLink>
                                                // </li>
                                                <li
                                                    key={route}
                                                    className={this.state.activeRoute === route.toLowerCase() ? "nav-click" : ""}
                                                >
                                                    <NavLink to={route}>
                                                        {item}
                                                    </NavLink>
                                                </li>

                                            )

                                        })}

                                    </ul>

                                )}

                            </li>

                            <li>

                                <div
                                    className={`sidebar-title ${this.state.openSideBars['Reports'] ? 'left-nav-active' : ''}`}
                                    onClick={(e) => this.toggleSideBarItem(e, 'Reports')}
                                >

                                    <span>
                                        <FontAwesomeIcon icon={faFileAlt} /> Reports
                                    </span>

                                    <span className="icon-down">
                                        <FontAwesomeIcon
                                            icon={this.state.openSideBars['Reports'] ? faChevronUp : faChevronDown}
                                        />
                                    </span>

                                </div>

                                {this.state.openSideBars['Reports'] && (

                                    <ul className="ul-leftnav">

                                        {ReportTitles.map((item) => {

                                            const route = "/" + item;

                                            return (
                                                <li
                                                    key={route}
                                                    className={this.state.activeRoute === route.toLowerCase() ? "nav-click" : ""}
                                                >
                                                    <NavLink to={route}>
                                                        {item}
                                                    </NavLink>
                                                </li>
                                            );

                                        })}

                                    </ul>

                                )}

                            </li>

                        </ul>

                    </div>

                )}

            </div>
        )
    }
}

export default withRouter(NavBar);