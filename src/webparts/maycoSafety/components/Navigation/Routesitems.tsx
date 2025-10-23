import * as React from "react";
import { Route, Routes, useLocation, useParams } from "react-router-dom";
import { Suspense } from "react";
import Home from "../Home/Home.component";
import Actions from "../Masters/Actions";
import SecondaryRootCauses from "../Masters/SecondaryRootCauses";
import MicroRootCauses from "../Masters/MicroRootCauses";
import InjuryTypes from "../Masters/InjuryTypes";
import Status from "../Masters/Status";
import AuditCategories from "../Masters/AuditCategories";
import SMATandEHSMapping from "../Masters/SMATandEHSMapping";
import JSRACategories from "../Masters/JSRACategories";
import JSRASubCategories from "../Masters/JSRASub-Categories";
import PPETypes from "../Masters/PPETypes";
import JSRADetails from "../Masters/JSRADetails";
import UATypes from "../Masters/UATypes";
import UASubTypes from "../Masters/UASub-Types";
import SEWOForm from "../Forms/SEWOForm";
import UCANForm from "../Forms/UCANForm";
import SMATForm from "../Forms/SMATForm";
import EHSForm from "../Forms/EHSForm";
import JSRAForm from "../Forms/JSRAForm";
import TAGForm from "../Forms/TAGForm";
import CHECKLISTSTEP1Form from "../Forms/CHECK-LISTSTEP1Form";
import CHECKLISTSTEP2Form from "../Forms/CHECK-LISTSTEP2Form";
import CHECKLISTSTEP3Form from "../Forms/CHECK-LISTSTEP3Form";
import CHECKLISTSTEP1View from "../Views/CHECK-LISTSTEP1View";
import CHECKLISTSTEP2View from "../Views/CHECK-LISTSTEP2View";
import CHECKLISTSTEP3View from "../Views/CHECK-LISTSTEP3View"
import SEWOView from "../Views/SEWOView";
import UCANView from "../Views/UCANView";
import SMATView from "../Views/SMATView";
import EHSView from "../Views/EHSView";
import JSRAView from "../Views/JSRAView";
import TAGView from "../Views/TAGView";
// import UnAuthorized from "../Unauthorized/Unauthorized.component";


export interface RoutesProps {
    isAuthorized: boolean,
    spContext: any;
}

export interface RoutesState {
}

class RoutesItems extends React.Component<RoutesProps, RoutesState> {

    public render() {
        const WrapperHome = (props: any, component: Comment) => {
            let params = useParams();
            return <Home {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        // Masters
        const WrapperActions = (props: any) => {
            let params = useParams();
            return <Actions {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperSecondaryRootCauses = (props: any) => {
            let params = useParams();
            return <SecondaryRootCauses {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperInjuryTypes = (props: any) => {
            let params = useParams();
            return <InjuryTypes {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperMicroRootCauses = (props: any) => {
            let params = useParams();
            return <MicroRootCauses {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperStatus = (props: any) => {
            let params = useParams();
            return <Status {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperAuditCategories = (props: any) => {
            let params = useParams();
            return <AuditCategories {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperSMATandEHSMapping = (props: any) => {
            let params = useParams();
            return <SMATandEHSMapping {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperJSRACategories = (props: any) => {
            let params = useParams();
            return <JSRACategories {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperJSRASubCategories = (props: any) => {
            let params = useParams();
            return <JSRASubCategories {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperPPETypes = (props: any) => {
            let params = useParams();
            return <PPETypes {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperJSRADetails = (props: any) => {
            let params = useParams();
            return <JSRADetails {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperUATypes = (props: any) => {
            let params = useParams();
            return <UATypes {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperUASubTypes = (props: any) => {
            let params = useParams();
            return <UASubTypes {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        // Forms
        const WrapperSEWOForm = (props: any) => {
            const location = useLocation();
            const params = useParams();
            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <SEWOForm key={key} {...allProps} />;
        };
        const WrapperUCANForm = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <UCANForm key={key} {...allProps} />;
        };
        const WrapperSMATForm = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <SMATForm key={key} {...allProps} />;
        };
        const WrapperEHSForm = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <EHSForm key={key} {...allProps} />;
        };
        const WrapperJSRAForm = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <JSRAForm key={key} {...allProps} />;
        };
        const WrapperTAGForm = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <TAGForm key={key} {...allProps} />;
        };
        const WrapperCHECKLISTSTEP1Form = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <CHECKLISTSTEP1Form key={key} {...allProps} />;
        };
        const WrapperCHECKLISTSTEP2Form = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <CHECKLISTSTEP2Form key={key} {...allProps} />;
        };
        const WrapperCHECKLISTSTEP3Form = (props: any) => {
            const location = useLocation();
            const params = useParams();

            const key = location.pathname + location.search;

            const allProps = {
                ...this.context,
                ...this.props,
                ...props,
                match: { params },
            };

            return <CHECKLISTSTEP3Form key={key} {...allProps} />;
        };
        //Views
        const WrapperCHECKLISTSTEP1View = (props: any) => {
            let params = useParams();
            return <CHECKLISTSTEP1View {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperCHECKLISTSTEP2View = (props: any) => {
            let params = useParams();
            return <CHECKLISTSTEP2View {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperCHECKLISTSTEP3View = (props: any) => {
            let params = useParams();
            return <CHECKLISTSTEP3View {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperSEWOView = (props: any) => {
            let params = useParams();
            return <SEWOView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperUCANView = (props: any) => {
            let params = useParams();
            return <UCANView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperSMATView = (props: any) => {
            let params = useParams();
            return <SMATView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperEHSView = (props: any) => {
            let params = useParams();
            return <EHSView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperJSRAView = (props: any) => {
            let params = useParams();
            return <JSRAView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperTAGView = (props: any) => {
            let params = useParams();
            return <TAGView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }


        return (
            <Suspense fallback={<div></div>}>
                <Routes>
                    <Route path="/" element={<WrapperHome />} />
                    <Route path="/Home" element={<WrapperHome />} />
                    {/* Masters */}
                    <Route path="/Actions" element={<WrapperActions />} />
                    <Route path="/SecondaryRootCauses" element={<WrapperSecondaryRootCauses />} />
                    <Route path="/MicroRootCauses" element={<WrapperMicroRootCauses />} />
                    <Route path="/InjuryTypes" element={<WrapperInjuryTypes />} />
                    <Route path="/Status" element={<WrapperStatus />} />
                    <Route path="/AuditCategories" element={<WrapperAuditCategories />} />
                    <Route path="/SMATandEHSMapping" element={<WrapperSMATandEHSMapping />} />
                    <Route path="/JSRACategories" element={<WrapperJSRACategories />} />
                    <Route path="/JSRASub-Categories" element={<WrapperJSRASubCategories />} />
                    <Route path="/PPETypes" element={<WrapperPPETypes />} />
                    <Route path="/JSRADetails" element={<WrapperJSRADetails />} />
                    <Route path="/UATypes" element={<WrapperUATypes />} />
                    <Route path="/UASub-Types" element={<WrapperUASubTypes />} />
                    {/* Forms */}
                    <Route path="/SEWOForm/:id?" element={<WrapperSEWOForm />} />
                    <Route path="/UCANForm/:id?" element={<WrapperUCANForm />} />
                    <Route path="/SMATForm/:id?" element={<WrapperSMATForm />} />
                    <Route path="/EHSForm/:id?" element={<WrapperEHSForm />} />
                    <Route path="/JSRAForm/:id?" element={<WrapperJSRAForm />} />
                    <Route path="/TAGForm/:id?" element={<WrapperTAGForm />} />
                    <Route path="/CHECK-LISTSTEP1Form/:id?" element={<WrapperCHECKLISTSTEP1Form />} />
                    <Route path="/CHECK-LISTSTEP2Form/:id?" element={<WrapperCHECKLISTSTEP2Form />} />
                    <Route path="/CHECK-LISTSTEP3Form/:id?" element={<WrapperCHECKLISTSTEP3Form />} />
                    {/* Views */}
                    <Route path="/CHECK-LISTSTEP1View" element={<WrapperCHECKLISTSTEP1View />} />
                    <Route path="/CHECK-LISTSTEP2View" element={<WrapperCHECKLISTSTEP2View />} />
                    <Route path="/CHECK-LISTSTEP3View" element={<WrapperCHECKLISTSTEP3View />} />
                    <Route path="/SEWOView" element={<WrapperSEWOView />} />
                    <Route path="/UCANView" element={<WrapperUCANView />} />
                    <Route path="/SMATView" element={<WrapperSMATView />} />
                    <Route path="/EHSView" element={<WrapperEHSView />} />
                    <Route path="/JSRAView" element={<WrapperJSRAView />} />
                    <Route path="/TAGView" element={<WrapperTAGView />} />





                </Routes>
            </Suspense>
        )
    }

}

export default RoutesItems;