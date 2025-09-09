import * as React from "react";
import { Route,Routes, useParams } from "react-router-dom";
import { Suspense } from "react";
import Home from "../Home/Home.component";
import Actions from "../Masters/Actions";
import SecondaryRootCauses from "../Masters/SecondaryRootCauses";
import MicroRootCauses from "../Masters/MicroRootCauses";
import InjuryTypes from "../Masters/InjuryTypes";
import Status from "../Masters/Status";
import  AuditCategories from "../Masters/AuditCategories";
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
// import UnAuthorized from "../Unauthorized/Unauthorized.component";


export interface RoutesProps {
    isSupplierTeam: boolean,
    isDTETeam: boolean,
    isProcurementTeam: boolean,
    isAuthorized: boolean,
    spContext: any;
}

export interface RoutesState {
}

class RoutesItems extends React.Component<RoutesProps,RoutesState> {

    public render() {
        const WrapperHome = (props:any,component:Comment) => {
            let params = useParams();
            return <Home {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        // Masters
        const WrapperActions = (props:any) => {
            let params = useParams();
            return <Actions {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
         const WrapperSecondaryRootCauses = (props:any) => {
            let params = useParams();
            return <SecondaryRootCauses {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperInjuryTypes = (props:any) => {
            let params = useParams();
            return <InjuryTypes {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperMicroRootCauses = (props:any) => {
            let params = useParams();
            return <MicroRootCauses {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperStatus = (props:any) => {
            let params = useParams();
            return <Status {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperAuditCategories = (props:any) => {
            let params = useParams();
            return <AuditCategories {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperSMATandEHSMapping = (props:any) => {
            let params = useParams();
            return <SMATandEHSMapping {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperJSRACategories = (props:any) => {
            let params = useParams();
            return <JSRACategories {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperJSRASubCategories = (props:any) => {
            let params = useParams();
            return <JSRASubCategories {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperPPETypes = (props:any) => {
            let params = useParams();
            return <PPETypes {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperJSRADetails = (props:any) => {
            let params = useParams();
            return <JSRADetails {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperUATypes = (props:any) => {
            let params = useParams();
            return <UATypes {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperUASubTypes = (props:any) => {
            let params = useParams();
            return <UASubTypes {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        // Forms
        const WrapperSEWOForm = (props:any) => {
            let params = useParams();
            return <SEWOForm {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperUCANForm = (props:any) => {
            let params = useParams();
            return <UCANForm {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperSMATForm = (props:any) => {
            let params = useParams();
            return <SMATForm {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperEHSForm = (props:any) => {
            let params = useParams();
            return <EHSForm {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperJSRAForm = (props:any) => {
            let params = useParams();
            return <JSRAForm {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperTAGForm = (props:any) => {
            let params = useParams();
            return <TAGForm {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperCHECKLISTSTEP1Form = (props:any) => {
            let params = useParams();
            return <CHECKLISTSTEP1Form {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperCHECKLISTSTEP2Form = (props:any) => {
            let params = useParams();
            return <CHECKLISTSTEP2Form {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        const WrapperCHECKLISTSTEP3Form = (props:any) => {
            let params = useParams();
            return <CHECKLISTSTEP3Form {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
       
        return(
            <Suspense fallback={<div></div>}>
                <Routes>
                    {/* <Route path="/" element={ this.props.isAuthorized ? <WrapperHome /> : <UnAuthorized {...this.props} /> }/>
                    <Route path="/Home" element={ this.props.isAuthorized ? <WrapperHome /> : <UnAuthorized {...this.props} /> }/> */}
                    <Route path="/" element={  <WrapperHome /> }/>
                    <Route path="/Home" element={ <WrapperHome />}/>
                    {/* Masters */}
                    <Route path="/Actions" element={ <WrapperActions />}/>
                    <Route path="/SecondaryRootCauses" element={ <WrapperSecondaryRootCauses />}/>
                    <Route path="/MicroRootCauses" element={ <WrapperMicroRootCauses />}/>
                    <Route path="/InjuryTypes" element={ <WrapperInjuryTypes />}/>
                    <Route path="/Status" element={ <WrapperStatus />}/>
                    <Route path="/AuditCategories" element={ <WrapperAuditCategories />}/>
                    <Route path="/SMATandEHSMapping" element={ <WrapperSMATandEHSMapping />}/>
                    <Route path="/JSRACategories" element={ <WrapperJSRACategories />}/>
                    <Route path="/JSRASub-Categories" element={ <WrapperJSRASubCategories />}/>
                    <Route path="/PPETypes" element={ <WrapperPPETypes />}/>
                    <Route path="/JSRADetails" element={ <WrapperJSRADetails />}/>
                    <Route path="/UATypes" element={ <WrapperUATypes />}/>
                    <Route path="/UASub-Types" element={ <WrapperUASubTypes />}/>
                    {/* Forms */}
                    <Route path="/SEWOForm/:id?" element={ <WrapperSEWOForm />}/>
                    <Route path="/UCANForm/:id?" element={ <WrapperUCANForm />}/>
                    <Route path="/SMATForm/:id?" element={ <WrapperSMATForm />}/>
                    <Route path="/EHSForm/:id?" element={ <WrapperEHSForm />}/>
                    <Route path="/JSRAForm/:id?" element={ <WrapperJSRAForm />}/>
                    <Route path="/TAGForm/:id?" element={ <WrapperTAGForm />}/>
                    <Route path="/CHECK-LISTSTEP1Form/:id?" element={ <WrapperCHECKLISTSTEP1Form />}/>
                    <Route path="/CHECK-LISTSTEP2Form/:id?" element={ <WrapperCHECKLISTSTEP2Form />}/>
                    <Route path="/CHECK-LISTSTEP3Form/:id?" element={ <WrapperCHECKLISTSTEP3Form />}/>
                    {/* Views */}
                </Routes>
            </Suspense>
        )
    }

}

export default RoutesItems;