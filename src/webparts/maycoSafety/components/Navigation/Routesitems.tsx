import * as React from "react";
import { Route,Routes, useParams } from "react-router-dom";
import { Suspense } from "react";
import Home from "../Home/Home.component";
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
        const WrapperHome = (props:any) => {
            let params = useParams();
            return <Home {...this.context} {...this.props} {...{...props, match:{params}}}  />
        }
        return(
            <Suspense fallback={<div></div>}>
                <Routes>
                    {/* <Route path="/" element={ this.props.isAuthorized ? <WrapperHome /> : <UnAuthorized {...this.props} /> }/>
                    <Route path="/Home" element={ this.props.isAuthorized ? <WrapperHome /> : <UnAuthorized {...this.props} /> }/> */}
                    <Route path="/" element={  <WrapperHome /> }/>
                    <Route path="/Home" element={ <WrapperHome />}/>
                </Routes>
            </Suspense>
        )
    }

}

export default RoutesItems;