import * as React from "react";
import { Route,Routes, useParams } from "react-router-dom";
import { Suspense } from "react";

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

        return(
            <Suspense fallback={<div></div>}>
                <Routes>
                </Routes>
            </Suspense>
        )
    }

}

export default RoutesItems;