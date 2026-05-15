import * as React from "react";
import { Route, Routes, useParams } from "react-router-dom";
import { Suspense } from "react";
import Home from "../Home/Home.component";
import AuditCategories from "../Masters/AuditCategories";
import KPI from "../Masters/KPI";
import LPAAuditorLevel from "../Masters/LpaAuditorLevels";
import LPAAuditors from "../Masters/LPAAuditors";
import LPACategories from "../Masters/LPACategories";
import LPA from "../Masters/LPA";
import LPAForm from "../Forms/LPA";
import LPAView from "../Views/LPAView";
import QAMatrixForm from "../Forms/QA-Matrix";
import QAMatrixView from "../Views/QAMatrixView";
import LPAReport from "../Reports/LPARports";



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
      
        const WrapperAuditcategory = (props: any) => {
            let params = useParams();
            return <AuditCategories {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }

         const WrapperKPI = (props: any) => {
            let params = useParams();
            return <KPI {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
           const WrapperLPAAuditorLevel = (props: any) => {
            let params = useParams();
            return <LPAAuditorLevel {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }

             const WrapperLPAAuditors = (props: any) => {
            let params = useParams();
            return <LPAAuditors {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
           const WrapperLPACategories = (props: any) => {
            let params = useParams();
            return <LPACategories {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
            const WrapperLPA = (props: any) => {
            let params = useParams();
            return <LPA {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
          const WrapperLPAForm = (props: any) => {
            let params = useParams();
            return <LPAForm {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
          const WrapperQAMatrixForm = (props: any) => {
            let params = useParams();
            return <QAMatrixForm {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }


       




            const WrapperLPAView = (props: any) => {
            let params = useParams();
            return <LPAView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
            const WrapperQAMatrixView = (props: any) => {
            let params = useParams();
            return <QAMatrixView {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
               const WrapperReports = (props: any) => {
            let params = useParams();
            return <LPAReport {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        
         
       
        // Forms



        return (
            <Suspense fallback={<div></div>}>
                <Routes>
                    <Route path="/" element={<WrapperHome />} />
                    <Route path="/Home" element={<WrapperHome />} />
                    {/* Masters */}
           
                    <Route path="/AuditCategories" element={<WrapperAuditcategory />} />
                    <Route path="/KPI" element={<WrapperKPI></WrapperKPI>}></Route>
                    <Route path="/LPAAuditorLevels" element={<WrapperLPAAuditorLevel></WrapperLPAAuditorLevel>}></Route>
                    <Route path="/LPAAuditors" element={<WrapperLPAAuditors></WrapperLPAAuditors>}></Route>
                    <Route path="/LPACategories" element={<WrapperLPACategories></WrapperLPACategories>}></Route>
                    <Route path="/LPA" element={<WrapperLPA></WrapperLPA>}></Route>
                    {/* Forms */}
                   <Route path="/LPAForm/:id?" element={<WrapperLPAForm />} />
                    <Route path="/QA-MatrixForm/:id?" element={<WrapperQAMatrixForm></WrapperQAMatrixForm>} />
                    {/* Views */}
                    <Route path="/LPAView" element={<WrapperLPAView></WrapperLPAView>}></Route>
                    <Route path="/QA-MatrixView" element={<WrapperQAMatrixView></WrapperQAMatrixView>}></Route>
                      
                        <Route path="/LPA-Report" element={<WrapperReports></WrapperReports>}></Route>
            

                </Routes>
            </Suspense>
        )
    }

}

export default RoutesItems;