import * as React from 'react';
//import DateUtilities from "../Utilities/DateUtilities";
import { format } from "date-fns";

interface ActionHistoryProps {
    HeaderData: any;
    HistoryData: any;
    spContext: any;
}

const ActionHistory = ({ HeaderData, HistoryData, spContext }: ActionHistoryProps) => {
    HistoryData=HistoryData.slice().reverse();
    return <div className='divActionHistory'><table className="table-bordered border  col-md-12">
        <thead className='theadActionHistory'>
            <tr>
                {HeaderData.map((header: any, index: number) => (
                    <th key={index}>{header}</th>
                ))}
            </tr>
        </thead>
        <tbody>
            {HistoryData.map((action: any, index: number) => (
                <tr key={index}>
                    <td>{action.ActionBy}</td>
                    <td>{format(new Date(action.ActionDateTime), "MM/dd/yyyy hh:mm aa")}</td>
                </tr>
            ))}
        </tbody>
    </table>
    </div>;
};

export default ActionHistory;