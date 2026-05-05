import * as React from 'react';
interface InputTextProps {
    label: string;
    name: string;
    checked: boolean;
    onChange: any;
    isdisable:boolean;
    title?:string;
    isRequired: boolean;
    id?:string
    width?:Number
    // refElement: any;
}

const InputCheckBox = ({ label, name, checked, onChange, isRequired, id, isdisable=false,title='', width}: InputTextProps) => {

    return (
        // <div className={ width == 6?"col-md-6" : "col-md-4"}>
            <div className='mt-3'>
                <input title={title} type='checkbox' checked={checked} required={isRequired} onChange={onChange} name={name} autoComplete="off" disabled={isdisable} id={id}/> <label title={title} className="col-form-label chkLbl" htmlFor={id}>{label}</label>
            </div>

    );
};

export default InputCheckBox;