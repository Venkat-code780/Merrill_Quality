import * as React from "react";
import Select from "react-select";

interface DropdownProps {
    label: string;
    Title: string;
    name: string;
    id: any;
    placeholderText?: string;
    className: string;
    selectedValue: any;
    selectedLabel?: any;
    OptionsList: any;
    OnChange: any;
    isRequired: boolean;
    disabled: boolean;
    refElement?: any;
    noOptionsMessage?: string;
    isMultiple?:boolean
}

const SearchableDropdown = ({ label, Title, name, id, placeholderText, className, selectedValue, OptionsList, OnChange, isRequired, disabled = false, refElement=null, noOptionsMessage = "No options", isMultiple= false }: DropdownProps) =>{
    const onBlur = () => {
        document.getElementById(id)?.classList.remove('searchMandatory');
    }
    return(
        <React.Fragment>
            {![null,""].includes(label)?<label className={disabled? "label-inactive": ''}>{label}
                { isRequired && <span className="mandatoryhastrick"> *</span>}
            </label>:''}
            <Select
                name={name}
                id={id}
                divId={'divSearch'}
                titleText={Title}
                placeholder={placeholderText}
                className={className}
                value={ isMultiple? selectedValue:( OptionsList.find( (option:any) => option.value === selectedValue )|| '')}
                options={OptionsList}
                onChange={(selectedOption: any, actionMeta: any) => { OnChange(selectedOption, actionMeta)}}
                onBlur={onBlur()}
                isDisabled = {disabled}
                ref = {refElement}
                isClearable = { !(['', "None",null,undefined].includes(selectedValue))}
                // isClearable = { true }
                isSearchable = {true}
                noOptionsMessage = {()=> noOptionsMessage}
                isMulti={isMultiple}
                trim={isMultiple}
            ></Select>
        </React.Fragment>
    )
}

export default SearchableDropdown;