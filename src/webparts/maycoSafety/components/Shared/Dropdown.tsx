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
    optionLabel: any;
    optionValue: any;
    selectedLabel?: any;
    OptionsList: any;
    OnChange: any;
    isRequired: boolean;
    disabled: boolean;
    refElement: any;
    noOptionsMessage?: string;
    isMultiple?:boolean
}

const SearchableDropdown = ({ label, Title, name, id, placeholderText, className, selectedValue, optionLabel, optionValue, selectedLabel, OptionsList, OnChange, isRequired, disabled = false, refElement, noOptionsMessage = "No options", isMultiple= false }: DropdownProps) =>{

    const options = OptionsList.map((item: any) => ({
        label: typeof(item) == "string" ? item: optionLabel.includes('.') ? item[optionLabel.split('.')[0]][optionLabel.split('.')[1]]:item[optionLabel],
        value: typeof(item) == "string" ? item: optionValue.includes('.') ? item[optionValue.split('.')[0]][optionValue.split('.')[1]]:item[optionValue],
        EMail: typeof(item) == "string" ? item: item["EMail"]
    }));
    const onBlur = () => {
        document.getElementById(id)?.classList.remove('searchMandatory');
    }
    return(
        <React.Fragment>
            <label>{label}
                { isRequired && <span className="mandatoryhastrick"> *</span>}
            </label>
            <Select
                name={name}
                id={id}
                divId={'divSearch'}
                title={Title}
                placeholder={placeholderText}
                className={className}
                value={ isMultiple? selectedValue:( options.find( (option:any) => option.value === selectedValue )|| '')}
                options={options}
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