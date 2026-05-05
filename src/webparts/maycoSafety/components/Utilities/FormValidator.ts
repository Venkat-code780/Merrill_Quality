import { ControlType } from "../Constants/Contants";

function ValidateForm(data:any){
    var element = document.getElementsByClassName("mandatory-FormContent-focus");
    if( element.length > 0 ){
        for( let i=0;i<element.length;i++){
            element[i].classList.remove("mandatory-FormContent-focus");
        }
    }
    var radioelement = document.getElementsByClassName("focus-Div");
    if( radioelement.length > 0 ){
        for( let i=0;i<radioelement.length;i++){
            radioelement[i].classList.remove("focus-Div");
        }
    }
    let status = true;
    let message ="";
    let propertieTypes={Number:ControlType.number,String:ControlType.string,MobileNumber:ControlType.mobileNumber,Email:ControlType.email,People:ControlType.people,Date:ControlType.date,DateTime:ControlType.DateTime,compareDates:ControlType.compareDates, reactSelect: ControlType.reactSelect, radio: ControlType.radio, lessthanTodayDate: ControlType.lessthanTodayDate, greaterthanTodayDate: ControlType.greaterthanTodayDate, compareNumbers: ControlType.compareNum, ArrayField: ControlType.array, limitedNumber: ControlType.limitedNumber};
    for (let key in data) {
        let value = data[key].val;
        let type =data[key].Type;
        let isrequired =data[key].required;
        if([undefined,null,''].includes(value) && ( propertieTypes.Number == type || propertieTypes.String == type || propertieTypes.MobileNumber == type || propertieTypes.Email == type ) && isrequired )
            // propertieTypes.People!=type && propertieTypes.Date!=type && propertieTypes.reactSelect != type && propertieTypes.radio != type && isrequired )
        {
            let prpel =data[key].Focusid;
            prpel.current.focus();
            prpel.current.classList.add('mandatory-FormContent-focus');
            message ="'"+data[key].Name+"' cannot be blank.";
            status = false;
            break;
        }
        else if( [undefined,null,'',0].includes(value) && propertieTypes.reactSelect == type && isrequired ){
            message ="'"+data[key].Name+"' cannot be blank.";
            let prpel =data[key].Focusid;
            document.getElementById(prpel)?.getElementsByTagName("input")[0].focus();
            setTimeout(() =>{document.getElementById(prpel)?.classList.add('searchMandatory')},100);
            status = false;
            break;
        }
        else if( [undefined,null,''].includes(value) && propertieTypes.radio == type && isrequired ){
            message ="'"+data[key].Name+"' cannot be blank.";
            let prpel =data[key].Focusid;
            let element = document.getElementById(prpel);
            if( element ){
                element.focus();
                element.classList.add('focus-Div');
            }
            status = false;
            break;
        }
        else if( !([undefined,null,''].includes(value)) && propertieTypes.Email == type && isrequired ){
            let prpel =data[key].Focusid;
            let isEmail = IsEmail(value);

            if( !isEmail ){
                prpel.current.focus();
                prpel.current.classList.add('mandatory-FormContent-focus');
                message ="Please enter valid Email";
                status = false;
                break;
            }
        }
        else if(((propertieTypes.People==type || propertieTypes.ArrayField==type) && isrequired) && (!value || value.length === 0))
        {
            message = "'"+data[key].Name+"' cannot be blank.";
            let prpIsreq =data[key].Focusid;
            let element = document.getElementById(prpIsreq);

            if( element ){
                element?.getElementsByTagName('input')[0]?.focus();
                element.classList.add('focus-Div');
            }
            status = false;
            break;
        }
        else if((propertieTypes.Date==type && isrequired) && [undefined,null,''].includes(value))
        {
            message ="'"+data[key].Name+"' cannot be blank.";
            let prpData =data[key].Focusid;
            let element = document.getElementById(prpData);
            element?.focus();
            setTimeout(() =>{element?.classList.add('mandatory-FormContent-focus');},100);
            status = false;
            break;
        }
        else if((propertieTypes.DateTime==type && isrequired) && [undefined,null,''].includes(value))
        {
            message ="'"+data[key].Name+"' cannot be blank.";
            let prpData =data[key].Focusid;
            let element = document.getElementById(prpData);
            element?.getElementsByTagName('input')[0].focus();
            setTimeout(() =>{
                element?.children[0].children[0].classList.remove('Mui-focused'); //default class removed
                element?.children[0].children[0].classList.add('focus-Div');
            },300);
            status = false;
            break;
        }
        else if(propertieTypes.MobileNumber ==type && ![undefined,null,''].includes(value)&& (!isNaN(value) || Math.floor(value) !=value))
        {
            let prpMob =data[key].Focusid;
            message ="'"+data[key].Name+"' enter valid number.";
            prpMob.current.focus();
            prpMob.current.classList.add('mandatory-FormContent-focus');
            status = false;
            break;
        }
        else if(propertieTypes.Number ==type && ![undefined,null,''].includes(value)&& (isNaN(value) || value <=0) )
        {
            let prpNum =data[key].Focusid;
            message ="'"+data[key].Name+"' enter valid number.";
            prpNum.current.focus();
            prpNum.current.classList.add('mandatory-FormContent-focus');
            status = false;
            break;
        }
        else if(propertieTypes.limitedNumber ==type && ![undefined,null,''].includes(value)&& (isNaN(value) || value <=0 || value.length !=  Number(data[key].limit )) )
        {
            let prpNum =data[key].Focusid;
            message ="'"+data[key].Name+"' must be exactly "+ data[key].limit +" digits.";
            prpNum.current.focus();
            prpNum.current.classList.add('mandatory-FormContent-focus');
            status = false;
            break;
        }
        else if(propertieTypes.compareDates == type)
        {
            let startDate = data[key].startDate;
            let EndDate = data[key].endDate;
            if( !([null, undefined, ''].includes(startDate)) && !([null, undefined, '']).includes(EndDate)){
                startDate=new Date(startDate).setHours(0,0,0,0);
                EndDate=new Date(EndDate).setHours(0,0,0,0);
                if(new Date(EndDate).getTime() < new Date(startDate).getTime()){
                    message ="'"+data[key].endDateName+"' must be greater than or equal to '"+data[key].startDateName+"'.";
                    let prpData =data[key].Focusid;
                    
                    let element = document.getElementById(prpData);
                    
                    if( element ){
                        element.focus();
                        setTimeout(() =>{element?.classList.add('mandatory-FormContent-focus');},100);
                    }
                    status = false;
                    break;
                }
            }
        }
        else if( propertieTypes.lessthanTodayDate == type && ![null, undefined, ''].includes(value) ){
            let dateVal = new Date(value).setHours(0,0,0,0);
            let todayDate = new Date().setHours(0,0,0,0);
            if( dateVal > todayDate ){
                message ="'"+data[key].Name+"' should be less than or equals to Today.";
                let prpData =data[key].Focusid;

                let element = document.getElementById(prpData);

                if( element ){
                    element.focus();
                    element.classList.add('mandatory-FormContent-focus');
                }
                status = false;
                break;
            }
        }
        else if( propertieTypes.greaterthanTodayDate == type && ![null, undefined, ''].includes(value) ){
            let dateVal = new Date(value).setHours(0,0,0,0);
            let todayDate = new Date().setHours(0,0,0,0);
            if( dateVal < todayDate ){
                message ="'"+data[key].Name+"' should be greater than or equals to Today.";
                let prpData =data[key].Focusid;

                let element = document.getElementById(prpData);

                if( element ){
                    element.focus();
                    element.classList.add('mandatory-FormContent-focus');
                }
                status = false;
                break;
            }
        }
        else if( propertieTypes.compareNumbers == type ){
            let higherValue = data[key].highVal;
            let lowerValue = data[key].lowVal;

            if( lowerValue > higherValue ){
                 message ="'"+data[key].lowName+"' must be lower than or equals to '"+data[key].highName+"'.";
                let prpel =data[key].Focusid;
                prpel.current.focus();
                prpel.current.classList.add('mandatory-FormContent-focus');
                status = false;
                break;
            }
        }
    }
    let retunobject ={message,status};
    return retunobject;
}
function ValidateInputFiles( fileArray: any, isRequired: boolean ){
    let status = true;
    let message ="";
    var regex = /^[A-Za-z0-9_\- ()]+$/;

    if( fileArray.length  == 0 && isRequired ){
        message = "Please upload any document";
        status =  false;
        let retunobject ={message,status};
        return retunobject;
    }

    var element = document.getElementsByClassName("focus-Div");
    if( element.length > 0 ){
        for( let i=0;i<element.length;i++){
            element[i].classList.remove("focus-Div");
        }
    }

    for (let i in  fileArray) {
        const fileName = fileArray[i].name;
        let fileSection = fileArray[i].category  == "Capability Brochure"?"cb":fileArray[i].category =="Presentations"?"pp":"cc";;
        if(!(regex.test(fileName.replace(/\.[^/.]+$/, "")))){
            message = "Special characters are not allowed in uploaded File '" + fileName + "'";
            status = false;

            let element = document.getElementById("li_"+fileSection+i);
            if( element ){
                element.focus();
                element.classList.add('focus-Div');
            }
            break;
        }
    }
    let retunobject ={message,status};
    return retunobject;
}

function IsEmail(email:any) {
    var regex = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9_\.\-\+])+\.)+([a-zA-Z0-9]{2,4})+$/;
    if (!regex.test(email)) {
        return false;
    } else {
        return true;
    }
}


class formValidation {
    public static FormValidation=(formData: any)=>{
        let status= ValidateForm(formData); 
        return status;
    }

    public static FilesValidation = ( filesArray: any, isRequired:boolean ) =>{
        let status = ValidateInputFiles(filesArray, isRequired);
        return status;
    }
 }
 export default formValidation;

