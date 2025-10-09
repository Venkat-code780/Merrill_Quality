import { faPaperclip, faCloudUploadAlt, faClose} from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { useRef } from "react";
import * as React from "react";


function FileUpload(props:any) {
    const isMultiAllowed = props.ismultiAllowed;
    const isFileCloseShow = props.isFileCloseShow;
    const isRequired = props.isRequired;
    const isDisabled = props.disabled;
    // const showFileInput = props.showFileInput;

    var fileArr = props.files[0];
    var delefileArr = props.files[1];
    let inputFileRef = useRef(null);


    function showFilePopup() {
        if( inputFileRef.current ){
            let inputCurr:any = inputFileRef.current;
            inputCurr.click();
        }
    }

    function handleFileUpload(e:any){
        e.preventDefault();

        let arrFiles = Array.from(e.target.files);
        let stateArrFiles = fileArr;
        arrFiles.map( ( selItem:any, index ) => {
            let filename:string = selItem["name"];

            let checkexisting = stateArrFiles.filter((file:any) => {
                return filename == file.name
            });
            selItem['IsNew'] = true;
            selItem['IsDeleted'] = false;
            if( checkexisting.length == 0 ){
                stateArrFiles.push(selItem);
            }
        });
        props.onFileChanges([stateArrFiles, delefileArr]);
        e.target.value = '';
    }

    function removeSelectedFile(fileName:any) {
        var element = document.getElementsByClassName("focus-Div");
        if( element.length > 0 ){
            for( let i=0;i<element.length;i++){
                element[i].classList.remove("focus-Div");
            }
        }
        let fileColl = fileArr;
        let fileCollAfterRemove = fileColl.filter((file:any) => {
            return file.name != fileName
        });
        let filearryRemove = fileColl.filter((file:any) => {
            return file.name == fileName && file.IsNew == false
        });
        if( filearryRemove.length > 0){
            delefileArr.push(filearryRemove[0]);
        }
        props.onFileChanges([fileCollAfterRemove, delefileArr]);
    }

    function renderFiles() {
        var files:any = fileArr;
        let fileInd = -1;
        const fsArr = files.map( (file:any) => {
            let fileName = file.name;
            let fileUrl = file.URL;
            fileInd++;

            if( fileUrl != undefined && fileUrl != null ){
                return ( <li className="hoverclass" id={"li_"+fileInd} title={fileName}><a target="_blank" download={fileName} href={fileUrl}><FontAwesomeIcon icon={faPaperclip}></FontAwesomeIcon> <span> {fileName} </span></a><span className="close" hidden={isFileCloseShow}><FontAwesomeIcon onClick={() => removeSelectedFile(fileName)} icon={faClose} /></span></li>)
            }
            else{
                
                return (<li className="hoverclass" id={"li_"+fileInd} title={fileName}><FontAwesomeIcon icon={faPaperclip} className="liNewFile"></FontAwesomeIcon> <span className="attachName" title={fileName}> {fileName} </span> <span className="close"> <FontAwesomeIcon onClick={() => removeSelectedFile(fileName)} icon={faClose} /></span></li>);
            }
        });
        return fsArr;
    }

    

    return(

        <div className="">
            <div className="p-3 light-box m-1">
                <div className="row">
                    <div className="col-4"><h6 className="my-2">Attachment { isRequired && <span className="text-danger">*</span>}</h6></div>
                    <div className="col-8">
                        <div className="">
                        <button type="button"  onClick={showFilePopup} className="btn btn-outline-dark" title="Choose File" disabled={isDisabled}>Choose File <FontAwesomeIcon icon={faCloudUploadAlt}></FontAwesomeIcon> </button>
                        <input multiple={isMultiAllowed} ref={inputFileRef} type="file" onChange={handleFileUpload} title="Please choose file" style={{ "display": "none" }} disabled={isDisabled} />
                        </div>
                    </div>
                </div>
                <div className="col-md-12 col-sm-12 col-xs-12 col-12">
                    <ul className="attachment-list px-2 row">
                        {renderFiles()}
                    </ul>

                </div>
            </div>
        </div>
    );
}


export default FileUpload;