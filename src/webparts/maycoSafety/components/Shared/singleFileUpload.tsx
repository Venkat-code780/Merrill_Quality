import { faCloudUploadAlt, faClose} from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { useRef } from "react";
import * as React from "react";


function SingleImageUpload(props:any) {
    const isRequired = props.isRequired;
    const isDisabled = props.disabled;

    const image = props.files[0];
    const onRemove = props.onRemove;
    const onChange = props.onChange;
    let inputFileRef = useRef(null);


    function showImageDialog() {
        if( inputFileRef.current ){
            let inputCurr:any = inputFileRef.current;
            inputCurr.click();
        }
    }

    function handleImageChange(e:any){
        e.preventDefault();
        const file = e.target.files[0];

        if (!file) return;

        // Only accept image files
        if (!file.type.startsWith("image/")) {
            alert("Only image files are allowed.");
            return;
        }

        file['IsNew'] = true;
        file['IsDeleted'] = false;

        onChange([file]);
        e.target.value = ''; // reset input
    }

    function removeImage() {
        onRemove([]);
    }

        return (
        <div className="p-3 light-box m-1">
            <div className="row">
                <div className="col-4">
                    <h6 className="my-2">
                        {props.fileLabel} {isRequired && <span className="text-danger">*</span>}
                    </h6>
                </div>
                <div className="col-8">
                    <button
                        type="button"
                        onClick={showImageDialog}
                        className="btn btn-outline-dark"
                        title="Choose Image"
                        disabled={isDisabled}
                    >
                        Choose Image <FontAwesomeIcon icon={faCloudUploadAlt} />
                    </button>
                    <input
                        type="file"
                        accept="image/*"
                        ref={inputFileRef}
                        onChange={handleImageChange}
                        style={{ display: "none" }}
                        disabled={isDisabled}
                    />
                </div>
            </div>

            {/* Image Preview */}
            <div className="col-12 mt-3">
                {image ? (
                    <div className="d-flex align-items-center">
                        <img
                            src={image.URL || URL.createObjectURL(image)}
                            alt={image.name}
                            style={{ maxWidth: '150px', maxHeight: '150px', objectFit: 'cover' }}
                            className="me-2"
                        />
                        <button className="btn btn-sm btn-outline-danger" onClick={removeImage}>
                            <FontAwesomeIcon icon={faClose} /> Remove
                        </button>
                    </div>
                ) : (
                    <p className="text-muted">No image selected.</p>
                )}
            </div>
        </div>
    );
}


export default SingleImageUpload;