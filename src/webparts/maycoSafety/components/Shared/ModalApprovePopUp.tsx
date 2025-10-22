import * as React from 'react';
interface modalProps {
  message: string;
  title: string;
  isVisible: boolean;
  isSuccess: boolean;
  isManager?:boolean;
  errorMessage: string;
  onConfirm: (e: any) => void;
  onCancel: () => void;
  comments: (e: any) => void;
  commentsValue: string;
  modalHeader: string;
}

const ModalApprovePopUp = ({ message, modalHeader, title, isVisible, isSuccess,isManager=true, onConfirm, onCancel, comments, errorMessage, commentsValue}: modalProps) => {
  return isVisible ? (
    <div className="modal" tabIndex={-1} style={{ display: 'block' }} >
      <div className="modal-dialog modal-dialog-centered">
        <div className="modal-content">
          <div className={modalHeader}>

            {!isSuccess && <h5 className="" color='rgb(232, 87, 87)'>{message}</h5>}
            {isSuccess && <h5 className="" color='#0D2F4B'>{message}</h5>}
          </div>
          { title == "Reject" && <div className="light-box border-box-shadow m-1 p-2">
            <div className="media-px-12">
              <div className="light-text height-auto">
                <label className="floatingTextarea2 top-12">Comments{!isSuccess && <span className='mandatoryhastrick'>*</span>} </label>
                <textarea className="position-static form-control requiredinput" onChange={comments} value={commentsValue} placeholder=""  id="txtComments" name="comments" disabled={false} title='Comments'></textarea>
              </div>
              <div>
                <span className='text-validator'> {errorMessage}</span>
              </div>

            </div>
          </div>}
          <div className="modal-footer">
            {<button type="button" onClick={onConfirm} id="btnOK" className="btn  btn-danger" data-dismiss="modal" title='Ok'>Ok</button>}
            <button type="button" onClick={onCancel} id="btnCancel" className="btn btn-secondary" data-dismiss="modal" title='Cancel'>Cancel</button>
          </div>
        </div>
      </div>
    </div>
  ) : null;
};

export default ModalApprovePopUp;