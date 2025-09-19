import * as React from 'react';
import  '../CSS/ImageUploader.css';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faClose } from '@fortawesome/free-solid-svg-icons';

interface IMultipleImageUploaderProps {
  onImageUpload: (imageData: string, lineIndex:number) => void;
  onRemoveImage: (lineIndex:number) => void;
  initialImageSrc?: string | null;
  index: number;
}

const MultipleImageUploader: React.FC<IMultipleImageUploaderProps> = ({ onImageUpload, onRemoveImage, initialImageSrc, index }) => {
  const [previewSrc, setPreviewSrc] = React.useState<string | null>(initialImageSrc || null);
  const [showModal, setShowModal] = React.useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const canvasRef = React.useRef<HTMLCanvasElement>(null);

  React.useEffect(() => {
    if (initialImageSrc !== previewSrc) {
      setPreviewSrc(initialImageSrc || null);
    }
  }, [initialImageSrc]);

  const handleImageChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.src = e.target?.result as string;

      img.onload = () => {
        const canvas = canvasRef.current;
        const ctx = canvas?.getContext('2d');
        if (!canvas || !ctx) return;

        let { width, height } = img;
        let MAX_WIDTH = 800;

        const srcLength = img.src.length;
        if (srcLength < 500000) MAX_WIDTH = 800;
        else if (srcLength < 1000000) MAX_WIDTH = 600;
        else if (srcLength < 1500000) MAX_WIDTH = 400;
        else MAX_WIDTH = 300;

        const scaleFactor = MAX_WIDTH / width;
        width = MAX_WIDTH;
        height = height * scaleFactor;

        canvas.width = width;
        canvas.height = height;

        ctx.drawImage(img, 0, 0, width, height);

        const dataUrl = canvas.toDataURL("image/jpeg", 0.5).replace(/ /g, "+");
        setPreviewSrc(dataUrl);
        onImageUpload(dataUrl, index); // callback to store base64
      };
    };
    reader.readAsDataURL(file);
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  const handleImageClick = () => {
    setShowModal(true);
  };

  const closeModal = () => {
    setShowModal(false);
  };

  const handleRemoveClick = () => {
    setPreviewSrc(null);
    onRemoveImage(index);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  }

  return (
    <div className="MultipleimageUploader">
      {/* <label className="label">Upload Image</label> */}
      <div className="uploadContainer">
        <div onClick={triggerFileSelect} style={{cursor:"pointer"}}>
            <img
            src="/sites/ventureglobal/mayco/merrill/sa/SiteAssets/Images/upload.png"
            title="Upload Image"
            className="uploadIcon"
            />
        </div>
        {previewSrc && (
          <div className='previewContainer '>
            <img
              src={previewSrc}
              className="previewImage"
              onClick={handleImageClick}
              alt="Uploaded"
              />
            <button className="removeButton" onClick={handleRemoveClick}>
              <FontAwesomeIcon icon={faClose} />
            </button>
          </div>
        )}
        <input
          type="file"
          accept="image/*"
          ref={fileInputRef}
          onChange={handleImageChange}
          style={{ display: 'none' }}
        />
        <canvas ref={canvasRef} style={{ display: 'none' }} />
      </div>
        

      {showModal && previewSrc && (
        <div className="modalOverlay" onClick={closeModal}>
          <div className="modalContent">
            <img src={previewSrc} className="fullImage" />
          </div>
        </div>
      )}
    </div>
  );
};

export default MultipleImageUploader;
