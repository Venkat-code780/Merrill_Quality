import * as React from 'react';
import { useState, useRef, useEffect, useImperativeHandle, forwardRef } from 'react';
import "../CSS/Sketch.css";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faEraser, faPen, faRedo, faSlash, faTimes, faUndo } from '@fortawesome/free-solid-svg-icons';

export interface SketchProps {
    initialImage?: string;
}

export interface SketchHandle {
    getImageData: () => string | null;
}

const Sketch = forwardRef<SketchHandle, SketchProps>((props, ref) => {
    const [tool, setTool] = useState<string>('pen');
    const [color, setColor] = useState<string>('#000000'); // Default pen color is black
    const [fontSize, setFontSize] = useState<number>(5); // Default font size
    const [isDrawing, setIsDrawing] = useState<boolean>(false);
    const [lastPosition, setLastPosition] = useState<{ x: number, y: number }>({ x: 0, y: 0 });
    const [startPosition, setStartPosition] = useState<{ x: number, y: number }>({ x: 0, y: 0 });
    const [currentShape, setCurrentShape] = useState<any>(null); // Used to store in-progress shapes
    const [history, setHistory] = useState<string[]>([]);
    const [redoHistory, setRedoHistory] = useState<string[]>([]);
    const canvasRef = useRef<HTMLCanvasElement | null>(null);

    useImperativeHandle(ref, () => ({
        getImageData: () => {
            if (canvasRef.current) {
                return canvasRef.current.toDataURL(); // returns base64 image
            }
            return null;
        }
    }));

    useEffect(() => {
        if (props.initialImage && canvasRef.current) {
            const img = new Image();
            img.src = props.initialImage;
            img.onload = () => {
                const ctx = getContext();
                if (ctx) ctx.drawImage(img, 0, 0);
            };
        }
    }, [props.initialImage]);


    // Set up the canvas context
    const getContext = () => {
        return canvasRef.current?.getContext('2d');
    };

    const numArr = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

    useEffect(() => {
        const ctx = getContext();
        if (ctx) {
            ctx.lineJoin = 'round';
            ctx.lineCap = 'round';
            ctx.lineWidth = fontSize;
        }
    }, [fontSize]);

    // Function to start drawing
    const startDrawing = (e: React.MouseEvent) => {
        const ctx = getContext();
        if (ctx && canvasRef.current) {
            const rect = canvasRef.current.getBoundingClientRect();
            const pos = { x: e.clientX - rect.left, y: e.clientY - rect.top };
            setLastPosition(pos);
            setStartPosition(pos);
            setIsDrawing(true);
            if (history.length === 0) {
                saveHistory(); // Save the initial blank state
            }
            //saveHistory(); // Save the current state of the canvas before starting drawing
        }
    };

    // Function to save the canvas state to history
    const saveHistory = () => {
        if (canvasRef.current) {
            //   const ctx = getContext();
            const imageUrl = canvasRef.current.toDataURL();
            setHistory((prevHistory) => [...prevHistory, imageUrl]);
        }
    };

    // Function to draw on canvas
    const draw = (e: React.MouseEvent) => {
        if (!isDrawing) return;

        const ctx = getContext();
        if (ctx && canvasRef.current) {
            const rect = canvasRef.current.getBoundingClientRect();
            const newPos = { x: e.clientX - rect.left, y: e.clientY - rect.top };

            if (tool === 'pen') {
                ctx.strokeStyle = color;
                ctx.beginPath();
                ctx.moveTo(lastPosition.x, lastPosition.y);
                ctx.lineTo(newPos.x, newPos.y);
                ctx.stroke();
            } else if (tool === 'eraser') {
                ctx.clearRect(newPos.x - 10, newPos.y - 10, 20, 20);
            } else if (tool === 'line' || tool === 'ellipse') {
                // Clear canvas and redraw last saved state
                const img = new Image();
                img.src = history[history.length - 1]; // last saved state
                img.onload = () => {
                    ctx.clearRect(0, 0, canvasRef.current!.width, canvasRef.current!.height);
                    ctx.drawImage(img, 0, 0);

                    const shape = {
                        type: tool,
                        start: startPosition,
                        end: newPos,
                    };
                    drawShape(shape);
                };
            }

            setLastPosition(newPos);
        }
    };


    // Stop drawing
    // const stopDrawing = () => {
    //     setIsDrawing(false);
    //     if (currentShape) {
    //         drawShape(currentShape);
    //     }
    //     setCurrentShape(null);
    // };
    const stopDrawing = (e: any, isMouseUp?: boolean) => {
        setIsDrawing(false);

        if (currentShape) {
            drawShape(currentShape);
        }
        setCurrentShape(null);
        if (isMouseUp) {
            saveHistory();
        }
    };

    // Function to draw the current shape
    const drawShape = (shape: any) => {
        const ctx = getContext();
        if (!ctx || !canvasRef.current) return;

        if (shape.type === 'line') {
            ctx.strokeStyle = color;
            ctx.beginPath();
            ctx.moveTo(shape.start.x, shape.start.y);
            ctx.lineTo(shape.end.x, shape.end.y);
            ctx.stroke();
        } else if (shape.type === 'ellipse') {
            const width = shape.end.x - shape.start.x;
            const height = shape.end.y - shape.start.y;
            ctx.strokeStyle = color;
            ctx.beginPath();
            ctx.ellipse(shape.start.x, shape.start.y, Math.abs(width), Math.abs(height), 0, 0, Math.PI * 2);
            ctx.stroke();
        }
    };

    // Clear the canvas
    const clearCanvas = () => {
        const ctx = getContext();
        if (ctx && canvasRef.current) {
            ctx.clearRect(0, 0, canvasRef.current.width, canvasRef.current.height);
            saveHistory(); // Save after clearing the canvas
        }
    };

    // Get image (save canvas as image)
    const getImage = () => {
        if (canvasRef.current) {
            const imageUrl = canvasRef.current.toDataURL(); // Get canvas as base64 image
            console.log(imageUrl);
        }
    };

    // Change drawing tool
    const changeTool = (newTool: string) => {
        setTool(newTool);
    };

    const changeColor = (newColor: string) => {
        setColor(newColor);
        setTool(tool);
    };

    // Font Size Change
    const handleFontSizeChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
        setFontSize(parseInt(e.target.value));
    };

    // Undo the last action
    const undo = () => {
        if (history.length === 0) return;

        const lastState = history[history.length - 1];
        const newHistory = history.slice(0, history.length - 1);
        setRedoHistory((prevRedo) => [lastState, ...prevRedo]);

        const img = new Image();
        // img.src = lastState;
        img.src = newHistory[newHistory.length - 1] || ''; // Load previous state or blank
        img.onload = () => {
            const ctx = getContext();
            if (ctx && canvasRef.current) {
                ctx.clearRect(0, 0, canvasRef.current.width, canvasRef.current.height);
                ctx.drawImage(img, 0, 0);
            }
        };

        setHistory(newHistory);
    };

    // Redo the last undone action
    const redo = () => {
        if (redoHistory.length === 0) return;

        const lastState = redoHistory[0];
        const newRedoHistory = redoHistory.slice(1);

        setHistory((prevHistory) => [...prevHistory, lastState]);

        const img = new Image();
        img.src = lastState;
        img.onload = () => {
            const ctx = getContext();
            if (ctx && canvasRef.current) {
                ctx.clearRect(0, 0, canvasRef.current.width, canvasRef.current.height);
                ctx.drawImage(img, 0, 0);
            }
        };

        setRedoHistory(newRedoHistory);
    };

    // Determine the cursor class based on the selected tool
    const getCursorClass = () => {
        switch (tool) {
            case "pen":
                return "cursor-pen";
            case "eraser":
                return "cursor-eraser";
            case "line":
            case "ellipse":
                return "cursor-plus";
            default:
                return "";
        }
    };

    return (
        <div className={`sketch-container`}>
            <div className="toolbar">
                <button type='button' onClick={undo} disabled={history.length === 0} title='Undo'><FontAwesomeIcon icon={faUndo} /></button>
                <button type='button' onClick={redo} disabled={redoHistory.length === 0} title='Redo'><FontAwesomeIcon icon={faRedo} /></button>
                <button type='button' onClick={clearCanvas} title='Clear'><FontAwesomeIcon icon={faTimes} /></button>
                {/* <button type='button' onClick={() => changeTool('ellipse')} title='Circle'><FontAwesomeIcon icon={faCircleRegular} /></button> */}
                <button type='button' onClick={() => changeTool('ellipse')} title='Circle'><p className='faCircle'></p></button>
                <button type='button' onClick={() => changeTool('line')} title='Line'><FontAwesomeIcon icon={faSlash} /></button>
                <button type='button' onClick={() => changeTool('pen')} title='Pen'><FontAwesomeIcon icon={faPen} /></button>
                <button type='button' onClick={() => changeTool('eraser')} title='Eraser'><FontAwesomeIcon icon={faEraser} /></button>
                <select onChange={handleFontSizeChange} value={fontSize}>
                    {/* <option value={5}>5px</option> */}
                    {numArr.map((num) => (<option value={num}>{num + "px"}</option>))}
                </select>
                <input type="color" onChange={(e) => changeColor(e.target.value)} />
                <button type='button' onClick={getImage} className='d-none'>Get Image</button>
            </div>

            <canvas
                ref={canvasRef}
                width={400}
                height={340}
                style={{ border: '1px solid #000' }}
                onMouseDown={startDrawing}
                onMouseMove={draw}
                onMouseUp={(e) => stopDrawing(e, true)}
                onMouseLeave={stopDrawing}
                className={`sketchCanvas ${getCursorClass()}`}
            />
        </div>
    );
});

export default Sketch;
