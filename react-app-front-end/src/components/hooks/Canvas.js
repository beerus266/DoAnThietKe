import React, {createRef, useEffect, useState } from 'react';

// Xử lý Bounding box dự đoán được và vẽ ra các Frames
export default function Canvas({ videoRef, model}){

    const [canvasRef] = useState(createRef(null));
    let goToRight = null;
    let x = 0;
    useEffect(() => {
        if (canvasRef.current && videoRef.current ) {
            const interval = setInterval(() => {
                const ctx = canvasRef.current?.getContext('2d');
                // Dự đoán các Bounding Box có thể có
                model.detect(videoRef.current).then(predictions => {
                    if (canvasRef.current){
                        model.renderPredictions(predictions, canvasRef.current, ctx, videoRef.current);
                        predictions.map(prediction => {      // bbox = [x, y, width, height]
                            if (predictions){
                                if ( x < predictions[0].bbox[0]) { // xác định thao tác bàn tay di chuyển sang Phải
                                    goToRight = true;
                                    console.log("RIGHT");
                                } else { // xác định thao tác bàn tay di chuyển sang Trái
                                    goToRight = false;
                                    console.log("LEFT");
                                }
                                x = predictions[0].bbox[0];
                            }
                        });
                    }
                });
            }, 0);
  
                return () => clearInterval(interval)
            } else {
                console.log("no canvas");
            }
    },[]);

    return (
            <canvas ref={canvasRef} width="420" height="300" />
    );
}