import Body, { ExtendedBodyPart } from "@mjcdev/react-body-highlighter";
import * as React from "react";
import "../CSS/BodyChart.css";

interface BodyPartProps {
  selectedBodyPart: ExtendedBodyPart[];
}

const BodyPart: React.FC<BodyPartProps> = ({ selectedBodyPart }) => {
  return (
    <React.Fragment>
      <div className="BodyChartDiv">
        {/* Front side body parts */}
        <Body
          data={selectedBodyPart}
          gender="male"
          side="front"
          scale={1.7}
          border="#dfdfdf"
        />
        {/* Back side body parts */}
        <Body
          data={selectedBodyPart}
          gender="male"
          side="back"
          scale={1.7}
          border="#dfdfdf"
        />
      </div>
    </React.Fragment>
  );
};

export default BodyPart;
