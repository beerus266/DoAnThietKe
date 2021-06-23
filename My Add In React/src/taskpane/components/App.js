import * as React from "react";
import PropTypes from "prop-types";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import ButtonCamera from "./hooks/ButtonCamera";
//  global console, Excel 

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  //============= Test Function ================
  click = async () => {   
    try {
      await Excel.run(async (context) => {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        // const range = context.workbook.getSelectedRange();
        var range = sheet.getRange("B1");
        // var addr = range.load("address");

        range.select();

        var sheet1 = context.workbook.worksheets.getItem("Sample");

        var range1 = sheet1.getRange("C3");
        range1.values = [[ 5 ]];
        range1.format.autofitColumns();

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
  //============ Move Selected Cell Up When Hand move up ============ 
  upCell = async () => {
    try {
      await Excel.run(async (context) => {

        // Get Address of Cell Selected
        var rangeSel = context.workbook.getSelectedRange();
        rangeSel.load("address");
        var pos = rangeSel.address;

        // Calculate next Cell
        pos = myArray.split(/([0-9]+)/);
        positionChar = myArray[0];
        positionInt = myArray[1];
        let nextPosition = String.fromCharCode(myArray[0].charCodeAt(0) ) + (parseInt(myArray[1]) - 1);
        
        // Move to next cell
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(nextPosition);
        range.select();

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
 //============ Move Selected Cell Down When Hand move down ============ 
  downCell = async () => {
    try {
      await Excel.run(async (context) => {

        var rangeSel = context.workbook.getSelectedRange();
        rangeSel.load("address");
        var pos = rangeSel.address;

        pos = myArray.split(/([0-9]+)/);
        positionChar = myArray[0];
        positionInt = myArray[1];
        let nextPosition = String.fromCharCode(myArray[0].charCodeAt(0) ) + (parseInt(myArray[1]) + 1);
        
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(nextPosition);

        range.select();

        await context.sync();
        console.log(`The range address was ${range.address}.`);

      });
    } catch (error) {
      console.error(error);
    }
  }
 //============ Move Selected Cell Left When Hand move Left ============ 
  leftCell = async () => {
    try {
      await Excel.run(async (context) => {

        var rangeSel = context.workbook.getSelectedRange();
        rangeSel.load("address");
        var pos = rangeSel.address;

        pos = myArray.split(/([0-9]+)/);
        positionChar = myArray[0];
        positionInt = myArray[1];
        let nextPosition = String.fromCharCode(myArray[0].charCodeAt(0) - 1) + (parseInt(myArray[1]));
        
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(nextPosition);

        range.select();

        await context.sync();
        console.log(`The range address was ${range.address}.`);

      });
    } catch (error) {
      console.error(error);
    }
  }
 //============ Move Selected Cell Right When Hand move Right ============ 
  rightCell = async () => {
    try {
      await Excel.run(async (context) => {

        var rangeSel = context.workbook.getSelectedRange();
        rangeSel.load("address");
        var pos = rangeSel.address;

        pos = myArray.split(/([0-9]+)/);
        positionChar = myArray[0];
        positionInt = myArray[1];
        let nextPosition = String.fromCharCode(myArray[0].charCodeAt(0) + 1 ) + (parseInt(myArray[1]));
        
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(nextPosition);

        range.select();

        await context.sync();
        console.log(`The range address was ${range.address}.`);

      });
    } catch (error) {
      console.error(error);
    }
  }

  nextChar = (c) => {
    return String.fromCharCode(c.charCodeAt(0) + 1);
  }
  nextInt = (i) => {
      return parseInt(i) + 1;
  }

  nextPosition

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <div style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
          }}
        >
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronUp" }}
            onClick={this.upCell}

          >
            Up
          </Button>
        </div>
        <div style={{display:"flex"}}
        >
          <Button
            style={{float: 'left'}}
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronLeft" }}
            onClick={this.leftCell}
          >
            Left
          </Button>
          <Button
            style={{float: 'right'}}
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.rightCell}
          >
            Right
          </Button>
        </div>
          <div style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
          }}
          >
            <Button
              className="ms-welcome__action"
              buttonType={ButtonType.hero}
              iconProps={{ iconName: "ChevronDown" }}
              onClick={this.downCell}
            >
              Down
            </Button>
          </div>
          {/* <ButtonCamera/> */}
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
