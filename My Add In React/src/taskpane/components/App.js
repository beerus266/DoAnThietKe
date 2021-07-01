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
                 const range = context.workbook.getSelectedRange();
                 // Read the range address
                 range.load("address");
                var sheet = context.workbook.worksheets.getActiveWorksheet();

                 await context.sync();
                 console.log(`The range address was ${range.address}.`);
                 var addr = range.address.slice(range.address.indexOf("!") + 1, range.address.length);
                 var val = this.calculateNextPosition('left', addr)
                //  range.values = [[ val ]];
                var range1 = sheet.getRange(val);
                range1.select();
      });
    } catch (error) {
      console.error(error);
    }
  }
  //============ Move Selected Cell Up When Hand move up ============ 
  upCell = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
       var sheet = context.workbook.worksheets.getActiveWorksheet();

        await context.sync();
        console.log(`The range address was ${range.address}.`);
        var addr = range.address.slice(range.address.indexOf("!") + 1, range.address.length);
        var val = this.calculateNextPosition('up', addr)
        //  range.values = [[ val ]];
        var range1 = sheet.getRange(val);
        range1.select();
      });
    } catch (error) {
      console.error(error);
    }
  }
 //============ Move Selected Cell Down When Hand move down ============ 
  downCell = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
       var sheet = context.workbook.worksheets.getActiveWorksheet();

        await context.sync();
        console.log(`The range address was ${range.address}.`);
        var addr = range.address.slice(range.address.indexOf("!") + 1, range.address.length);
        var val = this.calculateNextPosition('down', addr)
       //  range.values = [[ val ]];
       var range1 = sheet.getRange(val);
       range1.select();

      });
    } catch (error) {
      console.error(error);
    }
  }
 //============ Move Selected Cell Left When Hand move Left ============ 
  leftCell = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
       var sheet = context.workbook.worksheets.getActiveWorksheet();

        await context.sync();
        console.log(`The range address was ${range.address}.`);
        var addr = range.address.slice(range.address.indexOf("!") + 1, range.address.length);
        var val = this.calculateNextPosition('left', addr)
       //  range.values = [[ val ]];
       var range1 = sheet.getRange(val);
       range1.select();

      });
    } catch (error) {
      console.error(error);
    }
  }
 //============ Move Selected Cell Right When Hand move Right ============ 
  rightCell = async () => {
    try {
      await Excel.run(async (context) => {

        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
       var sheet = context.workbook.worksheets.getActiveWorksheet();

        await context.sync();
        console.log(`The range address was ${range.address}.`);
        var addr = range.address.slice(range.address.indexOf("!") + 1, range.address.length);
        var val = this.calculateNextPosition('right', addr)
       //  range.values = [[ val ]];
       var range1 = sheet.getRange(val);
       range1.select();

      });
    } catch (error) {
      console.error(error);
    }
  }

  calculateNextPosition = (direction, addr) => {
    var myArray = addr.split(/([0-9]+)/);
    // return positionChar;
    if (direction == 'up'){
      return String.fromCharCode(myArray[0].charCodeAt(0) ) + (parseInt(myArray[1]) - 1);
    } else if (direction == 'down'){
      return String.fromCharCode(myArray[0].charCodeAt(0) ) + (parseInt(myArray[1]) + 1);
    }else if (direction == 'left'){
      return String.fromCharCode(myArray[0].charCodeAt(0) - 1) + (parseInt(myArray[1]));
    }else if (direction == 'right'){
      return String.fromCharCode(myArray[0].charCodeAt(0) + 1 ) + (parseInt(myArray[1]));
    }
  }

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
