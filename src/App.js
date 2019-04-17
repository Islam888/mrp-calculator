import React, { Component } from "react";
import XLSX from "xlsx";
import contribution from "./data/contribution.json";
import "./App.css";
import TextField from "@material-ui/core/TextField";
import MenuItem from "@material-ui/core/MenuItem";
import Button from "@material-ui/core/Button";
import { withStyles } from "@material-ui/core/styles";

const styles = theme => ({
  container: {
    display: "flex",
    flexWrap: "wrap"
  },
  number: {
    marginLeft: theme.spacing.unit,
    marginRight: theme.spacing.unit,
    width: 150
  },
  type: {
    marginLeft: theme.spacing.unit,
    marginRight: theme.spacing.unit,
    width: 250
  },
  dense: {
    marginTop: 16
  },
  menu: {
    width: 200
  },
  button: {
    margin: theme.spacing.unit,
    fontSize: "20px"
  }
});

class App extends Component {
  state = {
    type: ""
  };
  calcMrpSum = workbook => {
    var mrpSum = 0;
    for (var i = 2; i < contribution.length; i++) {
      mrpSum += workbook.Sheets[workbook.SheetNames[0]][`M${i}`].v;
    }
    this.setState({
      mrpSum
    });
  };

  calcAverage = number => {
    this.setState({
      average: Math.round((this.state.mrpSum + Number(number)) / 2)
    });
  };

  handleChange = e => {
    const type = e.target.value;
    this.setState({
      type
    });
  };

  handleFileSelect = e => {
    const label = document.querySelector('.typeBorder')
    label.querySelector('span').innerText = e.target.value.split( '\\' ).pop();
  }

  handleKeyDown = e => {
    if (e.keyCode === 13 || e.keyCode === 32) {
      const file = document.querySelector('.type')
      file.click()
    }
    
  }

  handleClick = e => {
    e.preventDefault();
    var number = document.getElementById("number").value;
    var type = document.getElementById("type").value;
    var file = document.getElementById("file").files[0];
    if (number && type && file) {
      var reader = new FileReader();
      reader.onload = e => {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        this.calcMrpSum(workbook);
        this.calcAverage(number);
        if (this.state.mrpSum > this.state.average) {
          for (var n = 2; n < contribution.length; n++) {
            contribution.forEach(elem => {
              if (
                elem.site === workbook.Sheets[workbook.SheetNames[0]][`A${n}`].v
              ) {
                workbook.Sheets[workbook.SheetNames[0]][`M${n}`] = {
                  t: "n",
                  v: Math.round(
                    ((workbook.Sheets[workbook.SheetNames[0]][`M${n}`].v +
                      elem[`${type}`] * number) /
                      2 -
                      (((workbook.Sheets[workbook.SheetNames[0]][`M${n}`].v +
                        elem[`${type}`] * number) /
                        2) *
                        ((workbook.Sheets[workbook.SheetNames[0]][`M${n}`].v -
                          elem[`${type}`] * number) /
                          2)) /
                        this.state.average) *
                      (this.state.average / this.state.mrpSum)
                  )
                };
              }
            });
          }
        } else {
          for (var x = 2; x < contribution.length; x++) {
            contribution.forEach(elem => {
              if (
                elem.site === workbook.Sheets[workbook.SheetNames[0]][`A${x}`].v
              ) {
                workbook.Sheets[workbook.SheetNames[0]][`M${x}`] = {
                  t: "n",
                  v: Math.round(
                    ((workbook.Sheets[workbook.SheetNames[0]][`M${x}`].v +
                      elem[`${type}`] * number) /
                      2 +
                      (((workbook.Sheets[workbook.SheetNames[0]][`M${x}`].v +
                        elem[`${type}`] * number) /
                        2) *
                        ((workbook.Sheets[workbook.SheetNames[0]][`M${x}`].v -
                          elem[`${type}`] * number) /
                          2)) /
                        this.state.average) *
                      (this.state.average / this.state.mrpSum)
                  )
                };
              }
            });
          }
        }
        XLSX.writeFile(workbook, `calc-${type}-${number}.xlsx`);
      };

      reader.readAsArrayBuffer(file);
    }
  };

  render() {
    const { classes } = this.props;
    return (
      <div className="App">
        <header>
          <h1>MRP Calculator</h1>
        </header>
        <form>
          <TextField
            id="number"
            label="Target"
            type="number"
            className={classes.number}
            InputLabelProps={{
              shrink: true
            }}
            margin="normal"
            variant="outlined"
            helperText="Amount to distribute."
          />
          <br />
          <TextField
            id="type"
            select
            label="Type"
            className={classes.type}
            value={this.state.type}
            onChange={this.handleChange}
            SelectProps={{
              MenuProps: {
                className: classes.menu
              }
            }}
            helperText="Please select drug or cosm."
            margin="normal"
            variant="outlined"
          >
            <MenuItem key={"drug"} value={"drug"}>
              drug
            </MenuItem>
            <MenuItem key={"cosm"} value={"cosm"}>
              cosm
            </MenuItem>
          </TextField>
          <br />
          <br />
          <br />
          <label htmlFor="file" onKeyDown={this.handleKeyDown} className="typeBorder" tabIndex="0">
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="17" viewBox="0 0 20 17"><path fill="#000000de" d="M10 0l-5.2 4.9h3.3v5.1h3.8v-5.1h3.3l-5.2-4.9zm9.3 11.5l-3.2-2.1h-2l3.4 2.6h-3.5c-.1 0-.2.1-.2.1l-.8 2.3h-6l-.8-2.2c-.1-.1-.1-.2-.2-.2h-3.6l3.4-2.6h-2l-3.2 2.1c-.4.3-.7 1-.6 1.5l.6 3.1c.1.5.7.9 1.2.9h16.3c.6 0 1.1-.4 1.3-.9l.6-3.1c.1-.5-.2-1.2-.7-1.5z"></path></svg>
          <input type="file" id="file" accept="*.xlsx" onChange={this.handleFileSelect} className="type" required />
          <span>Choose a file</span>
          </label>
          <br />
          <br />
          <br />
          <Button
            variant="contained"
            size="large"
            onClick={this.handleClick}
            color="primary"
            className={classes.button}
          >
            Calculate
          </Button>
        </form>
      </div>
    );
  }
}

export default withStyles(styles)(App);
