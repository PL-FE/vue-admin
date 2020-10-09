import React, {Component} from 'react';
import GC from '@grapecity/spread-sheets'
import {SpreadSheets, Worksheet, Column} from '@grapecity/spread-sheets-react';
import './Style.css'
import dataService from '../dataService'

class StyleCon extends Component {

    constructor(props) {
        super(props);
        this.hostStyle = {
            top: '90px',
            bottom: '0px'
        };
        this.checkBoxCellType = new GC.Spread.Sheets.CellTypes.CheckBox();
        this.hyperLinkCelleType = new GC.Spread.Sheets.CellTypes.HyperLink();
        this.comboBoxCellType = new GC.Spread.Sheets.CellTypes.ComboBox();
        this.style = new GC.Spread.Sheets.Style();
        this.style.backColor = 'lightgray';
        this.comboBoxCellType.items([
            {text: 'US', value: 'US'},
            {text: 'UK', value: 'UK'},
            {text: 'Germany', value: 'Germany'},
            {text: 'Maxico', value: 'Maxico'}]);
        this.autoGenerateColumns = false;
        this.data = dataService.getEmployeesData();
    }

    render() {
        return (
            <div className="componentContainer" style={this.props.style}>
                <h3>样式</h3>
                <div>
                    <p>以下示例展示如何使用样式。</p>
                </div>
                <div className="spreadContainer" style={this.hostStyle}>
                    <SpreadSheets>
                        <Worksheet dataSource = {this.data} autoGenerateColumns={this.autoGenerateColumns}>
                            <Column dataField="name" width={150} style={this.style}/>
                            <Column dataField="phone" width={150}/>
                            <Column dataField="country" width={100} cellType = {this.comboBoxCellType}/>
                            <Column dataField="email" width={240} cellType = {this.hyperLinkCelleType}/>
                            <Column dataField="onJob" width={100} cellType = {this.checkBoxCellType}/>
                        </Worksheet>
                    </SpreadSheets>
                </div>
            </div>

        )
    }
}

export default StyleCon