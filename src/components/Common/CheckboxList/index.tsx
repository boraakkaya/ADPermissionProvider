import * as React from 'react';
import { connect } from 'react-redux';
import {
    Label, Checkbox, ICheckboxStyles,
    ICheckboxProps, autobind
} from 'office-ui-fabric-react';
export interface ICheckboxKeyValuePairObj {
    key: string,
    value: string
}
export interface CheckBoxListUCProps {
    /**
     * Callback returns all selected values as an Array
     */
    onUpdate: (val) => void
    /**
     * Array which will be used to render individual checkboxes
     */
    dataArray: Array<ICheckboxKeyValuePairObj>
    /**The Label.Title of the field */
    label: string
    /**Default selected items as an array type of ICheckboxKeyValuePairObj */
    defaultSelected: Array<ICheckboxKeyValuePairObj | any>
};
/**
 * Uncontrolled Checkbox List
 * It doesnt use state but callback to return all selected values
 */
export default class CheckBoxListUC extends React.Component<CheckBoxListUCProps, any> {
    public render(): JSX.Element {
        return (<div style={{ display: 'block' }}>
            <Label>{this.props.label}</Label>
            {this.props.dataArray.map((v: ICheckboxKeyValuePairObj, i) => {
                var isChecked = this.props.defaultSelected.filter((iv: string, ii) => { return iv.trim() == v.key.trim() }).length > 0 ? true : false;
                return <div key={i} style={{ width: '250px', display: 'inline-block',padding:'4px' }}>
                    <Checkbox
                        key={i}
                        label={v.value}
                        value={v.key}
                        defaultChecked={isChecked}
                        checked={isChecked}
                        onChange={(event, isChecked) => this.updateSelections(isChecked, v)}
                    /></div>
            })}
        </div>);
    }
    @autobind
    updateSelections(isChecked, v) {
        var checkedValues: string[] = this.props.defaultSelected;
        if (isChecked) {
            checkedValues.push(v.key);
            this.props.onUpdate(checkedValues);
        }
        else {
            checkedValues = checkedValues.filter((a, i) => {
                return a != v.key
            });
            this.props.onUpdate(checkedValues);
        }
    }
}


