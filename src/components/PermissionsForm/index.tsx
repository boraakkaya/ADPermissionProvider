import * as React from 'react';
import { connect } from 'react-redux';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { ITAPeoplePicker, ITAPersonaProps } from './../Common/ITAPeoplePicker'
import { autobind } from '@uifabric/utilities';
import { Button, PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField, ITextField } from 'office-ui-fabric-react/lib/TextField';
import { SPHttpClientConfiguration, SPHttpClient, HttpClient, IHttpClientOptions, GraphHttpClient, GraphHttpClientConfiguration, IGraphHttpClientOptions } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import CheckboxListUC from './../Common/CheckboxList';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
import { ICheckboxKeyValuePairObj } from './../Common/CheckboxList';
import styles from './../../webparts/adPermissionProvider/AdPermissionProviderWebPart.module.scss';
export interface PermissionFormProps {
    spContext: IWebPartContext
};
export interface PermissionFormState {
    userAccountEmail: string;
    propertyValue: string[];
    isLoading: boolean;
};
export enum promiseResultEnum {
    Success = "success",
    Error = "error"
}
export interface IAPIPromiseResult {
    graphResult: string,
    userProfileResult: promiseResultEnum
}
class PermissionForm extends React.Component<PermissionFormProps, PermissionFormState> {
    azureWebsiteLoaded: boolean;
    constructor(props) {
        super(props);
        this.state = { userAccountEmail: "", propertyValue: [], isLoading: true }
        this.executeOrDelayUntilAzureIFrameLoaded(() => {
            this.setState({ isLoading: false });
        })
    }
    public render(): JSX.Element {        
        if(this.state.isLoading)
        {
            return (<div>
                <Spinner size={SpinnerSize.large} label='Waiting for the Azure Web site to load...' ariaLive='assertive' />
                <iframe style={{ display: 'block' }} src="https://boratestfunction1.azurewebsites.net" onLoad={(e) => {
            this.azureWebsiteLoaded = true
        }} ></iframe>
        </div>)
        }
        if(!this.state.isLoading)
        {
        return (<div className={styles.formDiv}>
                        
            <h2>ITA Active Directory Dynamic Permissions</h2>
            {JSON.stringify(this.state)}
            <p>Disclaimer : Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis suscipit volutpat laoreet. Nullam vulputate nisl eget aliquet sodales. In nec leo libero. Sed hendrerit est felis, vel laoreet metus gravida sit amet. Fusce neque ligula, mattis vitae elit sollicitudin, malesuada laoreet metus. Phasellus lectus nunc, tincidunt sit amet congue eu, feugiat ac lorem. Proin elit urna, tristique vitae fermentum quis, elementum ac erat.</p>
            

            Employee Name : <br />
            <ITAPeoplePicker spContext={this.props.spContext} onChange={(a) => { this.handleITAPickerChange(a); }} itemLimit={1} />            
            <br />
            <CheckboxListUC defaultSelected={this.state.propertyValue} label="Permissions" dataArray={[{ key: "Pub", value: "Publisher" }, { key: "Cont", value: "Contributor" }, { key: "Edi", value: "Editor" }, { key: "Appr", value: "Approver" }]} onUpdate={(val) => { this.handleCheckBoxUpdate(val) }} />
            <br />
            <br />
            <DefaultButton className={styles.submitbutton} primary={true} disabled={this.state.isLoading || this.state.userAccountEmail.length==0} text="Update User Permissions" onClick={() => { this.callFunctionAPI() }} />
            <br />
            <br />
            <iframe style={{ display: 'none' }} src="https://boratestfunction1.azurewebsites.net" onLoad={(e) => {
            this.azureWebsiteLoaded = true
        }} ></iframe>
            <br />
            <br />
        </div>);
        }
    }
    @autobind
    handleCheckBoxUpdate(val: string[]) {
        this.setState({ propertyValue: val });
    }
    @autobind
    private callFunctionAPI() {
        if (Environment.type == EnvironmentType.SharePoint) {
            this.callFunctionAppITAADP().then((promiseResult: IAPIPromiseResult) => {
                console.log("promiseResult : ", promiseResult);
                if (promiseResult.graphResult == promiseResultEnum.Success && promiseResult.userProfileResult == promiseResultEnum.Success) {
                    alert("Updated both AD and UPS custom property successfully!");
                }
                else {
                    alert("An error occured while updating permission sets for this employee! Please check error logs..");
                }
            })
        }
        else {
            alert('Please use production SharePoint Online site');
        }
    }
    private executeOrDelayUntilAzureIFrameLoaded(func: Function): void {
        if (this.azureWebsiteLoaded) {
            func();
        } else {
            setTimeout((): void => { this.executeOrDelayUntilAzureIFrameLoaded(func); }, 100);
        }
    }

    private async callFunctionAppITAADP(): Promise<{}> {
        var resultObject = {};
        var customPropertyValue = this.state.propertyValue.length > 0 ? this.state.propertyValue.toString() : "";
        await this.props.spContext.httpClient.fetch(`https://boratestfunction1.azurewebsites.net/api/HttpTriggerCSharp1?userAccountEmail=${this.state.userAccountEmail}&propertyValue=${customPropertyValue}`, HttpClient.configurations.v1, {
            credentials: "include",
            mode: 'cors'
        }).then(async (result) => {
            console.log("Result ", result);
            await result.json().then((data: any) => {
                console.log("ResultObject ", data);
                resultObject = data;
            }
            );
        }),
            (error) => {
                resultObject = error;
            }
        return resultObject
    }

    @autobind
    private async handleITAPickerChange(val) {
        var pickerUserEmail = val.length > 0 ? val[0].secondaryText : "";
        //https://graph.microsoft.com/beta/users/boraakkaya@boraakkaya.onmicrosoft.com
        if (pickerUserEmail != "") {
            await this.getSelectedUserProperties(pickerUserEmail).then((res: any) => {
                console.log("Res", res);
                var propValue: string[] = res.extension_2dcfd4b97df04d62bfd57064d7db80c5_myCustomProperty1 != undefined ? res.extension_2dcfd4b97df04d62bfd57064d7db80c5_myCustomProperty1.split(',') : [];
                this.setState({ userAccountEmail: pickerUserEmail, propertyValue: propValue });
            });
        }
        else {
            this.setState({userAccountEmail:"", propertyValue: []});
        }
    }

    @autobind
    private async getSelectedUserProperties(userEmail): Promise<{}> {
        var resultObject = {};
        await this.props.spContext.graphHttpClient.get(`beta/users/${userEmail}`, GraphHttpClient.configurations.v1, {
        }).then(async (result) => {
            console.log("Result Graph", result);
            await result.json().then((data: any) => {
                console.log("ResultGraphObject ", data);
                resultObject = data;
            }
            );
        }),
            (error) => {
                resultObject = error;
            }
        return resultObject
    }



}

const mapStateToProps = (state) => {
    return { spContext: state.spContext };
};

export default connect(mapStateToProps)(PermissionForm);
