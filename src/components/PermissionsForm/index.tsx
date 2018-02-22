import * as React from 'react';
import { connect } from 'react-redux';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { ITAPeoplePicker, ITAPersonaProps } from './../Common/ITAPeoplePicker'
import { autobind } from '@uifabric/utilities';
import { Button, PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField, ITextField } from 'office-ui-fabric-react/lib/TextField';
import {SPHttpClientConfiguration, SPHttpClient, HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
export interface PermissionFormProps {
    spContext: IWebPartContext
};
export interface PermissionFormState {
    userAccountEmail: string;
    propertyValue: string;
    isLoading: boolean;
};
export enum promiseResultEnum{
    Success = "success",
    Error = "error"
}
export interface IAPIPromiseResult
{
    graphResult:string,
    userProfileResult:promiseResultEnum
}
class PermissionForm extends React.Component<PermissionFormProps, PermissionFormState> {
    azureWebsiteLoaded: boolean;
    constructor(props) {
        super(props);
        this.state = { userAccountEmail: "johndoe@boraakkaya.onmicrosoft.com", propertyValue: "Default Value", isLoading: true }
        this.executeOrDelayUntilAzureIFrameLoaded(() => {
            this.setState({ isLoading: false });
        })
    }
    public render(): JSX.Element {
        return (<div>

            {this.state.isLoading && <Spinner size={SpinnerSize.large} label='Waiting for the Azure Web site to load...' ariaLive='assertive' />}
            {JSON.stringify(this.state)}
            <h2>ITA Active Directory Permissions</h2>
            <p>Disclaimer : Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis suscipit volutpat laoreet. Nullam vulputate nisl eget aliquet sodales. In nec leo libero. Sed hendrerit est felis, vel laoreet metus gravida sit amet. Fusce neque ligula, mattis vitae elit sollicitudin, malesuada laoreet metus. Phasellus lectus nunc, tincidunt sit amet congue eu, feugiat ac lorem. Proin elit urna, tristique vitae fermentum quis, elementum ac erat.</p>

            Employee Name : <br />
            <ITAPeoplePicker spContext={this.props.spContext} onChange={(a) => { this.handleITAPickerChange(a); }} itemLimit={1} />
            <br />
            Permission Sets : <br />
            <TextField onChanged={(e) => { this.setState({ propertyValue: e }); }} placeholder="Custom Active Directory Permission Value" />
            <br />
            <DefaultButton style={{ backgroundColor: '#800000' }} primary={true} disabled={this.state.isLoading} text="Update User Permissions" onClick={() => { this.callFunctionAPI() }} />
            <br />
            {/*
            <br/>
            <DefaultButton style={{ backgroundColor: 'ms-bgColor-themePrimary' }} primary={true} disabled={this.state.isLoading} text="Update User Profile Property" onClick={() => { this.updateUPP() }} />
            <br/> 
            */}
            <br />
            <iframe style={{ display: 'block' }} src="https://boratestfunction1.azurewebsites.net" onLoad={() => { this.azureWebsiteLoaded = true }} ></iframe>
            <br />
            <br />
        </div>);
    }
    @autobind
    private callFunctionAPI() {
        if (Environment.type == EnvironmentType.SharePoint) {
            this.callFunctionAppITAADP().then((promiseResult:IAPIPromiseResult) => {
                console.log("promiseResult : ", promiseResult);
                if(promiseResult.graphResult == promiseResultEnum.Success && promiseResult.userProfileResult == promiseResultEnum.Success)
                {
                    alert("Updated both AD and UPS custom property successfully!");
                }
                else
                {
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
    @autobind
    private async updateUPP()
    {
        var contentBody = `{"accountName":"i:0#.f|membership|${this.state.userAccountEmail}","propertyName":"MyCustomProperty1","propertyValue":"${this.state.propertyValue}"}`
        this.props.spContext.spHttpClient.post('https://boraakkaya-admin.sharepoint.com/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty',SPHttpClient.configurations.v1,{body:contentBody}).then(async(result)=>{
            console.log("UPP Result ", result);
            await result.json().then((data)=>{
                console.log(data);
            })
        },(err)=>{
            console.log("Error occured : ", err);
        })
    }
    private async callFunctionAppITAADP():Promise<{}> {
        var resultObject = {};
        await this.props.spContext.httpClient.fetch(`https://boratestfunction1.azurewebsites.net/api/HttpTriggerCSharp1?userAccountEmail=${this.state.userAccountEmail}&propertyValue=${this.state.propertyValue}`, HttpClient.configurations.v1, {
            credentials: "include",
            mode: 'cors'
        }).then(async (result) => {
            console.log("Result ", result);
            await result.json().then((data: any) => {
                console.log("ResultObject ", data);
                resultObject =  data;
            }
            );
        }),
            (error) => 
            {
                resultObject =  error;
            }
            return resultObject
    }

    @autobind
    private handleITAPickerChange(val) {        
        this.setState({ userAccountEmail: val.length > 0 ? val[0].secondaryText : "" });
    }
}

const mapStateToProps = (state) => {
    return { spContext: state.spContext };
};

export default connect(mapStateToProps)(PermissionForm);
