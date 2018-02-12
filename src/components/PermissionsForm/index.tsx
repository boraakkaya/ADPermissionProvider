import * as React from 'react';
import { connect } from 'react-redux';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { ITAPeoplePicker, ITAPersonaProps } from './../Common/ITAPeoplePicker'
import { autobind } from '@uifabric/utilities';
import { Button, PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField, ITextField } from 'office-ui-fabric-react/lib/TextField';
import { SPHttpClient, HttpClient } from '@microsoft/sp-http';
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
    isLoading:boolean;
};
class PermissionForm extends React.Component<PermissionFormProps, PermissionFormState> {
    azureWebsiteLoaded: boolean;
    constructor(props) {
        super(props);
        this.state = { userAccountEmail: "johndoe@boraakkaya.onmicrosoft.com", propertyValue: "Default Value", isLoading:true }
        this.executeOrDelayUntilAzureIFrameLoaded(()=>{
            this.setState({isLoading:false});
        })
    }
    public render(): JSX.Element {
        return (<div>

            {this.state.isLoading && <Spinner size={ SpinnerSize.large } label='Waiting for the Azure Web site to load...' ariaLive='assertive' />}
            {JSON.stringify(this.state)}
            <h2>ITA Active Directory Permissions</h2>
            <p>Disclaimer : Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis suscipit volutpat laoreet. Nullam vulputate nisl eget aliquet sodales. In nec leo libero. Sed hendrerit est felis, vel laoreet metus gravida sit amet. Fusce neque ligula, mattis vitae elit sollicitudin, malesuada laoreet metus. Phasellus lectus nunc, tincidunt sit amet congue eu, feugiat ac lorem. Proin elit urna, tristique vitae fermentum quis, elementum ac erat.</p>

            Employee Name : <br />
            <ITAPeoplePicker spContext={this.props.spContext} onChange={(a) => { this.handleITAPickerChange(a); }} itemLimit={1} />
            <br />
            Permission Sets : <br />
            <TextField onChanged={(e) => { console.log(e); this.setState({ propertyValue: e }); }} placeholder="Custom Active Directory Permission Value" />
            <br />
            <DefaultButton style={{backgroundColor:'ms-bgColor-themePrimary'}}  primary={true} disabled={this.state.isLoading} text="Call Graph API" onClick={() => { this.callFunctionAPI() }} />
            <br />
            <br />
            <iframe style={{display:'none'}} src="https://functionappitaadp.azurewebsites.net" onLoad={()=>{this.azureWebsiteLoaded = true}} ></iframe>
            <br />
            <br />
        </div>);
    }
    @autobind
    private callFunctionAPI() {
        if (Environment.type == EnvironmentType.SharePoint) {
            this.callFunctionAppITAADP().then((a) => {
                console.log("A : ", a);
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
    private async callFunctionAppITAADP() {
        this.props.spContext.httpClient.fetch(`https://functionappitaadp.azurewebsites.net/api/UpdateAAD?userAccountEmail=${this.state.userAccountEmail}&propertyValue=${this.state.propertyValue}`, HttpClient.configurations.v1, {
            credentials: "include",
            mode: 'cors'
        }).then(async (result) => {
            console.log("Result ", result);
            await result.json().then((data) => {
                console.log("DSadsad ", data);
                return "Some DATA";
            });
        },
            (error) => {
                return error;
            }
        );
    }
    @autobind
    private handleITAPickerChange(val) {
        console.log(val);
        this.setState({ userAccountEmail: val.length > 0 ? val[0].secondaryText : "" });
    }
}

const mapStateToProps = (state) => {
    return { spContext: state.spContext };
};

export default connect(mapStateToProps)(PermissionForm);

