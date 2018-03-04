import * as React from 'react';
import { connect } from 'react-redux';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { ITAPeoplePicker, ITAPersonaProps } from './../Common/ITAPeoplePicker'
import { autobind } from '@uifabric/utilities';
import { Button, PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField, ITextField } from 'office-ui-fabric-react/lib/TextField';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
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
    dialog?: { active: boolean, title: string, subText: string },
    pickerUserName?: string;

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
    authIframe: HTMLIFrameElement = null;
    constructor(props) {
        super(props);
        this.state = { userAccountEmail: "", propertyValue: [], isLoading: true, dialog: { active: false, title: "", subText: "" } };
        this.executeOrDelayUntilAzureIFrameLoaded(() => {
            this.setState({ isLoading: false });
        });
        this.refreshFunctionAppToken();
    }
    public render(): JSX.Element {
        if (this.state.isLoading) {
            return (<div>
                <Spinner size={SpinnerSize.large} label='Loading page components...' ariaLive='assertive' />
                <iframe style={{ display: 'none' }} src="https://usercustomproperty.azurewebsites.net" onLoad={(e) => {
                    this.azureWebsiteLoaded = true
                }} ></iframe>
            </div>)
        }
        if (!this.state.isLoading) {
            return (<div className={styles.formDiv}>
                <div><Dialog
                    hidden={!this.state.dialog.active}
                    onDismiss={this.closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: this.state.dialog.title,
                        subText: this.state.dialog.subText
                    }}
                    modalProps={{
                        isBlocking: true,
                        containerClassName: 'ms-dialogMainOverride'
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={this.closeDialog} text='OK' />
                    </DialogFooter>
                </Dialog>
                </div>                
                <p>Disclaimer : Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis suscipit volutpat laoreet. Nullam vulputate nisl eget aliquet sodales. In nec leo libero. Sed hendrerit est felis, vel laoreet metus gravida sit amet. Fusce neque ligula, mattis vitae elit sollicitudin, malesuada laoreet metus. Phasellus lectus nunc, tincidunt sit amet congue eu, feugiat ac lorem. Proin elit urna, tristique vitae fermentum quis, elementum ac erat.</p>
                Employee Name : <br />
                <ITAPeoplePicker spContext={this.props.spContext} onChange={(a) => { this.handleITAPickerChange(a); }} itemLimit={1} />
                <br />
                <CheckboxListUC defaultSelected={this.state.propertyValue} label="Permissions" dataArray={[{ key: "Pub", value: "Publisher" }, { key: "Cont", value: "Contributor" }, { key: "Edi", value: "Editor" }, { key: "Appr", value: "Approver" }]} onUpdate={(val) => { this.handleCheckBoxUpdate(val) }} />
                <br />
                <br />
                <DefaultButton className={styles.submitbutton} primary={true} disabled={this.state.isLoading || this.state.userAccountEmail.length == 0} text="Update User Permissions" onClick={() => { this.callFunctionAPI() }} />
                <br />
                <br />
                <iframe style={{ display: 'none' }} src="https://usercustomproperty.azurewebsites.net" onLoad={(e) => {
                    this.azureWebsiteLoaded = true
                }} ></iframe>
                <iframe style={{ display: 'none' }} src="https://usercustomproperty.azurewebsites.net/.auth/login/aad/callback" ref={(e) => { this.authIframe = e }} ></iframe>
                <br />
                <br />
            </div>);
        }
    }
    autobind
    refreshFunctionAppToken() {
        if (this.authIframe != null) {
            setInterval(() => {
                console.log("Updating Source");
                this.authIframe.src = "https://usercustomproperty.azurewebsites.net/.auth/login/aad/callback";
            }, 300000);
        }
        else {
            setTimeout(() => {
                this.refreshFunctionAppToken();
            }, 300);
        }
    }
    @autobind
    closeDialog() {
        this.setState({ dialog: { active: false, title: "", subText: "" }});
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
                    //alert("Updated both AD and UPS custom property successfully!");
                    this.setState({ dialog: { active: true, title: "Success!", subText: "Updated both Active Directory and SharePoint UPS custom properties successfully!" } });
                }
                else {
                    //alert("An error occured while updating permission sets for this employee! Please check error logs..");
                    this.setState({ dialog: { active: true, title: "Error!", subText: "An error occured while updating permission sets for this employee! Please check error logs." } });
                }
            }, (err) => {
                console.log(err);
                //alert("An error occured while processing your request! Please contact your SharePoint administrators..");
                this.setState({ dialog: { active: true, title: "Error!", subText: "An error occured while calling remote API! Please check error logs." } });
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
        var requestBody = {
            userAccountEmail: this.state.userAccountEmail,
            propertyValue: customPropertyValue
        }
        await this.props.spContext.httpClient.fetch(`https://usercustomproperty.azurewebsites.net/api/UpdateProperty?userAccountEmail=${this.state.userAccountEmail}&propertyValue=${customPropertyValue}`, HttpClient.configurations.v1, {
            credentials: "include",
            mode: 'cors'
        }).then(async (result) => {
            await result.json().then((data: any) => {
                console.log("ResultObject ", data);
                resultObject = data;
            }
            );
        }),
            (error) => {
                console.log("Error : ", error);
                resultObject = error;
            }
        return resultObject;
    }
    @autobind
    private async handleITAPickerChange(val) {
        var pickerUserEmail = val.length > 0 ? val[0].secondaryText : "";
        var pickerUserName = val.length > 0 ? val[0].primaryText : "";
        //https://graph.microsoft.com/beta/users/boraakkaya@boraakkaya.onmicrosoft.com
        if (pickerUserEmail != "") {
            await this.getSelectedUserProperties(pickerUserEmail).then((res: any) => {
                console.log("Res", res);
                var propValue: string[] = res.extension_2dcfd4b97df04d62bfd57064d7db80c5_myCustomProperty1 != undefined ? res.extension_2dcfd4b97df04d62bfd57064d7db80c5_myCustomProperty1.split(',') : [];
                this.setState({ userAccountEmail: pickerUserEmail, propertyValue: propValue, pickerUserName: pickerUserName });
            });
        }
        else {
            this.setState({ userAccountEmail: "", propertyValue: [] });
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
