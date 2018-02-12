import * as React from 'react';
import { connect } from 'react-redux';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import {ITAPeoplePicker,ITAPersonaProps} from './../Common/ITAPeoplePicker'
import { autobind } from '@uifabric/utilities';
import { Button } from 'office-ui-fabric-react/lib/Button';
import {SPHttpClient, HttpClient}  from '@microsoft/sp-http';
import {Environment, EnvironmentType } from '@microsoft/sp-core-library';

export interface PermissionFormProps {
    spContext:IWebPartContext
};
export interface PermissionFormState {};
class PermissionForm extends React.Component<PermissionFormProps, PermissionFormState> {
    public render(): JSX.Element {
        return (<div>
            <h2>ITA Active Directory Permissions</h2>
            <p>Disclaimer : Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis suscipit volutpat laoreet. Nullam vulputate nisl eget aliquet sodales. In nec leo libero. Sed hendrerit est felis, vel laoreet metus gravida sit amet. Fusce neque ligula, mattis vitae elit sollicitudin, malesuada laoreet metus. Phasellus lectus nunc, tincidunt sit amet congue eu, feugiat ac lorem. Proin elit urna, tristique vitae fermentum quis, elementum ac erat.</p>
            
            Employee Name : <br/>
            <ITAPeoplePicker  spContext={this.props.spContext} onChange={(a) => { this.handleITAPickerChange(a);}} itemLimit={1} />
            <br/>
            Permission Sets : <br/>
            <br/>
            <Button text="Call Graph API" onClick={()=>{this.callGraphAPI()}}  />
            <br/>
            <br/>            
            <iframe src="https://functionapp1abcdefgh.azurewebsites.net"></iframe>
            <br/>
            <br/>

        </div>);
    }
    @autobind
    private callGraphAPI()
    {
        if(Environment.type == EnvironmentType.SharePoint)
        {
            this.getGraphData().then((a)=>{
                console.log("A : ",a );
            })
        }
        else
        {
            alert('Please use production SharePoint Online site');
        }
    }

private async getGraphData()
    {
        var userAccount = "boraakkaya";
        var propertyValue = "New Value from App 666";
        this.props.spContext.httpClient.fetch(`https://functionapp1abcdefgh.azurewebsites.net/api/HttpTriggerCSharp3?userAccount=johndoe&propertyValue=ValueFromSPFX&name=Bora2`,HttpClient.configurations.v1, { 
           credentials: "include",
           mode:'cors'
        }).then(async(result)=>{
            console.log("Result ", result);            
            await result.json().then((data)=>{                
                console.log("DSadsad ",data);
                return "Some DATA";
            });
        },
        (error)=>{
            return error;
        }
    );
    }
    @autobind
    private handleITAPickerChange(val)
    {
        console.log(val);
    }
}

const mapStateToProps = (state)=>
{
    return {spContext : state.spContext};
};

export default connect(mapStateToProps)(PermissionForm);

