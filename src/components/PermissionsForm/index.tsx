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
//userAccount propertyValue
//https://functionapp1abcdefgh.azurewebsites.net/api/HttpTriggerCSharp1
//https://functionapp1abcdefgh.azurewebsites.net/api/HttpTriggerCSharp1?userAccount=${userAccount}&propertyValue=${propertyValue}

//https://login.windows.net/cd1ed347-7cfb-48c9-981a-8ca2f80ba40f/oauth2/authorize?response_type=code+id_token&redirect_uri=https%3A%2F%2Ffunctionapp1abcdefgh.azurewebsites.net%2F.auth%2Flogin%2Faad%2Fcallback&client_id=7e3bb679-42ae-43c9-807b-a81f4168f6db&scope=openid+profile+email&response_mode=form_post&nonce=7bf82e3379764b20ac097ee72db95ad3_20180210232600&state=redir%3D%252Fapi%252FHttpTriggerCSharp1%253FuserAccount%253Dboraakkaya%2526propertyValue%253DNew%252520Value%252520from%252520App


//https://login.microsoftonline.com/cd1ed347-7cfb-48c9-981a-8ca2f80ba40f/oauth2/authorize?response_type=code+id_token&redirect_uri=https%3A%2F%2Ffunctionapp1abcdefgh.azurewebsites.net%2F.auth%2Flogin%2Faad%2Fcallback&client_id=7e3bb679-42ae-43c9-807b-a81f4168f6db&scope=openid+profile+email&response_mode=form_post&nonce=9f5a213a79284db1a791c2da2977704e_20180210232948&state=redir%3D%252Fapi%252FHttpTriggerCSharp1%253FuserAccount%253Dboraakkaya%2526propertyValue%253DNew%252520Value%252520from%252520App

private async getGraphData()
    {
        var userAccount = "boraakkaya";
        var propertyValue = "New Value from App 666";

        const requestHeaders: Headers = new Headers();
        //requestHeaders.append('credentials','true');
        //requestHeaders.append('Access-Control-Allow-Credentials', 'true');
        requestHeaders.append('Access-Control-Allow-Origin','https://boraakkaya.sharepoint.com');

        this.props.spContext.httpClient.fetch(`https://functionapp1abcdefgh.azurewebsites.net/api/HttpTriggerCSharp3?userAccount=johndoe&propertyValue=ValueFromSPFX&name=Bora2`,HttpClient.configurations.v1, { 
           credentials: "include",
           mode:'cors'
           //headers:requestHeaders
           //headers:{
           //'Access-Control-Allow-Origin':'https://boraakkaya.sharepoint.com'
          // }
        }).then(async(result)=>{
            console.log("Result ", result);
            //console.log(result.text());            
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

