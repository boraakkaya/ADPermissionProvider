import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartContext
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AdPermissionProviderWebPart.module.scss';
import * as strings from 'AdPermissionProviderWebPartStrings';

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Provider } from 'react-redux';
import store from './../../Store/';
import MainComponent from './../../components/MainComponent';
import { getCurrentContext } from '../../reducers/context';
import { SPUser } from '@microsoft/sp-page-context';

export interface IAdPermissionProviderWebPartProps {
  description: string;
}
declare const loggedInUser:SPUser;
export default class AdPermissionProviderWebPart extends BaseClientSideWebPart<IAdPermissionProviderWebPartProps> {
 
  public render(): void {
    const element = <Provider store={store}>
    <div>
      <h2 className="ms-font-xxl">ITA - Azure AD Permission Provider</h2>
      <MainComponent context={this.context} />
    </div>
  </Provider>;
  ReactDOM.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    console.log("CurrentStore : " , store.getState());
    store.dispatch(getCurrentContext(this.context)).then((data)=>{
      console.log("Datax : " , data);
      console.log("ContextState : " , store.getState());
      (window as any).loggedInUser = this.context.pageContext.user;
    });
    return super.onInit();
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
