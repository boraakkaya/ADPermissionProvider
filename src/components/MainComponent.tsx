import * as React from 'react';
import { IWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import { autobind } from "office-ui-fabric-react/lib/Utilities";
import PermissionsForm from './PermissionsForm';
import styles from './../webparts/adPermissionProvider/AdPermissionProviderWebPart.module.scss'
export interface MainComponentProps {
    context:IWebPartContext | WebPartContext;
}
export interface MainComponentState {}
class MainComponent extends React.Component<MainComponentProps, MainComponentState> {
    
    public render(): JSX.Element {
            
        return (<div className={styles.adPermissionProvider}>
                <PermissionsForm />                     
            </div>);
    }    
}



export default MainComponent;
