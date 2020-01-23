// <copyright file="errorPage.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { RouteComponentProps, Link } from "react-router-dom";
import { Text, Loader } from "@stardust-ui/react";
import { IAppSettings } from "./searchUserWrapperPage";
import { getErrorResourceStrings } from "../api/profileSearchApi";
import "../styles/site.css";

interface IResourceString {
    unauthorizedErrorMessage: string,
    forbiddenErrorMessage: string,
    generalErrorMessage: string,
    refreshLinkText: string
}

interface errorPageState {
    loader: boolean;
    resourceStrings: IResourceString,
}

export class ErrorPage extends React.Component<RouteComponentProps, errorPageState> {
    private appSettings: IAppSettings = {
        telemetry: "",
        theme: "",
        token: ""
    };

    constructor(props: any) {
        super(props);
        this.state = {
            loader: true,
            resourceStrings: {
                unauthorizedErrorMessage: "Sorry, an error occurred while trying to access this service.",
                forbiddenErrorMessage: "Sorry, seems like you don't have permission to access this page.",
                generalErrorMessage: "Oops! An unexpected error seems to have occured. Why not try refreshing your page? Or you can contact your administrator if the problem persists.",
                refreshLinkText: "Refresh"
            }
        };
        let storageValue = localStorage.getItem("appsettings");
        if (storageValue) {
            this.appSettings = JSON.parse(storageValue) as IAppSettings;
        }
    }

    async componentDidMount() {
        let response = await getErrorResourceStrings(this.appSettings.token);

        if (response.status === 200 && response.data) {
            this.setState({
                loader: false,
                resourceStrings: response.data
            });
        }
        else {
            this.setState({
                loader: false
            });
        }
    }

    /**
* Renders the component
*/
    public render(): JSX.Element {

        const params = this.props.match.params;
        let message = `${this.state.resourceStrings.generalErrorMessage}`;

        if ("id" in params) {
            const id = params["id"];
            if (id === "401") {
                message = `${this.state.resourceStrings.unauthorizedErrorMessage}`;
            } else if (id === "403") {
                message = `${this.state.resourceStrings.forbiddenErrorMessage}`;
            }
            else {
                message = `${this.state.resourceStrings.generalErrorMessage}`;
            }
        }
        if (!this.state.loader) {
            return (
                <div>
                    <Text content={message} className="error-message" error size="medium" />
                    <Link to="/" hidden={this.appSettings ? false : true} className="error-message refresh-page-link"> {this.state.resourceStrings.refreshLinkText} </Link>
                </div>
            );
        }
        else {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        }
    }
}