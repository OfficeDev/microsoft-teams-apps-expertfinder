// <copyright file="app.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { AppRoute } from "./router/router";
import Resources from "./constants/resources";
import { Provider, themes } from "@stardust-ui/react";
import { RouteComponentProps } from "react-router-dom";

export interface IAppState {
	theme: string;
	themeStyle: any;
}

export default class App extends React.Component<{}, IAppState> {

	constructor(props: any) {
		super(props);
		this.state = {
			theme: "",
			themeStyle: themes.teams,
		}
	}

	/**
	* Initializes Microsft Teams sdk and get current theme from teams context
	*/
	public componentDidMount() {
		microsoftTeams.initialize();
		microsoftTeams.getContext((context) => {
			let theme = context.theme || "";
			this.updateTheme(theme);
			this.setState({
				theme: theme
			});
		});

		microsoftTeams.registerOnThemeChangeHandler((theme) => {
			this.updateTheme(theme);
			this.setState({
				theme: theme,
			}, () => {
				this.forceUpdate();
			});
		});
	}

	/**
	* Set current theme state received from teams context
	* @param  {String} theme Current theme name
	*/
	private updateTheme = (theme: string) => {
		if (theme === Resources.dark) {
			this.setState({
				themeStyle: themes.teamsDark
			});
		} else if (theme === Resources.contrast) {
			this.setState({
				themeStyle: themes.teamsHighContrast
			});
		} else {
			this.setState({
				themeStyle: themes.teams
			});
		}

		if (theme) {
			// Possible values for theme: "default", "light", "dark" and "contrast"
			document.querySelector(Resources.body)
			document.body.className = Resources.theme + "-" + (theme === Resources.default ? Resources.light : theme);
		}
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element {

		return (
			<Provider theme={this.state.themeStyle}>
				<div className="appContainer">
					<AppRoute />
				</div>
			</Provider>
		);
	}

}
