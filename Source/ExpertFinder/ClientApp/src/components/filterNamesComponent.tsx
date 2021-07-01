// <copyright file="filterNamesComponent.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { MSTeamsIcon, MSTeamsIconWeight, MSTeamsIconType } from "msteams-ui-icons-react";
import Resources from "../constants/resources";
import "../styles/userProfile.css";
import { ISelectedFilter } from "./searchResultInitialMessage";

interface IFilterNamesComponentProps {
	selectedFilters: ISelectedFilter[],
	onFilterRemoved: (selectedFilter: ISelectedFilter) => void
}

export class FilterNamesComponent extends React.Component<IFilterNamesComponentProps> {

	constructor(props: IFilterNamesComponentProps) {
		super(props);
	}

	/**
	* Remove filter from selected filter collection
	* @param  {ISelectedFilter} filter User entered search text
	*/
    private onCloseClick = (filter: ISelectedFilter) => {
		this.props.onFilterRemoved(filter);
	}

	/**
	* Remove filter from selected filter collection
	* @param  {ISelectedFilter} filter User entered search text
	* @param  {Object} event Event object
	*/
    private onCloseKeyPress = (filter: ISelectedFilter, event) => {
		var keyCode = event.which || event.keyCode;
		if (keyCode === Resources.keyCodeEnter || keyCode === Resources.keyCodeSpace) {
			this.onCloseClick(filter);
		}
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element[] {

		return (this.props.selectedFilters.map((filter, key) => {
			return (
				<div key={key} className={"filter-name-block-container"}>
					<div>
						<div className={"filter-name-block"}>
							<div className={"filter-name-text"}>{filter.label}</div>
                            <span className={"filter-name-close-button" + " " + `${filter.label}keyPress`} tabIndex={0}
								role="button"
								onClick={() => this.onCloseClick(filter)}
								onKeyDown={(event) => this.onCloseKeyPress(filter, event)}>
								<MSTeamsIcon
									style={this.styles.icon}
									iconType={MSTeamsIconType.ChromeClose}
									iconWeight={MSTeamsIconWeight.Light} />
							</span>
						</div>
					</div>
				</div>
			);
		})
		)
	}

	styles = {
		icon: {
			fontSize: "1rem"
		},
	}
}