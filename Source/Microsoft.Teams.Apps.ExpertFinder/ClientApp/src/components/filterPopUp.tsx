// <copyright file="filterPopUp.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { MSTeamsIcon, MSTeamsIconWeight, MSTeamsIconType } from "msteams-ui-icons-react";
import { Flex, Popup } from "@stardust-ui/react";
import { ISelectedFilter } from "./searchResultInitialMessage";
import FilterCheckboxGroup from "./filterCheckboxGroup";

interface ITeamsSearchUserProps {
	selectedFilterValues: ISelectedFilter[],
	onFilterSelectionChange: (values: string[]) => void,
	SkillsLabel: string,
	InterestsLabel: string,
	SchoolsLabel: string
}

export default class FilterPopUp extends React.Component<ITeamsSearchUserProps, {}> {

	constructor(props: ITeamsSearchUserProps) {
		super(props);
	}

	/**
	* Notify parent component that filter selection change 
	* @param  {String Array} values Selected filters
	*/
	onGroupChecked = (values: string[]) => {
		this.props.onFilterSelectionChange(values);
	};

	/**
	* Renders the component
	*/
	public render(): JSX.Element {

		return (
			<Flex gap="gap.smaller">
				<Popup
					trapFocus
					trigger={
						<div className="search-filter-container">
							<MSTeamsIcon className="filter-icon" iconType={MSTeamsIconType.Filter} iconWeight={MSTeamsIconWeight.Light} />
						</div>
					}
					content={
						<FilterCheckboxGroup
							InterestsLabel={this.props.InterestsLabel}
							SchoolsLabel={this.props.SchoolsLabel}
							SkillsLabel={this.props.SkillsLabel}
							onFilterSelectionChange={this.onGroupChecked}
							selectedFilterValues={this.props.selectedFilterValues}
						/>
					}
				/>
			</Flex>
		);
	}
}


