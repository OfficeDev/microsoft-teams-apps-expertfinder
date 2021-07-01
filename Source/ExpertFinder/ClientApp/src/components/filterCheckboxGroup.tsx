// <copyright file="filterCheckboxGroup.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import Resources from "../constants/resources";
import { ISelectedFilter } from "./searchResultInitialMessage";
import { Checkbox } from "@stardust-ui/react";

interface IUserCheckboxGroupProps {
	selectedFilterValues: ISelectedFilter[],
	onFilterSelectionChange: (values: string[]) => void,
	SkillsLabel: string,
	InterestsLabel: string,
	SchoolsLabel: string
}

const FilterCheckboxGroup: React.FunctionComponent<IUserCheckboxGroupProps> = props => {

	let userSelectedFilterValues = props.selectedFilterValues.map(filter => filter.value)

	function isCheckboxChecked(value: string) {

		const selectedFilters = userSelectedFilterValues.filter(filterValue => filterValue === value);

		if (selectedFilters.length) {
			return true;
		} else {
			return false;
		}
	}

	function onCheckboxChecked(value: string) {
		let isFilterSelected = userSelectedFilterValues.includes(value)
		let selectedFilterValues: string[] = [];
		if (isFilterSelected) {
			selectedFilterValues = userSelectedFilterValues.filter(filterValue => filterValue !== value)
		}
		else {
			selectedFilterValues = [...userSelectedFilterValues, value]
		}
		props.onFilterSelectionChange(selectedFilterValues);
	}

	return (
		<div>
			<div>
				<Checkbox onChange={() => onCheckboxChecked(Resources.SkillsValue)} label={props.SkillsLabel} checked={isCheckboxChecked(Resources.SkillsValue)} />
			</div>
			<div>
				<Checkbox onChange={() => onCheckboxChecked(Resources.interestsValue)} label={props.InterestsLabel} checked={isCheckboxChecked(Resources.interestsValue)} />
			</div>
			<div>
				<Checkbox onChange={() => onCheckboxChecked(Resources.schoolsValue)} label={props.SchoolsLabel} checked={isCheckboxChecked(Resources.schoolsValue)} />
			</div>
		</div>
	);
}

export default FilterCheckboxGroup;