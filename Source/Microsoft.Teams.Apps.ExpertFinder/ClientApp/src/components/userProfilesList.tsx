// <copyright file="userProfilesList.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { IUserProfile } from "./searchResultInitialMessage";
import { Checkbox } from "@stardust-ui/react";

interface IUserProfilesListProps {
	searchResultList: IUserProfile[];
	selectedProfiles: IUserProfile[];
	onCheckboxSelected: (profile: IUserProfile, status: boolean) => void,
	SkillsLabel: string
}

const UserProfilesList: React.FunctionComponent<IUserProfilesListProps> = props => {

	/**
	* Used in checkbox component to decide whether checkbox is checked or not.
	* @param  {IUserProfile} value Selected user profile
	*/
	function isCheckboxChecked(value: IUserProfile) {
		const selectedProfile = props.selectedProfiles.filter(userProfile => userProfile.preferredName === value.preferredName);

		if (selectedProfile.length) {
			return true;
		} else {
			return false;
		}
	}

	/**
	* Notify parent component that profile selection change
	* @param  {IUserProfile} value Selected user profile
	*/
	function onCheckboxChecked(value: IUserProfile) {
		let isProfileSelected = isCheckboxChecked(value)
		props.onCheckboxSelected(value, !isProfileSelected);

	}

	let profilesNamesList = props.searchResultList.map((item, key) => {
		return (<div key={key}>
			<div className="user-profile-container">
				<div className="user-profile-checkbox">
					<Checkbox
						onChange={() => onCheckboxChecked(item)}
						checked={isCheckboxChecked(item)} />
				</div>
				<div className="user-profile-content">
					<div className="user-profile-nametext">{item.preferredName} </div>
					<div>{item.jobTitle}</div>
					<div className="user-profile-content-skills" title={item.skills}>{props.SkillsLabel}: {item.skills}</div>
				</div>
			</div>
		</div>);
	});
	return (
		<div>
			{profilesNamesList}
		</div>
	);
}

export default UserProfilesList;