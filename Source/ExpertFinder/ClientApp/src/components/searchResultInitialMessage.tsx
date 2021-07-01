// <copyright file="searchResultInitialMessage.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import EmptySearchResultMessage from "./emptySearchResultMessage";
import InitialResultMessage from "./initialResultMessage";
import UserProfilesList from "./userProfilesList";
import "../styles/userProfile.css";

export interface IUserProfile {
	aboutMe: string,
	interests: string,
	jobTitle: string,
	path: string,
	preferredName: string,
	schools: string,
	skills: string,
	workEmail: string
}

export interface ISearchResultProps {
	isSearching: boolean;
	searchResultList: IUserProfile[];
	selectedProfiles: IUserProfile[];
	onCheckboxSelected: (profile: IUserProfile, status: boolean) => void,
	InitialResultMessageHeaderText: string,
	InitialResultMessageBodyText: string,
	SkillsLabel: string,
	NoSearchResultFoundMessage: string
}

export interface ISelectedFilter {
	value: string;
	label: string;
}

export class SearchResultMessage extends React.Component<ISearchResultProps> {

	constructor(props: ISearchResultProps) {
		super(props);
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element {

		if (!this.props.isSearching) {
			return (
				<InitialResultMessage
					InitialResultMessageBodyText={this.props.InitialResultMessageBodyText}
					InitialResultMessageHeaderText={this.props.InitialResultMessageHeaderText} />
			);
		}
		else if (this.props.searchResultList.length > 0) {
			return (
				<UserProfilesList searchResultList={this.props.searchResultList} selectedProfiles={this.props.selectedProfiles} SkillsLabel={this.props.SkillsLabel} onCheckboxSelected={this.props.onCheckboxSelected} />
			);
		}
		else {
			return (
				<EmptySearchResultMessage NoSearchResultFoundMessage={this.props.NoSearchResultFoundMessage} />
			)
		}
	}
};