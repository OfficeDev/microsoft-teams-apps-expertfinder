// <copyright file="searchUserWrapperPage.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import ProfileSearchTextBoxComponent from "./profileSearchTextBox";
import { IUserProfile, SearchResultMessage, ISelectedFilter } from "./searchResultInitialMessage";
import FilterPopUp from "./filterPopUp";
import Resources from "../constants/resources";
import { getUserProfiles, getResourceStrings } from "../api/profileSearchApi";
import { FilterNamesComponent } from "./filterNamesComponent";
import { Provider, themes, Loader, Text, Button } from "@stardust-ui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import { ReactPlugin } from "@microsoft/applicationinsights-react-js";
import { createBrowserHistory } from "history";

import "../styles/site.css";
import "../styles/userProfile.css";

const browserHistory = createBrowserHistory({ basename: "" });
var reactPlugin = new ReactPlugin();

export interface ProfileSearchPageState {
    loader: boolean;
    isSearching: boolean;
    searchResults: IUserProfile[];
    selectedProfiles: IUserProfile[];
    isResourceAvailable: boolean,
    selectedFilterValues: ISelectedFilter[];
    showProfilesCountErrorMessage: boolean;
    resourceStrings: any,
    theme: string
}

export interface IAppSettings {
    token: string,
    telemetry: string,
    theme: string
}

export class ProfileSearchWrapperPage extends React.Component<{}, ProfileSearchPageState> {

    token?: string | null = null;
    telemetry?: any = null;
    theme?: string | null;
    appInsights: ApplicationInsights;

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.token = params.get("token");
        this.telemetry = params.get("telemetry");
        this.theme = params.get("theme");

        // store application settings in local storage to use later to allow refreshing page from error page.
        if (this.token) {
            let expertFinderSettings: IAppSettings = {
                token: this.token ? this.token : "",
                telemetry: this.telemetry ? this.telemetry : "",
                theme: this.theme ? this.theme : ""
            };
            localStorage.setItem("appsettings", JSON.stringify(expertFinderSettings));
        }
        else {
            let appSettings: any = null;
            let storageValue = localStorage.getItem("appsettings");
            if (storageValue) {
                appSettings = JSON.parse(storageValue) as IAppSettings;
            }
            if (appSettings) {
                this.token = appSettings.token;
                this.theme = appSettings.theme;
                this.telemetry = appSettings.telemetry;
            }
        }

        try {
            this.appInsights = new ApplicationInsights({
                config: {
                    instrumentationKey: this.telemetry,
                    extensions: [reactPlugin],
                    extensionConfig: {
                        [reactPlugin.identifier]: { history: browserHistory }
                    }
                }
            });
            this.appInsights.loadAppInsights();
        } catch (e) {
            this.appInsights = new ApplicationInsights({
                config: {
                    instrumentationKey: undefined,
                    extensions: [reactPlugin],
                    extensionConfig: {
                        [reactPlugin.identifier]: { history: browserHistory }
                    }
                }
            });
            console.log(e);
        }

        this.state = {
            loader: false,
            isSearching: false,
            searchResults: [],
            selectedProfiles: [],
            isResourceAvailable: false,
            selectedFilterValues: [],
            showProfilesCountErrorMessage: false,
            resourceStrings: null,
            theme: this.theme ? this.theme : Resources.default
        }
    }

    /**
    * Used to initialize Microsoft Teams sdk
    */
    async componentDidMount() {
        microsoftTeams.initialize();
        let response = await getResourceStrings(this.token);

        if (response.status === 200 && response.data) {

            this.setState({
                loader: false,
                isSearching: false,
                searchResults: this.state.searchResults,
                resourceStrings: response.data,
                selectedProfiles: [],
                selectedFilterValues: [
                    {
                        value: Resources.SkillsValue,
                        label: response.data.skillsTitle
                    },
                    {
                        value: Resources.interestsValue,
                        label: response.data.interestTitle
                    },
                    {
                        value: Resources.schoolsValue,
                        label: response.data.schoolsTitle
                    }],
                isResourceAvailable: true,
                showProfilesCountErrorMessage: false
            });
        } else {
            this.appInsights.trackTrace({ message: `"getResourceStrings" - Request failed:`, severityLevel: SeverityLevel.Warning });
        }
    }

    /**
    * Get user profile details from api
    * @param  {String} searchText Selected user profile
    */
    handleUserSearch = (searchText: string) => {
        if (searchText) {
            this.setState({
                loader: true,
                isSearching: true,
                searchResults: [],
                selectedProfiles: [],
                selectedFilterValues: this.state.selectedFilterValues,
                showProfilesCountErrorMessage: false
            });

            this.getUserProfile(searchText);
        }
    }

    /**
    * Call api to get user profile datails
    * @param  {String} searchText Selected user profile
    */
    private getUserProfile = async (searchText: string) => {
        try {
            let selectedFilterValues: string[] = [];
            this.state.selectedFilterValues.forEach(filter => {
                selectedFilterValues.push(filter.value);
            });
            let response = await getUserProfiles(searchText, selectedFilterValues, this.token);

            if (response.status === 200 && response.data) {

                this.setState({
                    loader: false,
                    isSearching: true,
                    searchResults: response.data,
                    selectedProfiles: [],
                    selectedFilterValues: this.state.selectedFilterValues,
                    showProfilesCountErrorMessage: false
                });
            } else {
                this.appInsights.trackTrace({ message: `"getUserProfile" - Request failed:`, severityLevel: SeverityLevel.Warning });

                this.setState({
                    loader: false,
                    isSearching: false,
                    searchResults: [],
                    selectedProfiles: [],
                    selectedFilterValues: this.state.selectedFilterValues,
                    showProfilesCountErrorMessage: false
                });
            }
        }
        catch (error) {
            this.appInsights.trackException(error);
            this.appInsights.trackTrace({ message: `"getUserProfile" - Request failed:`, severityLevel: SeverityLevel.Warning });
            console.error(error);
        }
    }

    /**
    * Submit task module task and pass selected user profile details to bot
    */
    private onViewButtonClicked = () => {
        let toBot = {
            command: Resources.UserSearchBotCommand,
            searchresults: this.state.selectedProfiles
        };
        microsoftTeams.tasks.submitTask(toBot);
        this.setState({
            loader: true,
        });
    }

    /**
    * Update user selected filters collection when user changes filter
     * @param  {String Array} values User selected filter names
    */
    private onFilterSelectionChange = (values: string[]) => {

        let selectedValues: ISelectedFilter[] = [];

        values.map(value => {
            switch (value) {
                case Resources.schoolsValue:
                    selectedValues.push(
                        {
                            value: Resources.schoolsValue,
                            label: this.state.resourceStrings.schoolsTitle
                        });
                    break;
                case Resources.SkillsValue:
                    selectedValues.push(
                        {
                            value: Resources.SkillsValue,
                            label: this.state.resourceStrings.skillsTitle
                        });
                    break;
                case Resources.interestsValue:
                    selectedValues.push(
                        {
                            value: Resources.interestsValue,
                            label: this.state.resourceStrings.interestTitle
                        });
                    break;
            }
        });

        this.setState(
            {
                isSearching: this.state.isSearching,
                searchResults: this.state.searchResults,
                selectedProfiles: this.state.selectedProfiles,
                selectedFilterValues: selectedValues,
                showProfilesCountErrorMessage: false
            });
    }

    /**
    * Update user selected filters collection when user removes filter
    * @param  {ISelectedFilters} selectedFilter User selected filter names
    */
    private onFilterNameRemoved = (selectedFilter: ISelectedFilter) => {

        let filteredNames = this.state.selectedFilterValues.filter((filter) => {
            return filter.value !== selectedFilter.value;
        });

        this.setState(
            {
                isSearching: this.state.isSearching,
                searchResults: this.state.searchResults,
                selectedProfiles: this.state.selectedProfiles,
                selectedFilterValues: filteredNames,
                showProfilesCountErrorMessage: false
            });
    }

    onUserProfileSelected = (profile: IUserProfile, status: boolean) => {
        let allProfiles = this.state.selectedProfiles;
        let isProfileMaxLimitReached = false;

        if (status) {
            let filteredProfiles = allProfiles.filter(
                userProfile => (userProfile.preferredName === profile.preferredName));

            if (filteredProfiles.length < 1) {

                if (allProfiles.length < Resources.MaxUserProfileLimit) {
                    allProfiles.push(profile);
                }
                else {
                    isProfileMaxLimitReached = true;
                }
            }
        }
        else {
            allProfiles = allProfiles.filter(userProfile => userProfile.preferredName !== profile.preferredName);
        }

        this.setState(
            {
                isSearching: true,
                searchResults: this.state.searchResults,
                selectedProfiles: allProfiles,
                selectedFilterValues: this.state.selectedFilterValues,
                showProfilesCountErrorMessage: isProfileMaxLimitReached
            }
        );

    }

    /**
    * Renders the component
    */
    public render(): JSX.Element {

        const styleProps: any = {};
        switch (this.state.theme) {
            case Resources.dark:
                styleProps.theme = themes.teamsDark;
                break;
            case Resources.contrast:
                styleProps.theme = themes.teamsHighContrast;
                break;
            case Resources.light:
            default:
                styleProps.theme = themes.teams;
        }

        return (
            <div>
                {this.getWrapperPage(styleProps.theme)}
            </div>
        );
    }

    private getWrapperPage = (theme: any) => {
        if (!this.state.isResourceAvailable) {
            return (
                <Provider theme={theme}>
                    <div className="Loader">
                        <Loader />
                    </div>
                </Provider>
            );
        } else {
            return (
                <div>
                    <div className="search-textbox-header">
                    </div>
                    <div className="search-textbox-container">
                        <ProfileSearchTextBoxComponent selectSearchText={this.handleUserSearch} placeHolderText={this.state.resourceStrings.searchTextBoxPlaceholder} />
                    </div>
                    <div className="search-filter-container">
                        <Provider theme={theme}>
                            <FilterPopUp
                                selectedFilterValues={this.state.selectedFilterValues}
                                onFilterSelectionChange={this.onFilterSelectionChange}
                                SkillsLabel={this.state.resourceStrings.skillsTitle}
                                InterestsLabel={this.state.resourceStrings.interestTitle}
                                SchoolsLabel={this.state.resourceStrings.schoolsTitle}
                            />
                        </Provider>
                    </div>
                    <div className="selected-filters-container">
                        <div className="selected-filters-innercontainer">
                            <FilterNamesComponent selectedFilters={this.state.selectedFilterValues} onFilterRemoved={this.onFilterNameRemoved} />
                        </div>
                    </div>
                    <div className="search-profiles-seperator" />
                    {this.getUserProfileListComponent(theme)}
                </div>
            );
        }
    }

    private getUserProfileListComponent = (theme: any) => {

        if (this.state.loader) {
            return (
                <Provider theme={theme}>
                    <div className="Loader">
                        <Loader />
                    </div>
                </Provider>
            );
        } else {
            return (
                <div className="user-search-container">
                    <div className="user-search-list-view">
                        <SearchResultMessage
                            isSearching={this.state.isSearching}
                            searchResultList={this.state.searchResults}
                            selectedProfiles={this.state.selectedProfiles}
                            onCheckboxSelected={this.onUserProfileSelected}
                            InitialResultMessageHeaderText={this.state.resourceStrings.initialSearchResultMessageHeaderText}
                            InitialResultMessageBodyText={this.state.resourceStrings.initialSearchResultMessageBodyText}
                            NoSearchResultFoundMessage={this.state.resourceStrings.searchResultNoItemsText}
                            SkillsLabel={this.state.resourceStrings.skillsTitle} />
                    </div>
                    <div className="view-profile-button-container">
                        <div className="view-profile-inner-container">
                            {this.getUserProfileLimitError(this.state.showProfilesCountErrorMessage, theme)}
                            <div className="view-button-container">
                                <Button
                                    content={this.state.resourceStrings.viewButtonText}
                                    primary
                                    onClick={this.onViewButtonClicked}
                                    disabled={this.state.selectedProfiles.length < 1}
                                />
                            </div>
                        </div>
                    </div>
                </div>
            );
        }
    }

    private getUserProfileLimitError = (showProfilesCountErrorMessage: boolean, theme: any) => {

        if (showProfilesCountErrorMessage) {
            return (
                <Provider theme={theme}>
                    <div className="error-message-container">
                        <Text content={this.state.resourceStrings.maxUserProfilesError} className="profile-error-message" error size="medium" />
                    </div>
                </Provider>
            );
        }
        else {
            return (
                <div />
            );
        }
    }
}

