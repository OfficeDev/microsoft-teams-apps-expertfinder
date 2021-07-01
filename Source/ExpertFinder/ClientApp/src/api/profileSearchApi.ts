﻿// <copyright file="profileSearchApi.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axiosDecorator";

const baseAxiosUrl = window.location.origin;

/**
* Get user profiles from api
* @param  {String} searchText User entered search text
* @param  {String Array} filters User selected filters
* @param  {String | Null} token Custom jwt token
*/
export const getUserProfiles = async (searchText: string, filters: string[], token: any): Promise<any> => {

	let url = baseAxiosUrl + "/api/users";
	const data = {
		searchText: searchText,
		SearchFilters: filters
    };
    return await axios.post(url, data, token);
}

/**
* Get localized resource strings from api
*/
export const getResourceStrings = async (token: any, locale: string): Promise<any> => {

	let url = baseAxiosUrl + "/api/resources/strings";
	return await axios.get(url, token, locale);
}