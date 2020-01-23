// <copyright file="axiosDecorator.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios, { AxiosResponse, AxiosRequestConfig } from "axios";

export class AxiosJWTDecorator {

	/**
	* Post data to api
	* @param  {String} url Resource uri
	* @param  {Object} data Request body data
	* @param  {String} token Custom jwt token
	*/
	public async post<T = any, R = AxiosResponse<T>>(
		url: string,
		data?: any,
		token?: string
	): Promise<R> {
		try {
			let config: AxiosRequestConfig = axios.defaults;
			config.headers["Authorization"] = `Bearer ${token}`;

			return await axios.post(url, data, config);
		} catch (error) {
			this.handleError(error);
			throw error;
		}
	}

	/**
	* Get data to api
	* @param  {String} token Custom jwt token
	*/
	public async get<T = any, R = AxiosResponse<T>>(
		url: string,
		token?: string
	): Promise<R> {
		try {
			let config: AxiosRequestConfig = axios.defaults;
			config.headers["Authorization"] = `Bearer ${token}`;
			return await axios.get(url, config);
		} catch (error) {
			this.handleError(error);
			throw error;
		}
	}

	/**
	* Handle error occured during api call.
	* @param  {Object} error Error response object
	*/
	private handleError(error: any): void {
		if (error.hasOwnProperty("response")) {
			const errorStatus = error.response.status;
			if (errorStatus === 403) {
				window.location.href = "/errorpage/403";
			} else if (errorStatus === 401) {
				window.location.href = "/errorpage/401";
			} else {
				window.location.href = "/errorpage";
			}
		} else {
			window.location.href = "/errorpage";
		}
	}

}

const axiosJWTDecoratorInstance = new AxiosJWTDecorator();
export default axiosJWTDecoratorInstance;