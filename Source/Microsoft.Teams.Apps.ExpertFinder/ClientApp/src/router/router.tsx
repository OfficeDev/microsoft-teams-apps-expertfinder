// <copyright file="router.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import { ProfileSearchWrapperPage } from "../components/searchUserWrapperPage";
import { ErrorPage } from "../components/errorPage";

export const AppRoute: React.FunctionComponent<{}> = () => {

	return (
		<BrowserRouter>
			<Switch>
				<Route exact path="/" component={ProfileSearchWrapperPage} />
				<Route exact path="/errorpage" component={ErrorPage} />
				<Route exact path="/errorpage/:id" component={ErrorPage} />
			</Switch>
		</BrowserRouter>

	);
};


