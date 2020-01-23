// <copyright file="emptySearchResultMessage.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { MSTeamsIcon, MSTeamsIconWeight, MSTeamsIconType } from "msteams-ui-icons-react";

interface IEmptySearchResultMessageProps {
	NoSearchResultFoundMessage: string,
}

const EmptySearchResultMessage: React.FunctionComponent<IEmptySearchResultMessageProps> = props => {

	return (
		<div>
			<div className="initial-result-message-container">
				<div className="initial-result-message-icon">
					<MSTeamsIcon className="result-message-filter-icon" iconType={MSTeamsIconType.PresenceUnknown} iconWeight={MSTeamsIconWeight.Light} />
				</div>

				<div className="initial-result-message-text">
					<div>
						{props.NoSearchResultFoundMessage}
					</div>
				</div>
			</div>
		</div>
	);
}

export default EmptySearchResultMessage;