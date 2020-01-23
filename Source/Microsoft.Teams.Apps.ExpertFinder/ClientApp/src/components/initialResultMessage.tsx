// <copyright file="initialResultMessage.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { MSTeamsIcon, MSTeamsIconWeight, MSTeamsIconType } from "msteams-ui-icons-react";

interface IInitialResultMessageProps {
	InitialResultMessageHeaderText: string,
	InitialResultMessageBodyText: string,
}

const InitialResultMessage: React.FunctionComponent<IInitialResultMessageProps> = props => {

	return (
		<div>
			<div className="initial-result-message-container">
				<div className="initial-result-message-icon">
					<MSTeamsIcon
						className="result-message-filter-icon"
						iconType={MSTeamsIconType.PresenceUnknown}
						iconWeight={MSTeamsIconWeight.Light} />
				</div>

				<div className="initial-result-message-text">
					<div className="initial-message-header">
						{props.InitialResultMessageHeaderText}
					</div>
					<div>
						{props.InitialResultMessageBodyText}
					</div>
				</div>
			</div>
		</div>
		);
}

export default InitialResultMessage;