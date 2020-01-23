// <copyright file="profileSearchTextBox.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Input } from "@stardust-ui/react";
import Resources from "../constants/resources";


interface IInputControlProps {
	selectSearchText: (searchText: string) => void,
	placeHolderText: string
}

interface IInputControlState {
	value: string
}

export default class ProfileSearchTextBoxComponent extends React.Component<IInputControlProps, IInputControlState> {

	constructor(props: IInputControlProps) {
		super(props);
		this.state = { value: "" };
		this.handleChange = this.handleChange.bind(this);
		this.handleKeyPress = this.handleKeyPress.bind(this);
	}

	/**
	* Set State value of textbox input control
	* @param  {Any} e Event object
	*/
	handleChange(e: any) {
		this.setState({ value: e.target.value });
	}

	/**
	* Used to call parent search method on enter keypress in textbox
	* @param  {Any} event Event object
	*/
	handleKeyPress(event: any) {
		var keyCode = event.which || event.keyCode;
		if (keyCode === Resources.keyCodeEnter) {
			this.props.selectSearchText(event.target.value);
		}
	}

	/**
	* Renders the component
	*/
	public render() {
		return (
			<div className="search-textbox">
				<Input 
					fluid
					placeholder={this.props.placeHolderText}
					autoFocus
					required
					value={this.state.value}
					onChange={this.handleChange}
					onKeyUp={this.handleKeyPress}
					maxLength={30}
					aria-label={this.props.placeHolderText}
				/>
			</div>
		);
	}

}