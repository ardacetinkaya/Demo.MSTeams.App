import React, { Component } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";

export class Home extends Component {
    static displayName = Home.name;
    constructor(props) {
        super(props);
        this.state = {
            userId: "",
            userName: "Unknown",
            token: ""
        };
    }

    componentDidMount() {
        microsoftTeams.initialize();

        microsoftTeams.getContext((context) => {
            this.setState({ userId: context.userObjectId });


        });


        if (this.state.userName === "Unknown") {

        }
    }

    render() {
        return (
            <div>
                <p style={{ wordWrap: "break-word", fontSize:"10px" }}></p>
            </div>
        );
    }
}
