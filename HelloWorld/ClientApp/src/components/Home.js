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
            this.setState({ userName: context.userObjectId });
            this.setState({ userId: context.userObjectId });

            microsoftTeams.authentication.getAuthToken({
                successCallback: (token) => {
                    microsoftTeams.appInitialization.notifySuccess();
                    fetch(`api/graph/beta/me`,
                        {
                            method: "GET",
                            headers: {
                                "Authorization": `${token}`
                            }
                        })
                        .then(response => response.json())
                        .then(data => {
                            if (data.error) {
                                console.error("!!!ERROR!!!");
                                console.error(data.error);
                            }
                            else {
                                console.log(data);
                                this.setState({ userName: data.data.displayName });
                            }

                        })
                        .catch(error => {
                            console.error('Unable to get user info', error);

                        });
                },
                failureCallback: (error) => {
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        error,
                    });
                }
            });
        });


        if (this.state.userName === "Unknown") {

        }
    }

    render() {
        return (
            <div>
                <h4>Hello, {this.state.userName}</h4>
                <p style={{ wordWrap: "break-word", fontSize:"10px" }}></p>
            </div>
        );
    }
}
