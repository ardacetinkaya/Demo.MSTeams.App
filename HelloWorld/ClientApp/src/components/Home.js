import React, { Component } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { webPubSub } from "@azure/web-pubsub";

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
                    this.setState({ token: token });

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
                                console.error(data.error);
                            }
                            else {
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
                <p>{this.state.token}</p>
            </div>
        );
    }
}
