import React, { Component } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";

export class Home extends Component {
    static displayName = Home.name;
    constructor(props) {
        super(props);
        this.state = {
            userId: "",
            token: "",
            users:[]
        };
    }

    componentDidMount() {
        microsoftTeams.initialize();

        microsoftTeams.authentication.getAuthToken({
            successCallback: (token) => {
                microsoftTeams.appInitialization.notifySuccess();
                fetch(`api/graph/beta/users`,
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
                            this.setState({ users: data.data.value });
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
    }

    render() {
        return (
            <div>
                <p style={{ wordWrap: "break-word", fontSize: "10px" }}>Current AD User count is {this.state.users.length}</p>
            </div>
        );
    }
}
