import React, { Component } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";

import { Collapse, Container, Navbar, NavbarBrand, NavbarToggler, NavItem, NavLink } from 'reactstrap';
import { Link } from 'react-router-dom';
import './NavMenu.css';

export class NavMenu extends Component {
    static displayName = NavMenu.name;

    constructor(props) {
        super(props);

        this.toggleNavbar = this.toggleNavbar.bind(this);
        this.state = {
            collapsed: true,
            userName: "",
            token: ""
        };
    }

    componentDidMount() {
        microsoftTeams.initialize();

        microsoftTeams.getContext((context) => {
            this.setState({ userName: context.userObjectId });
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
    }

    toggleNavbar() {
        this.setState({
            collapsed: !this.state.collapsed
        });


    }

    render() {
        return (
            <header>
                <Navbar className="navbar-expand-sm navbar-toggleable-sm ng-white border-bottom box-shadow mb-3" light>
                    <Container>
                        <NavbarBrand tag={Link} to="/">Hello {this.state.userName}</NavbarBrand>
                        <NavbarToggler onClick={this.toggleNavbar} className="mr-2" />
                        <Collapse className="d-sm-inline-flex flex-sm-row-reverse" isOpen={!this.state.collapsed} navbar>
                            <ul className="navbar-nav flex-grow">
                                <NavItem>
                                    <NavLink tag={Link} className="text-dark" to="/">Home</NavLink>
                                </NavItem>
                                <NavItem>
                                    <NavLink tag={Link} className="text-dark" to="/counter">Counter</NavLink>
                                </NavItem>
                            </ul>
                        </Collapse>
                    </Container>
                </Navbar>
            </header>
        );
    }
}
