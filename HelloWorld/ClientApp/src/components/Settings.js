import React, { Component } from 'react';

export class Settings extends Component {
  static displayName = Settings.name;

  constructor(props) {
    super(props);
    this.state = { theme: "default" };

  }



  render() {
    return (
      <div>
        <h1>Settings</h1>

        <p>This is a simple example of a Teams App setting.</p>

      </div>
    );
  }
}
