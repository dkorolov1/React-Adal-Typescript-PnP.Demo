import './App.css';

import * as React from 'react';

import logo from './logo.svg';

import { Web } from "@pnp/sp";
import adalContext, { endpoint } from './adalConfig';

interface IAppState {
  webTitle: string
}

class App extends React.Component<{}, IAppState> {
  constructor(props: any) {
    super(props);
    this.state = {webTitle: ''};
    this.onLogOut = this.onLogOut.bind(this);
  }

  public componentWillMount() {
    const web = new Web(endpoint);
    web.select("Title").get().then(w => {
      this.setState({
        webTitle : w.Title
      });
    });
  }

  public onLogOut() {
    adalContext.LogOut();
  }

  public render() {
    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <h1 className="App-title">Welcome to React</h1>
        </header>
        <p className="App-intro">
          To get started, edit <code>src/App.tsx</code> and save to reload.
        </p>
        <h1>
          SharePoint Web Title: {this.state.webTitle}
        </h1>
        <button onClick={this.onLogOut}>
          Log out
        </button>
      </div>
    );
  }
}

export default App;
