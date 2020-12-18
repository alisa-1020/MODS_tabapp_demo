// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import { MODS } from "teamsauth";
import { Client } from "@microsoft/microsoft-graph-client";

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component {
  constructor(props) {
    super(props)
    this.state = {
      userInfo: {},
      profile: {},
      photoObjectURL: "",
      showLoginBtn: false,
      manager: {},
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  async componentDidMount() {
    // Next steps: Error handling using the error object
    await this.initMODS();

    await this.callGraphSilent();
    // Test Case 29
    // Test auth provider api
    // uncomment the following code. and comment the previous line code.
    // await this.callAuthProviderClient();
  }

  async initMODS() {
    // test case 25
    // check MODS.init called before all other client site sdk Apis.
    // uncomment the following code.
    await MODS.init("https://wenyutangtabapp1-runtime-293280.azurewebsites.net", "auth-start.html", "https://wenyutangtabapp1-be-293280.azurewebsites.net");
    
    var userInfo = MODS.getUserInfo();
    
    this.setState({
      userInfo: userInfo
    });
  }

  async callGraphSilent() {
    try {
      var graphClient = await MODS.getMicrosoftGraphClient();
      
      // // test case 27
      // // teams app need graph client.
      // // uncomment the following code.
      // graphClient = null;

      var profile = await graphClient.api("/me").get();
      var location = await graphClient.api("/location").get();
      var photoBlob = await graphClient.api("/me/photos('120x120')/$value").get();
      
      this.setState({
        profile: profile,
        photoObjectURL: URL.createObjectURL(photoBlob),
      });

      // // test case 26
      // // graph client display user manager by calling api ("/me/manager")
      // // uncomment the folloing code
      // var manager = await graphClient.api("/me/manager").get();
      // this.setState({
      //   manager: manager,
      // });
    }
    catch (err) {
      alert("You need to click login button to consent the access: " + err.message);
      this.setState({
        showLoginBtn: true
      });
    }
  }

  async callAuthProviderClient() {
    try {
      var authProvider = MODS.getMicrosoftGraphAuthProvider();
      var client = Client.init({
        defaultVersion: "v1.0",
        debugLogging: true,
        authProvider: authProvider
      })
      var profile = await client.api("/me").get();
      // var location = await client.api("/location").get();
      var photoBlob = await client.api("/me/photos('120x120')/$value").get();

      this.setState({
        profile: profile,
        photoObjectURL: URL.createObjectURL(photoBlob),
      });
    }
    catch(err) {
      alert();
      this.setState({
        showLoginBtn: true
      });
    }
  }

  async loginBtnClick() {
    try {
        await MODS.popupLoginPage();
    }
    catch(err) {
        alert("Login failed: " + err);
        return;
    }

    await this.callGraphSilent();
    // test case 28
    // uncomment the following code and comment the previous line code
    // await this.callAuthProviderClient();
  }

  async callFunction() {
    await MODS.callFunction("httpTrigger", "post", "hello");
  }

  render() {
    return (
      <div>
        <h2>Basic info from SSO</h2>
        <p><b>Name:</b> {this.state.userInfo.userName}</p>
        <p><b>E-mail:</b> {this.state.userInfo.preferredUserName}</p>

        {this.state.showLoginBtn && <button onClick={() => this.loginBtnClick()}>Auth and show</button>}

        <p>
          <h2>Profile from Microsoft Graph</h2>
          <div>
            <div><b>Name:</b> {this.state.profile.displayName}</div>
            <div><b>Job title:</b> {this.state.profile.jobTitle}</div>
            <div><b>E-mail:</b> {this.state.profile.mail}</div>
            <div><b>UPN:</b> {this.state.profile.userPrincipalName}</div>
            <div><b>Object id:</b> {this.state.profile.id}</div>
            <div><b>Office Localtiuon: </b>{this.state.profile.localtion}</div>
            {/* test case  26 
                display all users info
                uncommand the following code.*/}
            {/* <div><b>SurName:</b> {this.state.profile.surname} </div>
            <div><b>Given Name:</b> {this.state.profile.givenName} </div>
            <div><b>Mobile Phone:</b> {this.state.profile.mobilePhone} </div>
            <div><b>Preferred Language:</b> {this.state.profile.preferredLanguage} </div> */}

            {/* test case 27
            display user manager infomation */}
            {/* <div><b>manager display name:</b>{this.state.manager.displayName}</div>
            <div><b>manager job title:</b> {this.state.manager.jobTitle} </div>
            <div><b>manager mail:</b>{this.state.manager.mail} </div> */}
          </div>
        </p>

        <p>
          <h2>User Photo from Microsoft Graph</h2>
          <div >
            {this.state.photoObjectURL && <img src={this.state.photoObjectURL} alt="" />}
          </div>
        </p>

      </div>
    );
  }
}
export default Tab;