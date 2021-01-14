// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import './Tab.css'
import { MODS } from "mods-client";
import { StorageClient } from './StorageClient';

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component {

  storageClient;

  constructor(props) {
    super(props)
    this.state = {
      userInfo: {},
      items: [],
      newItemContent: "",
      showLoginBtn: false,
      photoObjectURL: "",
      photoDescription: "",
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  async componentDidMount() {
    // Next steps: Error handling using the error object
    await this.initMODS();
    await this.getData();
  }

  async initMODS() {
    var modsEndpoint = process.env.REACT_APP_MODS_ENDPOINT;
    var startLoginPageUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL;
    var functionEndpoint = process.env.REACT_APP_FUNC_ENDPOINT;
    // Initialize MODS SDK
    await MODS.init(modsEndpoint, startLoginPageUrl, functionEndpoint);
    // Get user information
    var userInfo = MODS.getUserInfo();
    this.setState({
      userInfo: userInfo
    });
  }

  async getData() {
    try {
      // Get Microsoft Graph client
      const graphClient = MODS.getMicrosoftGraphClient(['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All']);

      try {
        var photoBlob = await graphClient.api("/me/photos('120x120')/$value").get();
        this.setState({
          photoObjectURL: URL.createObjectURL(photoBlob),
          photoDescription: "avatar"
        });
      } catch (error) {
        this.setState({
          photoDescription: "no avatar found"
        });
      }

      this.storageClient = new StorageClient(graphClient);
      this.setState({
        items: await this.storageClient.getItems(),
        showLoginBtn: false
      });
    }
    catch (err) {
      alert("You need to click login button to consent the access: " + err.message);
      this.setState({
        showLoginBtn: true
      });
    }
  }

  async loginBtnClick() {
    try {
      // Popup login page to get Microsoft Graph access token
      await MODS.popupLoginPage();
    }
    catch (err) {
      alert("Login failed: " + err);
      return;
    }

    await this.getData();
  }

  async onAddItem() {
    await this.storageClient.addItem(this.state.newItemContent);
    this.setState({
      newItemContent: ""
    });
    this.refresh();
  }

  async onDeleteItem(id) {
    await this.storageClient.deleteItem(id);
    this.refresh();
  }

  async onCompletionStatusChange(id, index, isComplete) {
    this.handleInputChange(index, "isComplete", isComplete);
    await this.storageClient.updateItemCompltionStatus(id, isComplete);
  }

  handleInputChange(index, property, value) {
    const tmp = JSON.parse(JSON.stringify(this.state.items))
    tmp[index][property] = value;
    this.setState({
      items: tmp
    })
  }

  async refresh() {
    this.setState({
      items: await this.storageClient.getItems(),
    });
  }

  render() {
    const items = this.state.items?.map((item, index) =>
      <div key={item.id} className="item">
        <span className="number">
          <input
            type="checkbox"
            checked={this.state.items[index].isComplete}
            onChange={(ev) => this.onCompletionStatusChange(item.id, index, ev.target.checked)}
            className="isComplete"
          />
          {index + 1}
        </span>
        <span className="content">
          <input
            type="text"
            value={this.state.items[index].Title}
            onChange={(ev) => this.handleInputChange(index, "content", ev.target.value)}
            className="text"
          />

          <button onClick={() => this.onDeleteItem(item.id)}>Delete</button>
        </span>
      </div>
    );

    return (
      <div>
        <div className="profile">
          <h2>To Do List</h2>
          <p><b>Name:</b> {this.state.userInfo.userName}</p>
          <p><b>E-mail:</b> {this.state.userInfo.preferredUserName}</p>
          {this.state.photoDescription && <p><img src={this.state.photoObjectURL} alt={this.state.photoDescription} /></p>}
          {this.state.showLoginBtn && <button onClick={() => this.loginBtnClick()}>Auth and show</button>}
        </div>

        {!this.state.showLoginBtn && <div className="todo">
          <div className="add">
            <input
              type="text"
              value={this.state.newItemContent}
              onChange={(ev) => this.setState({ newItemContent: ev.target.value })}
            />
            <button onClick={() => this.onAddItem()}>Add</button>
          </div>
          <div className="item">
            <span className="number">Number</span>
            <span className="content">Item </span>
          </div>
          {items}
        </div>}
      </div>
    );
  }
}
export default Tab;
