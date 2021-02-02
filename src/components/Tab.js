// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Loader } from '@fluentui/react-northstar';
import { withMsal } from '@azure/msal-react';

import './App.css';

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component {
  constructor(props) {
    super(props)

    this.state = {
      context: {},
      ssoToken: "",
      consentRequired: false,
      consentProvided: false,
      graphAccessToken: "",
      profile: {},
      error: false
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  componentDidMount() {
    //Initialize the Microsoft Teams SDK
    microsoftTeams.initialize();

    //Get the user context from Teams and set it in the state
    microsoftTeams.getContext((context, error) => {
      this.setState({context:context});
    });

    //Perform single sign-on authentication
    let authTokenRequestOptions = {
      successCallback: (result) => { this.ssoLoginSuccess(result) }, //The result variable is the SSO token.
      failureCallback: (error) => {this.ssoLoginFailure(error)}
    };

    microsoftTeams.authentication.getAuthToken(authTokenRequestOptions);
  }

  ssoLoginSuccess = async (result) => {
    this.setState({ssoToken:result});
    this.exchangeClientTokenForServerToken(result);
  }

  ssoLoginFailure = (error) => {
    console.error("SSO failed: ", error);
    this.setState({error:true});
  }

  //Exchange the SSO access token for a Graph access token
  //Learn more: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow
  exchangeClientTokenForServerToken = async (token) => {

    let serverURL = `${process.env.REACT_APP_BASE_URL}/getGraphAccessToken?ssoToken=${token}`;
    let response = await fetch(serverURL).catch(this.unhandledFetchError); //This calls getGraphAccessToken route in /api-server/app.js
    let data = await response.json().catch(this.unhandledFetchError);

    if (data.subError === 'consent_required') {
      //A consent_required error means it's the first time a user is logging into to the app, so they must consent to sharing their Graph data with the app.
      //They may also see this error if MFA is required.
      this.setState({consentRequired:true}); //This displays the consent required message.
      
      //Show a popup dialogue prompting the user to consent to the required API permissions.
      //Learn more: https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-tab-aad#initiate-authentication-flow
      this.props.msalContext.instance
        .acquireTokenPopup({
          scopes: ["user.read"],
        })
        .then(response => {
          console.log(response);
          this.consentSuccess(response.accessToken);
        })
        .catch((error) => {
          console.log(error);
          this.consentFailure(error);
        });
      
    } else if (!response.ok) {
      //Unknown error
      console.error(data);
      this.setState({error:true});
    } else {
      //Server side token exchange worked. Save the access_token to state, so that it can be picked up and used by the componentDidMount lifecycle method.
      this.setState({graphAccessToken:data['accessToken']});
    }
  }

  //React lifecycle method that gets called after a component's state or props updates
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidupdate
  componentDidUpdate = async (prevProps, prevState) => {
    
    //Check to see if a Graph access token is now in state AND that it didn't exist previously
    if ((prevState.graphAccessToken === "") && (this.state.graphAccessToken !== "")) {
      this.callGraphFromClient();
    }
  }  

  //Callback function for a successful authorization
  consentSuccess = (result) => {
    //Save the Graph access token in state
    this.setState({
      graphAccessToken: result,
      consentProvided: true
    });
  }

  consentFailure = (reason) => {
    console.error("Consent failed: ", reason);
    this.setState({error:true});
  }

  // Fetch the user's profile photo from Graph using the access token retrieved either from the server 
  // or microsoftTeams.authentication.authenticate
  callGraphFromClient = async () => {
    let upn = this.state.context['upn'];

    let graphPhotoEndpoint = `https://graph.microsoft.com/v1.0/users/${upn}`;

    let graphRequestParams = {
      method: 'GET',
      headers: {
        "authorization": `bearer ${this.state.graphAccessToken}`
      }
    }

    let response = await fetch(graphPhotoEndpoint, graphRequestParams).catch(this.unhandledFetchError);

    if(!response.ok) {
      console.error("ERROR: ", response);
      this.setState({error:true});
    }
    
    let profileData = await response.json().catch(this.unhandledFetchError); 
    console.log(profileData);

    this.setState({profile: profileData})
  }

  //Generic error handler ( avoids having to do async fetch in try/catch block )
  unhandledFetchError = (err) => {
    console.error("Unhandled fetch error: ", err);
    this.setState({error:true});
  }

  render() {

      let title = Object.keys(this.state.context).length > 0 ?
        'Congratulations ' + this.state.context['upn'] + '! This is your tab' : <Loader/>;

      let ssoMessage = this.state.ssoToken === "" ?
        <Loader label='Performing Azure AD single sign-on authentication...'/>: null;
      
      let serverExchangeMessage = (this.state.ssoToken !== "") && (!this.state.consentRequired) && (this.state.photo==="") ?
        <Loader label='Exchanging SSO access token for Graph access token...'/> : null;

      let consentMessage = (this.state.consentRequired && !this.state.consentProvided) ?
        <Loader label='Consent required.'/> : null;

      let profile = this.state.profile;

      let content;

      if (this.state.error) {
        content = <h1>ERROR: Please ensure pop-ups are allowed for this website and retry</h1>
      } else {
        content =
          <div>
            <h1>{title}</h1>
            <h3>{ssoMessage}</h3>
            <h3>{serverExchangeMessage}</h3>          
            <h3>{consentMessage}</h3>
            <pre>{JSON.stringify(profile)}</pre>
          </div>
      }
      
      return (
        <div>
          {content}
        </div>
      );
  }
}
export default Tab = withMsal(Tab);