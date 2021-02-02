// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";
import { BrowserRouter as Router, Route } from "react-router-dom";

import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";

import Privacy from "./Privacy";
import TermsOfUse from "./TermsOfUse";
import Tab from "./Tab";

const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_AZURE_APP_REGISTRATION_ID,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: `${process.env.REACT_APP_BASE_URL}/tab`,
  }
}

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
function App() {

  // Initialize the Microsoft Teams SDK
  microsoftTeams.initialize();

  const msalInstance = new PublicClientApplication(msalConfig);

  // Display the app home page hosted in Teams
  return (
    <MsalProvider instance={msalInstance}>
      <Router>
        <Route exact path="/privacy" component={Privacy} />
        <Route exact path="/termsofuse" component={TermsOfUse} />
        <Route exact path="/tab" component={Tab} />
      </Router>
    </MsalProvider>
  );
}

export default App;
