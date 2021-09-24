/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-64.png";
import "../../assets/icon-80.png";

/* global document, Office, require */

const ssoAuthHelper = require("./../helpers/ssoauthhelper");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getGraphDataButton").onclick = ssoAuthHelper.getGraphData;
  }
});
