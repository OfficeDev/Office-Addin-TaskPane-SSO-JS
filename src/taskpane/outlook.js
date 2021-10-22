/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
/* eslint-disable no-unused-vars */
import icon16 from "../../assets/icon-16.png";
import icon32 from "../../assets/icon-32.png";
import icon64 from "../../assets/icon-64.png";
import icon80 from "../../assets/icon-80.png";
import icon128 from "../../assets/icon-128.png";
/* eslint-enable no-unused-vars */

/* global document, Office, require */

const ssoAuthHelper = require("./../helpers/ssoauthhelper");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getGraphDataButton").onclick = ssoAuthHelper.getGraphData;
  }
});
