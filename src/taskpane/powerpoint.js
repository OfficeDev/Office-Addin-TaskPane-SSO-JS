/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getUserProfile } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("getProfileButton").onclick = run;
  }
});

export async function run() {
  getUserProfile(writeDataToOfficeDocument);
}

function writeDataToOfficeDocument(result) {
  let data = [];
  let userProfileInfo = filterUserProfileInfo(result);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
