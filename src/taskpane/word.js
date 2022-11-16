/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { getUserProfile } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("getProfileButton").onclick = run;
  }
});

export async function run() {
  getUserProfile(writeDataToOfficeDocument);
}

function writeDataToOfficeDocument(result) {
  return Word.run(function (context) {
    let data = [];
    let userProfileInfo = filterUserProfileInfo(result);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        data.push(userProfileInfo[i]);
      }
    }

    const documentBody = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
