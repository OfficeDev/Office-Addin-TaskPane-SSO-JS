/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Excel, Office */

import { getUserData } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("getProfileButton").onclick = run;
  }
});

export async function run() {
  getUserData(writeDataToOfficeDocument);
}

function writeDataToOfficeDocument(result) {
  return Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data = [];
    let userProfileInfo = filterUserProfileInfo(result);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        let innerArray = [];
        innerArray.push(userProfileInfo[i]);
        data.push(innerArray);
      }
    }
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
