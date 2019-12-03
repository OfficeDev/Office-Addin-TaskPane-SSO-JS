/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global OfficeRuntime, require */

const documentHelper = require("./documentHelper");
const errorHandler = require("./../../node_modules/office-addin-sso/lib/error-handler");
const fallbackAuthHelper = require("./fallbackAuthHelper");
const msGraphHelper = require("./../../node_modules/office-addin-sso/lib/msgraph-helper");
const messageHelper = require("./../../node_modules/office-addin-sso/lib/message-helper");
let retryGetAccessToken = 0;

export async function getGraphData() {
  try {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      forMSGraphAccess: true
    });
    let exchangeResponse = await msGraphHelper.MSGraphHelper.getGraphToken(
      bootstrapToken
    );
    if (exchangeResponse.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({
        authChallenge: exchangeResponse.claims
      });
      exchangeResponse = msGraphHelper.MSGraphHelper.getGraphToken(
        mfaBootstrapToken
      );
    }

    if (exchangeResponse.error) {
      // AAD errors are returned to the client with HTTP code 200, so they do not trigger
      // the catch block below.
      handleAADErrors(exchangeResponse);
    } else {
      // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
      // in the .fail callback of that call
      const response = await msGraphHelper.MSGraphHelper.makeGraphApiCall(
        exchangeResponse.access_token
      );
      documentHelper.writeDataToOfficeDocument(response);
      messageHelper.showMessage("Your data has been added to the document.");
    }
  } catch (exception) {
    if (exception.code) {
      if (errorHandler.handleClientSideErrors(exception)) {
        fallbackAuthHelper.dialogFallback();
      }
    } else {
      messageHelper.showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
  }
}

function handleAADErrors(exchangeResponse) {
  // On rare occasions the bootstrap token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired bootstrap token.
  if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1) && (retryGetAccessToken <= 0))  {
    retryGetAccessToken++;
    getGraphData();
  } else {
    fallbackAuthHelper.dialogFallback();
  }
}
