/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dialogFallback } from "./fallbackauthdialog.js";
import { getUserData } from "./middle-tier-calls";
import { showMessage } from "./message-helper";
import { handleClientSideErrors } from "./error-handler";

/* global Office */

let retryGetMiddletierToken = 0;

export async function getUserProfile(callback) {
  try {
    let middletierToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    let response = await getUserData(middletierToken);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaMiddletierToken = await Office.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = getUserData(mfaMiddletierToken);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

function handleAADErrors(response, callback) {
  // On rare occasions the middle tier token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired middle tier token.

  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken <= 0) {
    retryGetMiddletierToken++;
    getUserProfile(callback);
  } else {
    dialogFallback(callback);
  }
}
