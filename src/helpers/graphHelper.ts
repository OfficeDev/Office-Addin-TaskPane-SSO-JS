
import { handleAADErrors, handleClientSideErrors } from './errorHandler'
import { showMessage } from './messageHelper';
import { writeFileNamesToOfficeDocument } from './documentHelper';

export async function getGraphData(): Promise<void> {
    try {
        let bootstrapToken: string = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true });
        let exchangeResponse: any = await getGraphToken(bootstrapToken);
        if (exchangeResponse.claims) {
            // Microsoft Graph requires an additional form of authentication. Have the Office host 
            // get a new token using the Claims string, which tells AAD to prompt the user for all 
            // required forms of authentication.
            let mfaBootstrapToken: string = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
            exchangeResponse = await getGraphToken(mfaBootstrapToken);
        }

        if (exchangeResponse.error) {
            // AAD errors are returned to the client with HTTP code 200, so they do not trigger
            // the catch block below.
            handleAADErrors(exchangeResponse);
        }
        else {
            // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
            // in the .fail callback of that call, not in the catch block below.
            makeGraphApiCall(exchangeResponse.access_token);
        }
    }
    catch (exception) {
        // The only exceptions caught here are exceptions in your code in the try block
        // and errors returned from the call of `getAccessToken` above.
        if (exception.code) {
            handleClientSideErrors(exception);
        }
        else {
            showMessage("EXCEPTION: " + JSON.stringify(exception));
        }
    }
}

export function makeGraphApiCall(accessToken: string): void {
    $.ajax({
        type: "GET",
        url: "/getuserdata",
        headers: { "access_token": accessToken },
        cache: false
    }).done(function (response) {

        writeFileNamesToOfficeDocument(response)
            .then(function () {
                showMessage("Your data has been added to the document.");
            })
            .catch(function (error) {
                // The error from writeFileNamesToOfficeDocument will begin 
                // "Unable to add filenames to document."
                showMessage(error);
            });
    })
        .fail(function (errorResult) {
            // This error is relayed from `app.get('/getuserdata` in app.js file.
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
}

export async function getGraphToken(bootstrapToken): Promise<any> {
    let response = await $.ajax({
        type: "GET",
        url: "/auth",
        headers: { "Authorization": "Bearer " + bootstrapToken },
        cache: false
    });
    return response;
}