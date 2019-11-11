import { getGraphData } from './graphHelper';
import { dialogFallback } from './fallbackAuthHelper';
import { showMessage } from './messageHelper';

export function handleClientSideErrors(error: any) {
    switch (error.code) {

        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one 
            // is logged into Office, then the first call of getAccessToken should pass the 
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see this error
            showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.");
            break;
        case 13006:
            // Only seen in Office on the Web.
            showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again.");
            break;
        case 13008:
            // Only seen in Office on the Web.
            showMessage("Office is still working on the last operation. When it completes, try this operation again.");
            break;
        case 13010:
            // Only seen in Office on the Web.
            showMessage("Follow the instructions to change your browser's zone configuration.");
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in.
            dialogFallback();
            break;
    }
}

export function handleAADErrors(exchangeResponse: any) {
    // On rare occasions the bootstrap token is unexpired when Office validates it,
    // but expires by the time it is sent to AAD for exchange. AAD will respond
    // with "The provided value for the 'assertion' is not valid. The assertion has expired."
    // Retry the call of getAccessToken (no more than once). This time Office will return a 
    // new unexpired bootstrap token.

    let retryGetAccessToken = 0;
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) {
        retryGetAccessToken++;
        getGraphData();
    }
    else {
        dialogFallback();
    }
}