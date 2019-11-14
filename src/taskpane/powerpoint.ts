/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, document, Office */

import { getGraphData } from "./../helpers/graphHelper";

Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
        $(document).ready(function () {
            $("#getGraphDataButton").click(getGraphData);
        });
    }
});

export function writeDataToOfficeDocument(result: string[]) {
    let data: string = "";
    for (let i = 0; i < result.length; i++) {
        if (result[i] !== null) {
            data += result[i] + "\n";
        }
    }

    Office.context.document.setSelectedDataAsync(data, function (
        asyncResult
    ) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
        }
    });
}
