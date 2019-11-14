/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, document, Office, Word */

import { getGraphData } from "./../helpers/graphHelper";

Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
        $(document).ready(function () {
            $("#getGraphDataButton").click(getGraphData);
        });
    }
});

export function writeDataToOfficeDocument(result: string[]) {
    return Word.run(function (context) {
        const documentBody: Word.Body = context.document.body;
        for (let i = 0; i < result.length; i++) {
            if (result[i] !== null) {
            documentBody.insertParagraph(result[i], "End");
            }
        }

        return context.sync();
    });
}
