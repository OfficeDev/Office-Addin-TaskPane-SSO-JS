/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, document, Excel, Office */

import { getGraphData } from "../helpers/graphHelper";

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(function() {
      $("#getGraphDataButton").click(getGraphData);
    });
  }
});

export function writeDataToOfficeDocument(result) {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    let data = [];
    let i;
    for (i = 0; i < result.length; i++) {
      var innerArray = [];
      innerArray.push(result[i]);
      data.push(innerArray);
    }

    const rangeAddress = `B5:B${5 + (result.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
