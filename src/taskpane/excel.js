/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, document, Office, require */

const graphHelper = require("./../helpers/graphHelper");

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(function() {
      $("#getGraphDataButton").click(graphHelper.getGraphData);
    });
  }
});
