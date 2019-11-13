/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, document, Office */

import { getGraphData } from "./../helpers/graphHelper";

Office.onReady(function() {
  $(document).ready(function() {
    $("#getGraphDataButton").click(getGraphData);
  });
});
