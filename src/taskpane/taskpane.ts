/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. 
 *
 * This file shows how to use the SSO API to get a bootstrap token.
 */
import { getGraphData } from './../helpers/graphHelper';

Office.onReady(function(info) {
    $(document).ready(function() {
        $('#getGraphDataButton').click(getGraphData);
    });
});
