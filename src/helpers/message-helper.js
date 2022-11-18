/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document */

export function showMessage(text) {
  document.getElementById("message-area").style.display = "flex";
  document.getElementById("message-area").innerText = text;
}

export function clearMessage() {
  document.getElementById("message-area").style.display = "flex";
  document.getElementById("message-area").innerText = "---<br>";
}

export function hideMessage() {
  document.getElementById("message-area").style.display = "none";
  document.getElementById("message-area").innerText = "---<br>";
}
