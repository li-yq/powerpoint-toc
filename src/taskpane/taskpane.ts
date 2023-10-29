/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { base64Image } from "../base64Image";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("load-config").onclick = () => clearMessage(loadConfig);
    document.getElementById("save-config").onclick = () => clearMessage(saveConfig);
    // TODO4: Assign event handler for insert-text button.
    // TODO6: Assign event handler for get-slide-metadata button.
    // TODO8: Assign event handlers for add-slides and the four navigation buttons.
  }
});

function loadConfig() {
  PowerPoint.run(async function (context) {
    const slides = context.presentation.getSelectedSlides()
    slides.load("tags/key, tags/value");
    slides.load()
    await context.sync()
    
    if (slides.items.length == 0) {
      setMessage("no slide selected")
      return
    }

    var config = slides.getItemAt(0).tags.getItemOrNullObject("KEY")
    config.load("value")
    await context.sync()
    let value= config.isNullObject ? "{}":config.value

    setMessage(value)

  })
}

function saveConfig(){
  PowerPoint.run(async (context: PowerPoint.RequestContext) => {
    const slides = context.presentation.getSelectedSlides()
    slides.load("items")
    await context.sync()

    if (slides.items.length == 0) {
      setMessage("no slide selected")
      return
    }

    slides.getItemAt(0).tags.add("KEY", JSON.stringify({
      key1: "val1"
    }))

  })



}

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
  document.getElementById("message").innerText = message;
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    document.getElementById("message").innerText = "";
    await callback();
  } catch (error) {
    setMessage("Error: " + error.toString());
  }
}