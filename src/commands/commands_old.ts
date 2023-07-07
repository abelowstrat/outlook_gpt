/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;

import { Configuration, OpenAIApi } from "openai";
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function getSelectedText(): Promise<any> {
  return new Office.Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
        const configuration = new Configuration({
          apiKey: "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",
        });
        const openai = new OpenAIApi(configuration);
        const response = await openai.createCompletion({
          model: "gpt-3.5-turbo-instruct",
          prompt: "Turn the following text into a professional business mail: " + asyncResult.value,
          temperature: 0.7,
          max_tokens: 300,
        });

        resolve(response.data.choices[0].text);
      });
    } catch (error) {
      reject(error);
    }
  });
}

function action(event: Office.AddinCommands.Event) {
  getSelectedText().then(function (selectedText) {
    Office.context.mailbox.item.setSelectedDataAsync(selectedText, { coercionType: Office.CoercionType.Text });
    event.completed();
  });
}
