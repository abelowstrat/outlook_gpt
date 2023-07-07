import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { Configuration, OpenAIApi } from "openai";
/* import Office
/* global require */
/* global  Office */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  generatedText: string;
  startText: string;
  finalMailText: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props) {
    super(props);
    this.state = {
      generatedText: "",
      startText: "",
      finalMailText: "",
    };
  }

  generateText = async () => {
    const configuration = new Configuration({
      apiKey: "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",
    });
    const openai = new OpenAIApi(configuration);
    try {
      const response = await openai.createCompletion({
        model: "text-davinci-003",
        prompt: "Turn the following text into a professional business mail: " + this.state.startText,
        temperature: 0.7,
        max_tokens: 300,
      });

      console.log(response); // Log the response object
      console.log(response.data); // Log the response data
  
      this.setState({ generatedText: response.data.choices[0].text });
    } catch (error) {
      console.error(error); // Log any error that occurs during the API call
    }
  };

  insertIntoMail = () => {
    const finalText = this.state.finalMailText.length === 0 ? this.state.generatedText : this.state.finalMailText;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Html,
    });
  };

  render() {
    return (
      <div>
        <main>
          <h2> Open AI business e-mail generator </h2>
          <p> Briefly describe what you want to communicate in the mail:</p>
          <textarea onChange={(e) => this.setState({ startText: e.target.value })} rows={10} cols={40} />
          <p>
            <DefaultButton onClick={this.generateText}>Generate text</DefaultButton>
          </p>
          <textarea
            defaultValue={this.state.generatedText}
            onChange={(e) => this.setState({ finalMailText: e.target.value })}
            rows={10}
            cols={40}
          />
          <p>
            <DefaultButton onClick={this.insertIntoMail}>Insert into mail</DefaultButton>
          </p>
        </main>
      </div>
    );
  }
}
