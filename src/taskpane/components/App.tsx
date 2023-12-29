import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { Configuration, OpenAIApi } from "openai";
import config from "./config";
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
      apiKey: config.apiKey,
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
      <div style={{ padding: "20px", backgroundColor: "#333", color: "#fff" }}>
        <main style={{ maxWidth: "600px", margin: "0 auto", backgroundColor: "#333", paddingBottom: "200px" }}>
          <h2 style={{ textAlign: "center", marginBottom: "20px" }}> Open AI business e-mail generator </h2>
          <p style={{ marginBottom: "10px" }}> Briefly describe what you want to communicate in the mail:</p>
          <textarea onChange={(e) => this.setState({ startText: e.target.value })} 
            rows={10}
            cols={40}
            style={{ marginBottom: "5px", width: "95%", padding: "10px" }} 
          />
          <p>
            <DefaultButton onClick={this.generateText} style={{ marginBottom: "15px" }}>Generate text</DefaultButton>
          </p>
          <textarea
            defaultValue={this.state.generatedText}
            onChange={(e) => this.setState({ finalMailText: e.target.value })}
            rows={10}
            cols={40}
            style={{ marginBottom: "5px", width: "95%", padding: "10px" }}
          />
          <p>
            <DefaultButton onClick={this.insertIntoMail} style={{ marginBottom: "15px" }}>Insert into mail</DefaultButton>
          </p>
        </main>
      </div>
    );
  }
}
