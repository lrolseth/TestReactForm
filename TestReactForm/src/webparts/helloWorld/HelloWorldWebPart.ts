import * as React from "react";
import * as ReactDom from "react-dom";
import styles from "./components/HelloWorld.module.scss";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "HelloWorldWebPartStrings";
import { HttpClient } from "@microsoft/sp-http";
import HelloWorld from "./components/HelloWorld";
import { IHelloWorldProps } from "./components/IHelloWorldProps";

export interface IHelloWorldWebPartProps {
  ClientID: string;
  APIUrl: string;
  ToEmail: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<
  IHelloWorldWebPartProps
> {
  private error: string = null;
  private result: string = null;
  private token: string = null;
  private decodedToken: any = null;

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        context: this.context,
        ToEmail: this.properties.ToEmail,
        APIUrl: this.properties.APIUrl,
        ClientID: this.properties.ClientID
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /*
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.spfxAad}">
        <div class="${styles.container}">
          <h2>Result</h2>
          ${
            this.error
              ? `<div class="${styles.error}">${this.error}</div>`
              : `<div class="${styles.result}">${this.result}</div>`
          }
          ${
            this.decodedToken
              ? `<h2>Decoded JWT</h2>
                <table "${styles.decodedToken}">
                  <tr><td><b>Claim</b></td><td><b>Value</b></td></tr>
                  ${Object.keys(this.decodedToken)
                    .map(
                      k =>
                        `<tr><td>${k}</td><td>${this.decodedToken[k]}</td></tr>`
                    )
                    .join("")}
               </table>`
              : ""
          }
          ${
            this.token
              ? `<h2>Raw JWT</h2>
               <div class="${styles.token}">${this.token}</div>`
              : ""
          }
        </div>
      </div>`;
  }
  */

  public async onInit(): Promise<void> {
    if (
      !this.properties.ClientID ||
      this.properties.ClientID === "00000000-0000-0000-0000-000000000000"
    ) {
      this.error = "Please set Client ID property in the webpart settings";
      return;
    }

    this.GetSimpleToken();
  }

  private async GetSimpleToken() {
    console.log("Starting GetSimpleToken...");
    try {
      // Get token PROVIDER
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      // Get TOKEN
      this.token = await tokenProvider.getToken(this.properties.ClientID);
      console.log("Token");
      console.log(this.token);
    } catch (error) {
      this.error = error;
      console.log("Error getting token");
      console.log(error);
    }

    if (this.token) {
      try {
        // Call secured API with token
        let headers = new Headers();
        headers.set("Authorization", `Bearer ${this.token}`);
        headers.set("Access-Control-Allow-Origin", "*");
        headers.set("Content-Type", "application/x-www-form-urlencoded");
        headers.set("Access-Control-Allow-Methods", "POST");
        headers.set(
          "Access-Control-Allow-Headers",
          "Origin, Content-Type, X-Auth-Token"
        );

        const response = await this.context.httpClient.get(
          this.properties.APIUrl,
          HttpClient.configurations.v1,
          { headers }
        );
        console.log("Response");
        console.log(response);

        if (response.ok) {
          const result = await response.json();
          this.result = result.message;
          console.log(this.result);
        } else {
          throw new Error(`${response.status}: ${response.statusText}`);
        }
      } catch (error) {
        console.log(
          `Error calling API URL: ${this.properties.APIUrl}. ${error}`
        );
      }
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("ClientID", {
                  label: "ClientID"
                }),
                PropertyPaneTextField("APIUrl", {
                  label: "APIUrl"
                }),
                PropertyPaneTextField("ToEmail", {
                  label: "ToEmail"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
