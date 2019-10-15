import * as React from "react";
import styles from "./HelloWorld.module.scss";
import { IHelloWorldProps } from "./IHelloWorldProps";
import * as strings from "HelloWorldWebPartStrings";
import { escape } from "@microsoft/sp-lodash-subset";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { EmailProperties, sp } from "@pnp/sp";
import {
  SPHttpClient,
  ISPHttpClientOptions,
  SPHttpClientResponse,
  HttpClient,
  HttpClientResponse,
  AadTokenProvider
} from "@microsoft/sp-http";
import { IHelloWorldState } from "./IHelloWorldState";

import {
  autobind,
  PrimaryButton,
  TextField,
  Label,
  CheckboxVisibility,
  SelectionMode
} from "office-ui-fabric-react";
import * as $ from "jquery";

import { AadHttpClient, IHttpClientOptions } from "@microsoft/sp-http";

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  IHelloWorldState
> {
  private error: string = null;
  public result: string = null;
  private token: string = null;
  private decodedToken: any = null;

  constructor(props) {
    super(props);

    this.handleSubmit = this.handleSubmit.bind(this);
  }

  private handleSubmit() {
    event.preventDefault();

    console.log("laura.rawlings@rbaconsulting.com");
    var emailList = [];
    emailList.push("laura.rawlings@rbaconsulting.com");

    //this.GetSimpleToken();
    this.SendEmailWithAadHttpClient();
  }

  //need async for await
  private async SendEmailWithAadHttpClient() {
    console.log("Starting SendEmailWithAadHttpClient...");
    if (this.props.context) {
      if (this.props.context.aadHttpClientFactory) {
        this.props.context.aadHttpClientFactory
          .getClient(this.props.ClientID)
          .then((client: AadHttpClient): void => {
            const header: Headers = new Headers();

            //when using the AADHttpClient class, it automatically sends the token in a Authorization Header
            /*
          •	AAD Integration for 3rd Party APIs
              o	Option #1 – Work with the authorization token directly – AadTokenProvider class 
              o	Option #2 – Use a client that automatically adds the token as an authorization header for you – AadHttpClient class (tutorial)
              https://urldefense.proofpoint.com/v2/url?u=https-3A__docs.microsoft.com_en-2Dus_sharepoint_dev_spfx_use-2Daadhttpclient&d=DwMFAg&c=h2Fu29X6B8vuaRkwP4jdQg&r=SsXd-HFWmbEdYTTgErh0FciqP9FpQUR3OPdFUBEbeFj91yT7l1ZVNzVdXniIOyKN&m=tvjEDChPC6g5BMRME83rq2hk2F43MTO31ggZwqlyF1U&s=DScBUC432dx14So_XIsPZyuJLITjbLSkQ7Uq2Wy78cg&e=
          */
            header.append("Content-Type", "application/x-www-form-urlencoded");
            header.append("Access-Control-Allow-Origin", "*");
            header.append("Access-Control-Allow-Methods", "POST");
            header.append(
              "Access-Control-Allow-Headers",
              "Origin, Content-Type, X-Auth-Token"
            );

            var emailToList = [];
            emailToList.push("laura.rawlings@rbaconsulting.com");
            var emailCCList = [];
            emailCCList.push("laura_rawlings@crgl-thirdparty.com");

            var data = JSON.stringify({
              to: emailToList,
              cc: emailCCList,
              from: "donotreply@cargill.com",
              subject:
                "TESTING ONLY:  Test for Global Travel Email using the Integration Team email API",
              body:
                "PGgyPlRyYXZlbCBSZXF1ZXN0PC9oMj48dGFibGUgY2xhc3M9ZW1haWxUYWJsZV81ZTlhZDAwYj4gPHRyPiAgIDx0ZCBzdHlsZT0iIHdpZHRoOiAxNTBweDsgYmFja2dyb3VuZC1jb2xvcjogI2VhZWFlYTtmb250LXdlaWdodDogNTAwOyI+TmFtZSAgIDwvdGQ+ICAgPHRkPjxkaXY+TGF1cmEgUmF3bGluZ3M8L2Rpdj4gICA8L3RkPiA8L3RyPiA8dHI+ICAgPHRkIHN0eWxlPSIgd2lkdGg6IDE1MHB4OyBiYWNrZ3JvdW5kLWNvbG9yOiAjZWFlYWVhO2ZvbnQtd2VpZ2h0OiA1MDA7Ij5EU0lEICAgPC90ZD4gICA8dGQ+PGRpdj4gMTIzPC9kaXY+ICAgPC90ZD4gPC90cj4gPHRyPiAgIDx0ZCBzdHlsZT0iIHdpZHRoOiAxNTBweDsgYmFja2dyb3VuZC1jb2xvcjogI2VhZWFlYTtmb250LXdlaWdodDogNTAwOyI+VHJhdmVsZXIgRW1haWwgICA8L3RkPiAgIDx0ZD48ZGl2PiBsYXVyYV9yYXdsaW5nc0BjcmdsLXRoaXJkcGFydHkuY29tPC9kaXY+ICAgPC90ZD4gPC90cj4gPHRyPiAgIDx0ZCBzdHlsZT0iIHdpZHRoOiAxNTBweDsgYmFja2dyb3VuZC1jb2xvcjogI2VhZWFlYTtmb250LXdlaWdodDogNTAwOyI+VHJhdmVsZXIgUGhvbmUgICA8L3RkPiAgIDx0ZD48ZGl2PiA2MTItNTU1LTEyMTI8L2Rpdj4gICA8L3RkPiA8L3RyPiA8dHI+ICAgPHRkIHN0eWxlPSIgd2lkdGg6IDE1MHB4OyBiYWNrZ3JvdW5kLWNvbG9yOiAjZWFlYWVhO2ZvbnQtd2VpZ2h0OiA1MDA7Ij5Db3VudHJ5ICAgPC90ZD4gICA8dGQ+PGRpdj5Vbml0ZWQgU3RhdGVzPC9kaXY+ICAgPC90ZD4gPC90cj4gPHRyPiAgIDx0ZCBzdHlsZT0iIHdpZHRoOiAxNTBweDsgYmFja2dyb3VuZC1jb2xvcjogI2VhZWFlYTtmb250LXdlaWdodDogNTAwOyI+UGxhdGZvcm0gLSBCdXNpbmVzcyBVbml0ICAgPC90ZD4gICA8dGQ+PGRpdj4oMDIxKSBBZ3JpY3VsdHVyYWwgU3VwcGx5IENoYWluIC0gKDc2MDEpIENBU0MgQVBBQyBCT1NDPC9kaXY+ICAgPC90ZD4gPC90cj4gPHRyPiAgIDx0ZCBzdHlsZT0iIHdpZHRoOiAxNTBweDsgYmFja2dyb3VuZC1jb2xvcjogI2VhZWFlYTtmb250LXdlaWdodDogNTAwOyI+UHVycG9zZSBvZiBUcmlwICAgPC90ZD4gICA8dGQ+PGRpdj5FeHRlcm5hbCB2aXNpdCAtIFByb3NwZWN0cyBjdXN0b21lci9zdXBwbGllcjwvZGl2PiAgIDwvdGQ+IDwvdHI+IDx0cj4gICA8dGQgc3R5bGU9IiB3aWR0aDogMTUwcHg7IGJhY2tncm91bmQtY29sb3I6ICNlYWVhZWE7Zm9udC13ZWlnaHQ6IDUwMDsiPlByb2ZpbGUgc3RhdHVzICAgPC90ZD4gICA8dGQ+PGRpdj5OZXcgdHJhdmVsZXIgb3IgdHJhdmVsZXIgcHJvZmlsZSBkb2VzIG5vdCBleGlzdC48L2Rpdj4gICA8L3RkPiA8L3RyPiA8dHI+ICAgPHRkIHN0eWxlPSIgd2lkdGg6IDE1MHB4OyBiYWNrZ3JvdW5kLWNvbG9yOiAjZWFlYWVhO2ZvbnQtd2VpZ2h0OiA1MDA7Ij5Db21tZW50cyAgIDwvdGQ+ICAgPHRkPjxkaXY+IDwvZGl2PiAgIDwvdGQ+PC90YWJsZT48aDQ+UmFpbDwvaDQ+PHRhYmxlIGNsYXNzPWVtYWlsVGFibGVfNWU5YWQwMGI+IDx0cj4gICA8dGQgc3R5bGU9IiB3aWR0aDogMTUwcHg7IGJhY2tncm91bmQtY29sb3I6ICNlYWVhZWE7Zm9udC13ZWlnaHQ6IDUwMDsiPkRlcGFydCBDaXR5ICAgPC90ZD4gICA8dGQgc3R5bGU9IndpZHRoOiAzMDBweDsgYmFja2dyb3VuZC1jb2xvcjogI2ZmZmZmZjsiPjxkaXY+IG1zcDwvZGl2PiAgIDwvdGQ+ICAgPHRkIHN0eWxlPSIgd2lkdGg6IDE1MHB4OyBiYWNrZ3JvdW5kLWNvbG9yOiAjZWFlYWVhO2ZvbnQtd2VpZ2h0OiA1MDA7Ij5BcnJpdmFsIENpdHkgICA8L3RkPiAgIDx0ZCBzdHlsZT0id2lkdGg6IDMwMHB4OyBiYWNrZ3JvdW5kLWNvbG9yOiAjZmZmZmZmOyI+PGRpdj4gY2hpPC9kaXY+ICAgPC90ZD4gPC90cj4gPHRyPiAgIDx0ZCBzdHlsZT0iIHdpZHRoOiAxNTBweDsgYmFja2dyb3VuZC1jb2xvcjogI2VhZWFlYTtmb250LXdlaWdodDogNTAwOyI+RGVwYXJ0IERhdGUgICA8L3RkPiAgIDx0ZCBzdHlsZT0id2lkdGg6IDMwMHB4OyBiYWNrZ3JvdW5kLWNvbG9yOiAjZmZmZmZmOyI+PGRpdj5EZWNlbWJlciAyOHRoIDIwMTggPGJyLz5Bbnl0aW1lPC9kaXY+ICAgPC90ZD4gICA8dGQ+ICAgPC90ZD4gICA8dGQ+ICAgPC90ZD4gPC90cj48L3RhYmxlPiA8aHI+IDxocj48YnIvPjxici8+PGJyLz48YnIvPg=="
            });
          
            const httpClientOptions: IHttpClientOptions = {
              body: data,
              headers: header,
              method: "POST"
            };

            client
              .post(
                this.props.APIUrl,
                AadHttpClient.configurations.v1,
                httpClientOptions
              )
              .then(
                (response: HttpClientResponse): Promise<HttpClientResponse> => {
                  console.log("REST API response received.");
                  //this.result = response.json();
                  return response.json();
                }
              );
          });
      } else {
        console.log("this.props.context.aadHttpClientFactory is false");
      }
    } else {
      console.log("this.props.context is false");
    }
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div>
        <div>
          <button onClick={this.handleSubmit}>Send email!</button>
        </div>
        <div className={styles.error}>{this.error}</div>
        <div className={styles.result}>{this.result}</div>
      </div>
    );
  }
}
