import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHelloWorldProps {
  context: WebPartContext;
  ClientID: string;
  APIUrl: string;
  ToEmail: string;
}
