import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IMaycoSafetyProps {
  description: string;
  userDisplayName: string;
  spHttpClient:any;
  spContext:any;
  context:WebPartContext
}
