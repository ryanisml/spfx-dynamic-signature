import { SPHttpClient } from '@microsoft/sp-http';

export interface IPdfESignatureProps {
  description: string;
  spHttpClient: SPHttpClient;
  currentWebUrl: string;
}
