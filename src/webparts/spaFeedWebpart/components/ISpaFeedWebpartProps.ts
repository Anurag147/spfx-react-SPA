import { SPHttpClient } from '@microsoft/sp-http'; 
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISpaFeedWebpartProps {
  spHttpClient: SPHttpClient;  
  siteUrl: string;
  context:WebPartContext;
}
