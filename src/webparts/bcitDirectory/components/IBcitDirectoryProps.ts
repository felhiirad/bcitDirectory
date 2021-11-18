import { WebPartContext } from "@microsoft/sp-webpart-base";

import { DisplayMode } from '@microsoft/sp-core-library'; 
export interface IBcitDirectoryProps {
 
  description: string;
  siteUrl:string;
  context:WebPartContext;
}

