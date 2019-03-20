import { IWebPartContext } from "@microsoft/sp-webpart-base";

interface IShrelloAssetsViewWebPartProps {
  description: string;
  assetsProvisioned: boolean;
  context: IWebPartContext;
}

export default IShrelloAssetsViewWebPartProps;