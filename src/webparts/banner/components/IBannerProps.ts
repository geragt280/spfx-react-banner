import { IPropertyPaneAccessor } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface IBannerProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  sp: SPFI;
  propertyPane: IPropertyPaneAccessor;
  domElement: HTMLElement;
  useParallaxInt: boolean;
  bannerText: string;
  bannerImage: string;
  bannerLink: string;
  bannerHeight: number;
  headerFontSize: number;
  textFontSize: number;
  cardOpacity: number;
  allViewNewsLink: string;
}
