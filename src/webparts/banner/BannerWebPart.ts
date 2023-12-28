import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  // PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import * as strings from 'BannerWebPartStrings';
import Banner from './components/Banner';
import { IBannerProps } from './components/IBannerProps';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IBannerWebPartProps {
  description: string;
  bannerText: string;
  bannerImage: string;
  bannerLink: string;
  bannerHeight: number;
  fullWidth: boolean;
  useParallax: boolean;
  useParallaxInt: boolean;
  headingFontSize:number;
  textFontSize: number;
  cardOpacity: number;
  allNewLink: string;
}

export default class BannerWebPart extends BaseClientSideWebPart<IBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp: SPFI = null;
  

  public render(): void {
    const element: React.ReactElement<IBannerProps> = React.createElement(
      Banner,
      {
        ...this.properties,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        sp: this._sp,
        propertyPane: this.context.propertyPane,
        domElement: this.context.domElement,
        useParallaxInt: this.displayMode === DisplayMode.Read && !!this.properties.bannerImage && this.properties.useParallax,
        headerFontSize: this.properties.headingFontSize,
        textFontSize: this.properties.textFontSize,
        cardOpacity: this.properties.cardOpacity,
        allViewNewsLink: this.properties.allNewLink
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      this._sp = spfi().using(SPFx(this.context));
    });
  }

  private _validateImageField(imgVal: string): string {
    if (imgVal) {
      const urlSplit = imgVal.split(".");
      if (urlSplit && urlSplit.length > 0) {
        const extName = urlSplit.pop().toLowerCase();
        if (["jpg", "jpeg", "png", "gif"].indexOf(extName) === -1) {
          return "Please provide a link to an image";
        }
      }
    }
    return "";
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Shows latest happening from your feeds"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('bannerText', {
                  label: "Overlay image text (Optional)",
                  placeholder:'leave blank for no heading',
                  multiline: true,
                  maxLength: 200,
                  value: this.properties.bannerText
                }),
                PropertyPaneTextField('bannerImage', {
                  label: "Image URL *",
                  onGetErrorMessage: this._validateImageField,
                  value: this.properties.bannerImage,
                }),
                // PropertyPaneTextField('bannerLink', {
                //   label: "Link URL",
                //   value: this.properties.bannerLink
                // }),
                PropertyFieldNumber('bannerHeight', {
                  key: "bannerHeight",
                  label: "Banner height",
                  placeholder: "(420 - 520)",
                  value: this.properties.bannerHeight,
                  maxValue: 520,
                  minValue: 420
                }),
                PropertyFieldNumber('headingFontSize', {
                  key: "headingFontSize",
                  label: "Heading font size",
                  placeholder: "(10 - 16)",
                  value: this.properties.headingFontSize,
                  maxValue: 16,
                  minValue: 10
                }),
                PropertyFieldNumber('textFontSize', {
                  key: "textFontSize",
                  label: "Text font size",
                  placeholder: "(9 - 15)",
                  value: this.properties.textFontSize,
                  maxValue: 15,
                  minValue: 9
                }),
                PropertyFieldNumber('cardOpacity', {
                  key: "cardOpacity",
                  label: "Card Opacity",
                  placeholder: "(0.1 - 1)",
                  value: this.properties.cardOpacity,
                  maxValue: 1,
                  minValue: 0
                }),
                PropertyPaneTextField('allNewLink', {
                  label: "Provide view all news link",
                  placeholder:'like (https://www.google.com)',
                  multiline: true,
                  maxLength: 200,
                  value: this.properties.allNewLink
                })
                // PropertyPaneToggle('useParallax', {
                //   label: "Enable parallax effect",
                //   checked: this.properties.useParallax
                // })
              ]
            }
          ]
        }
      ]
    };
  }
}
