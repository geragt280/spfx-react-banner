import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import * as React from "react";
import styles from "./Banner.module.scss";
import { IBannerProps } from "./IBannerProps";
import {
  DocumentCard,
  // DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardDetails,
  DocumentCardTitle,
  // IDocumentCardPreviewProps,
  // DocumentCardLocation,
  DocumentCardType,
} from "office-ui-fabric-react/lib/DocumentCard";
import { ImageFit } from "office-ui-fabric-react/lib/Image";
import "./card.css";
import {
  // Stack, IStackTokens,
  IStackProps,
  css,
  // Text,
} from "office-ui-fabric-react";

interface IBannerStates {
  NewsItems: any[];
}

export default class Banner extends React.Component<
  IBannerProps,
  IBannerStates
> {
  private _scrollElm: HTMLElement = null;
  private _scrollElmRect: ClientRect = null;
  private _parallaxElm: HTMLElement = null;
  private horizontalAlignment: IStackProps["horizontalAlign"];
  private verticalAlignment: IStackProps["verticalAlign"];
  // private wrapStackTokens: IStackTokens = { childrenGap: 20 };

  // private stackStyles: IStackStyles = {
  //   root: {
  //     overflow: 'hidden',
  //     width: `100%`,
  //   },
  // };

  // private items: any[] = [
  //   {
  //     thumbnail:
  //       "https://shell.cdn.office.net/shellux/images/beach.65c12d3945a2f93e82607434bc6f5018.gif",
  //     title: "Adventures in SPFx",
  //     name: "Perry Losselyong",
  //     profileImageSrc:
  //       "https://shell.cdn.office.net/shellux/images/beach.65c12d3945a2f93e82607434bc6f5018.gif",
  //     location: "SharePoint",
  //     activity: "3/13/2019",
  //   },
  //   {
  //     thumbnail:
  //       "https://shell.cdn.office.net/shellux/images/beach.65c12d3945a2f93e82607434bc6f5018.gif",
  //     title: "The Wild, Untold Story of SharePoint!",
  //     name: "Ebonee Gallyhaock",
  //     profileImageSrc:
  //       "https://shell.cdn.office.net/shellux/images/beach.65c12d3945a2f93e82607434bc6f5018.gif",
  //     location: "SharePoint",
  //     activity: "6/29/2019",
  //   },
  //   {
  //     thumbnail:
  //       "https://shell.cdn.office.net/shellux/images/beach.65c12d3945a2f93e82607434bc6f5018.gif",
  //     title: "Low Code Solutions: PowerApps",
  //     name: "Seward Keith",
  //     profileImageSrc:
  //       "https://shell.cdn.office.net/shellux/images/beach.65c12d3945a2f93e82607434bc6f5018.gif",
  //     location: "PowerApps",
  //     activity: "12/31/2018",
  //   },
  // ];

  constructor(props: IBannerProps) {
    super(props);

    this.state = {
      NewsItems: [],
    };

    void this.getLast3News();
  }

  async getLast3News(): Promise<void> {
    const { sp } = this.props;
    const newsItems = await sp.web.lists
      .getByTitle("Site Pages")
      .items.select(
        "Title",
        "BannerImageUrl",
        "Description",
        "Created",
        "FileRef"
      )
      .filter("PromotedState eq '2'")();

    this.setState({
      NewsItems: newsItems.slice(-4),
    });
    console.log("All news", newsItems);
  }

  formatDate(inputDate: string): string {
    const date = new Date(inputDate);
    let day = date.getDate().toString();
    let month = (date.getMonth() + 1).toString(); // Months are zero-based
    const year = date.getFullYear().toString();

    // Ensure two-digit day and month
    if (day.length === 1) {
      day = "0" + day;
    }

    if (month.length === 1) {
      month = "0" + month;
    }

    return `${day}/${month}/${year}`;
  }

  private _onConfigure = () => {
    this.props.propertyPane.open();
  };

  private _getScrollableParent(): HTMLElement {
    const scrollElm = document.querySelector(
      'div[data-is-scrollable="true"]'
    ) as HTMLElement;
    if (scrollElm) {
      return scrollElm;
    }
    return null;
  }

  private _setTranslate(vector: number) {
    const r = `translate3d(0px, ${vector}px, 0px)`;
    this._parallaxElm.style.transform = r;
  }

  private _setParallaxEffect = () => {
    window.requestAnimationFrame(() => {
      const scrollElmTop = this._scrollElmRect.top;
      const clientElmRect = this.props.domElement.getBoundingClientRect();
      const clientElmTop = clientElmRect.top;
      const clientElmBottom = clientElmRect.bottom;

      if (clientElmTop < scrollElmTop && clientElmBottom > scrollElmTop) {
        const vector = Math.round((scrollElmTop - clientElmTop) / 1.81);
        this._setTranslate(vector);
      } else if (clientElmTop >= scrollElmTop) {
        this._setTranslate(0);
      }
    });
  };

  private _removeParallaxBinding() {
    if (this._scrollElm) {
      // Unbind the scroll event
      this._scrollElm.removeEventListener("scroll", this._setParallaxEffect);
    }
  }

  private _parallaxBinding() {
    if (this.props.useParallaxInt) {
      this._scrollElm = this._getScrollableParent();
      this._parallaxElm = this.props.domElement.querySelector(
        `.${styles.bannerImg}`
      ) as HTMLElement;
      if (this._scrollElm && this._parallaxElm) {
        // Get client rect info
        this._scrollElmRect = this._scrollElm.getBoundingClientRect();
        // Bind the scroll event
        this._scrollElm.addEventListener("scroll", this._setParallaxEffect);
      }
    } else {
      this._removeParallaxBinding();
    }
  }

  public componentDidMount(): void {
    console.log(
      "Alignment got",
      this.verticalAlignment,
      this.horizontalAlignment
    );
    this._parallaxBinding();
  }

  public componentDidUpdate(prevProps: IBannerProps): void {
    this._parallaxBinding();
  }

  public componentWillUnmount(): void {
    this._removeParallaxBinding();
  }

  private _onRenderGridItem = (item: any, index: number): JSX.Element => {
    const { headerFontSize, textFontSize, allViewNewsLink, cardOpacity } = this.props;
    console.error("debug");
    const truncatedString =
      item.Description?.length > 200
        ? item.Description.slice(0, 200) + "..."
        : item.Description;
    if (index === 1 || index === 3) {
      return (
        <div
          className={styles.msGridcol}
          data-is-focusable={true}
          role="listitem"
          aria-label={item.title}
        >
          <DocumentCard
            styles={{
              root: {
                padding: 10,
                // height: 140
                maxWidth: "100%",
                backgroundColor: `rgba(255, 255, 255, ${cardOpacity ? cardOpacity : 0.8})`
              },
            }}
            className={styles.singleItemsBackGround}
            aria-label="Document Card with document preview. Revenue stream proposal fiscal year 2016 version 2.
          Created by Roko Kolar a few minutes ago"
            type={DocumentCardType.compact}
            onClickHref={window.location.origin + item.FileRef}
            onClickTarget={"tab"}
          >
            <DocumentCardDetails
              styles={{
                root: {
                  minWidth: 340,
                  justifyContent: "flex-start",
                },
              }}
            >
              <div
                style={{
                  display: "flex",
                  flexDirection: "row",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <DocumentCardTitle
                  title={item.Title}
                  styles={{
                    root: { height: 27, fontSize: `${headerFontSize}px !Important` },
                  }}
                  className={styles.singleItemsDate}
                />
                <DocumentCardTitle
                  title={this.formatDate(item.Created)}
                  styles={{
                    root: { height: 15, fontSize: `${headerFontSize}px !important` },
                  }}
                  className={styles.singleItemsDate}
                />
              </div>
              <DocumentCardTitle
                title={truncatedString}
                className={styles.singleItemsDescription}
                styles={{
                  root: { height: 100, fontSize: `${textFontSize}px !important` },
                }}
                showAsSecondaryTitle
              />
              {item.Description && item.Description.length > 200 && (
                <DocumentCardTitle
                  title={"See more"}
                  className={styles.singleItemsSeeMore}
                  showAsSecondaryTitle
                />
              )}
            </DocumentCardDetails>
            {item.BannerImageUrl && item.BannerImageUrl.Url && (
              <DocumentCardPreview
                previewImages={[
                  {
                    previewImageSrc: item.BannerImageUrl.Url,
                    imageFit: ImageFit.cover,
                    height: 130,
                    width: 130,
                  },
                ]}
              />
            )}
          </DocumentCard>
          <div
            style={{
              height: index === 3 ? 10 : 5,
              fontSize: index === 3 ? 10 : 0,
              padding: index === 3 ? 5 : 0,
              backgroundColor: "#000",
              
            }}
          >
            {index === 3 ? <a href={allViewNewsLink} style={{cursor: 'pointer'}} target={"_blank"} >View All News</a> : ""}
          </div>
        </div>
      );
    } else {
      return (
        <div
          className={styles.msGridcol}
          data-is-focusable={true}
          role="listitem"
          aria-label={item.title}
        >
          <DocumentCard
            styles={{ root: { 
              padding: 10, 
              maxWidth: "100%",
              backgroundColor: `rgba(255, 255, 255, ${cardOpacity ? cardOpacity : 0.8})` 
            }}}
            className={styles.singleItemsBackGround}
            aria-label="Document Card with document preview. Revenue stream proposal fiscal year 2016 version 2.
          Created by Roko Kolar a few minutes ago"
            type={DocumentCardType.compact}
            onClickHref={window.location.origin + item.FileRef}
            onClickTarget={"tab"}
          >
            <DocumentCardDetails
              styles={{
                root: {
                  minWidth: 340,
                  justifyContent: "flex-start",
                },
              }}
            >
              <div
                style={{
                  display: "flex",
                  flexDirection: "row",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <DocumentCardTitle
                  title={item.Title}
                  styles={{
                    root: { height: 27, fontSize: `${headerFontSize}px !important` },
                  }}
                  className={styles.singleItemsDate}
                />
                <DocumentCardTitle
                  title={this.formatDate(item.Created)}
                  styles={{
                    root: { height: 15, fontSize: `${headerFontSize}px !important` },
                  }}
                  className={styles.singleItemsDate}
                />
              </div>
              <DocumentCardTitle
                title={truncatedString}
                className={styles.singleItemsDescription}
                styles={{ root: { fontSize: `${textFontSize}px !important` } }}
                showAsSecondaryTitle
              />
              {item.Description && item.Description.length > 200 && (
                <DocumentCardTitle
                  title={"See more"}
                  className={styles.singleItemsSeeMore}
                  showAsSecondaryTitle
                />
              )}
            </DocumentCardDetails>
            {item.BannerImageUrl && item.BannerImageUrl.Url && (
              <DocumentCardPreview
                previewImages={[
                  {
                    previewImageSrc: item.BannerImageUrl.Url,
                    imageFit: ImageFit.cover,
                    height: 130,
                    width: 130,
                  },
                ]}
              />
            )}
          </DocumentCard>
          <div style={{ height: 5, backgroundColor: "#000" }}></div>
        </div>
      );
    }
  };

  public render(): React.ReactElement<IBannerProps> {
    const { bannerImage, bannerHeight, bannerLink, bannerText } = this.props;
    const { NewsItems } = this.state;

    if (this.props.bannerImage) {
      return (
        <div
          className={styles.banner}
          style={{
            height: bannerHeight ? `${bannerHeight}px` : `400px`,
          }}
        >
          <div
            className={styles.bannerImg}
            style={{
              backgroundImage: `url('${bannerImage}')`,
            }}
          ></div>
          <div className={styles.bannerOverlay}></div>
          <div className={css(styles.bannerText, styles.msGrid)}>
            {bannerLink ? (
              <a href={bannerLink} title={escape(bannerText)}>
                {escape(bannerText)}
              </a>
            ) : (
              <span>{bannerText}</span>
            )}
            <div
              className={css("bbg", styles.msGridrow)}
              // style={{ height: 320 }}
            >
              <div>
                {NewsItems.map((item, index) => {
                  return this._onRenderGridItem(item, index);
                })}
              </div>
            </div>
          </div>
        </div>
      );
    } else {
      return (
        <Placeholder
          iconName="ImagePixel"
          iconText={"Configure your web part"}
          description={"Please specify the banner configuration."}
          buttonLabel={"Configure"}
          onConfigure={this._onConfigure}
        />
      );
    }
  }
}
