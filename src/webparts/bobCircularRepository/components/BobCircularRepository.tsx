import * as React from 'react';
import styles from './BobCircularRepository.module.scss';
import type { IBobCircularRepositoryProps } from './IBobCircularRepositoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import CircularSearch from './Search/CircularSearch';
import { FluentProvider, IdPrefixProvider, webDarkTheme, webLightTheme, Theme } from "@fluentui/react-components";
import { ContextProvider } from "../DataContext/DataContext";
import Header from "./Header/Header";
import CircularForm from './CircularForm/CircularForm';
import { Constants } from '../Constants/Constants';
import EditDashBoard from './EditDashBoard/EditDashBoard';
import { Text } from '@microsoft/sp-core-library'
import { withResponsiveMode } from '@fluentui/react';

export const customLightTheme: Theme = {
  ...webLightTheme,
  colorBrandBackground: "#f26522",
  colorBrandBackgroundHover: "#f26522",
  colorBrandBackgroundSelected: "#f26522",
  colorBrandForegroundOnLightPressed: "#f26522",
  colorNeutralForeground2BrandHover: "#f26522",
  colorSubtleBackgroundHover: "#ffff",
  colorSubtleBackgroundPressed: "#ffff",
  colorNeutralForegroundDisabled: "#514b4bc7"
}

export interface IBobCircularRepositoryState {
  isHome?: boolean;
  isCreateCircular?: boolean;
  isEditCircular?: boolean;
  isPendingCompliance: boolean;
  isPendingChecker: boolean;
  isArchived?: boolean;
  isRejected?: boolean;
  isMasterCircularEdit?: boolean
}

@withResponsiveMode
export default class BobCircularRepository extends React.Component<IBobCircularRepositoryProps, IBobCircularRepositoryState> {

  private formRef;
  private navRef;

  constructor(props) {
    super(props)

    this.state = {
      isHome: true,
      isCreateCircular: false,
      isEditCircular: false,
      isPendingChecker: false,
      isPendingCompliance: false,
      isArchived: false,
      isRejected: false,
      isMasterCircularEdit: false

    }

    this.formRef = React.createRef();
    this.navRef = React.createRef();
  }


  public hideSharePointComponents() {
    const { hasTeamsContext, context } = this.props;

    setTimeout(() => {

      let bodyElement: HTMLElement = document.querySelector("body");
      bodyElement.style.cssText = "margin:0;padding:0";

      let controlZone: HTMLElement = document.querySelector(".ControlZone");
      if (controlZone != null) {
        controlZone.style.cssText = "padding:0px;margin:0px";
      }

      let canvasZone: HTMLElement = document.querySelector(".CanvasZone.row");
      if (canvasZone != null) {
        canvasZone.children[0].removeAttribute("class");
        canvasZone.style.cssText =
          "margin:0px -22px;width:auto;max-width:none;";
      }

      let commentWrapper: HTMLElement =
        document.querySelector("#CommentsWrapper");
      if (commentWrapper != null) {
        commentWrapper.style.cssText = "display:none";
      }

      let hideAppBar: HTMLElement = document.querySelector(".sp-appBar");
      if (hideAppBar != null) {
        hideAppBar.style.cssText = "display:none";
      }

      // let mainContent: HTMLElement = document.querySelector(".mainContent");
      // if (mainContent != null) {
      //   mainContent.style.cssText =
      //     "background: radial-gradient(circle, hsla(0, 0%, 100%, 1) 65%, hsla(228, 100%, 97%, 1) 100%);"; //overflow:hidden
      //   // console.log(mainContent.scrollWidth)
      // }

      let spSiteHeader: HTMLElement = document.querySelector(
        "div[id='spSiteHeader']"
      );
      if (spSiteHeader != null) {
        spSiteHeader.style.cssText = "display:none";
      }




      let divRoleMain: HTMLElement = document.querySelector("div[role='main']");
      if (divRoleMain != null) {
        divRoleMain.style.overflow = "visible";
      }

      let spCommandBar: HTMLElement = document.querySelector(
        "div[id='spCommandBar']"
      );
      if (spCommandBar != null) {
        spCommandBar.style.cssText = "display:none";
      }

      let SuiteNavWrapper: HTMLElement = document.querySelector(
        "div[id='SuiteNavWrapper']"
      );
      if (SuiteNavWrapper != null) {
        SuiteNavWrapper.style.cssText = "display:none";
      }

      let spPageChrome: HTMLElement = document.querySelector(".SPPageChrome");
      if (spPageChrome != null && hasTeamsContext) {
        spPageChrome.style.cssText = "height:auto";
      }

      let footerContainer: HTMLElement = document.querySelector(
        "div[class^='simpleFooterContainer']"
      );
      if (footerContainer != null) {
        footerContainer.style.cssText = "display:none";
      }

      //if (hasTeamsContext) {
      const head: any = document.getElementsByTagName("head")[0];
      const meta: any = document.createElement("meta");
      meta.setAttribute("name", "viewport");
      meta.setAttribute(
        "content",
        "width=device-width, initial-scale=1, maximum-scale=1,user-scalable=yes"
      );
      head.appendChild(meta);
      //}

      // console.log(this.navRef?.current)

    }, 1500);

    let workpagecontent: HTMLElement = document.querySelector(
      "div[id='workbenchPageContent']"
    );
    if (workpagecontent != null) {
      workpagecontent.style.cssText = "max-width:100%";
    }
    // 
    localStorage.setItem("loadDashBoard", "false");
  }

  public componentDidMount(): void {
    this.hideSharePointComponents();
  }

  public render(): React.ReactElement<IBobCircularRepositoryProps> {
    const { isCreateCircular, isHome, isEditCircular, isPendingCompliance, isPendingChecker, isRejected } = this.state;
    const { userInformation } = this.props
    const editCircularFilterString = Text.format(Constants.editCircularFilterString, userInformation.department);
    const complianceFilterString = Text.format(Constants.compliancePendingFilterString, userInformation.department)
    const checkerPendingFilterString = Text.format(Constants.checkerPendingFilterString, userInformation.department);
    const rejectedFilterString = Text.format(Constants.rejectedFilterString, userInformation.department);

    return (
      <>
        <IdPrefixProvider value={"CRAT_"} >
          <FluentProvider theme={customLightTheme}>
            <div className={`${styles.mainContainer} `}>
              <div className={`${styles.container}`}>
                <div className={`${styles.row}`}>
                  {/* <div className={`${styles.column2}`} ref={ref => this.navRef = ref} style={{ padding: 0 }}>
                  <ContextProvider value={this.props}>
                    <LeftNav >
                    </LeftNav>
                  </ContextProvider>
                </div> */}
                  <div className={`${styles.column12}`} ref={ref => this.formRef = ref} style={{ padding: 0 }}>
                    <ContextProvider value={this.props}>
                      <Header onMenuSubMenuLinkClick={this.onMenuSubMenuClick}></Header>
                    </ContextProvider>
                    {isHome && <ContextProvider value={this.props}>
                      <CircularSearch />
                    </ContextProvider>
                    }
                    {isCreateCircular && <>
                      <ContextProvider value={this.props}>
                        <CircularForm displayMode={Constants.lblNew}
                          onGoBack={() => {
                            this.setState({
                              isHome: true,
                              isCreateCircular: false
                            })
                          }} />
                      </ContextProvider>
                    </>
                    }
                    {
                      isEditCircular && <>
                        <ContextProvider value={this.props}>
                          <EditDashBoard
                            stateKey={new Date().toString()}
                            filterString={editCircularFilterString}
                            currentPage={`${Constants.makerGroup}`}
                          />
                        </ContextProvider>

                      </>
                    }
                    {
                      (isPendingCompliance) && <>
                        <ContextProvider value={this.props}>
                          <EditDashBoard stateKey={new Date().toString()}
                            filterString={complianceFilterString}
                            currentPage={`${Constants.complianceGroup}`}
                          />
                        </ContextProvider></>
                    }

                    {
                      (isPendingChecker) && <>
                        <ContextProvider value={this.props}>
                          <EditDashBoard stateKey={new Date().toString()}
                            filterString={checkerPendingFilterString}
                            currentPage={`${Constants.checkerGroup}`}
                          />
                        </ContextProvider></>
                    }

                    {
                      (isRejected) && <>
                        <ContextProvider value={this.props}>
                          <EditDashBoard stateKey={new Date().toString()}
                            filterString={rejectedFilterString}
                            currentPage={`${Constants.rejectedGroup}`}
                          />
                        </ContextProvider></>
                    }

                  </div>
                </div>

              </div>
            </div>
          </FluentProvider>
        </IdPrefixProvider>

      </>
    );
  }

  private onMenuSubMenuClick = (labelName) => {
    switch (labelName) {
      case Constants.lblHome: this.setState({
        isHome: true, isCreateCircular: false, isEditCircular: false,
        isPendingCompliance: false, isPendingChecker: false, isRejected: false
      });
        break;
      case Constants.lblAddCircular:
        localStorage.setItem("loadDashBoard", "false");
        this.setState({
          isHome: false, isCreateCircular: true, isEditCircular: false,
          isPendingCompliance: false, isPendingChecker: false, isRejected: false
        });
        break;
      case Constants.lblEditCircular:
        localStorage.setItem("loadDashBoard", "false");
        this.setState({
          isHome: false, isCreateCircular: false, isEditCircular: true,
          isPendingCompliance: false, isPendingChecker: false, isRejected: false
        });
        break;
      case Constants.lblPendingCompliance:
        localStorage.setItem("loadDashBoard", "true");
        this.setState({
          isHome: false, isCreateCircular: false, isEditCircular: false,
          isPendingCompliance: true, isPendingChecker: false, isRejected: false
        });
        break;
      case Constants.lblPendingChecker:
        localStorage.setItem("loadDashBoard", "true");
        this.setState({
          isHome: false, isCreateCircular: false, isEditCircular: false,
          isPendingCompliance: false, isPendingChecker: true, isRejected: false
        });
        break;
      case Constants.lblRejectedRequest:
        localStorage.setItem("loadDashBoard", "true");
        this.setState({
          isHome: false, isCreateCircular: false, isEditCircular: false,
          isPendingCompliance: false, isPendingChecker: false, isRejected: true
        });
        break;
    }

  }


}
