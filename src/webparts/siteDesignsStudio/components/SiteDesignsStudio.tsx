import * as React from 'react';
import styles from './SiteDesignsStudio.module.scss';
import { find } from '@microsoft/sp-lodash-subset';
import { SiteDesignEditor } from "../../../components/siteDesign/SiteDesignEditor";
import { SiteScriptEditor } from "../../../components/siteScript/SiteScriptEditor";
import { SiteDesignsListInContext } from "../../../components/siteDesign/SiteDesignsListInContext";
import { SiteScriptsList } from "../../../components/siteScript/SiteScriptsList";
import { IApplicationState } from '../../../app/ApplicationState';
import { App, useAppContext } from '../../../app/App';
import { Reducers } from '../../../app/ApplicationReducers';
import { ActionType, IGoToActionArgs, ISetAllAvailableSiteDesigns, ISetAllAvailableSiteScripts, IEditSiteDesignActionArgs, IEditSiteScriptActionArgs } from '../../../app/IApplicationAction';
import { Debugger } from "../../../components/common/debugger/Debugger";

import { Nav, INavLink, INav } from 'office-ui-fabric-react/lib/Nav';

import { useEffect, useState, useRef } from 'react';
import { SiteDesignsServiceKey } from '../../../services/siteDesigns/SiteDesignsService';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar } from 'office-ui-fabric-react/lib/MessageBar';
import { useTraceUpdate } from '../../../helpers/hooks';

const AppLayout = (props: any) => {
  const [appState, execute] = useAppContext<IApplicationState, ActionType>();

  const [expandedMenuItems, setExpandedMenuItems] = useState<string[]>([]);

  function isExpanded(menuItemKey: string): boolean {
    return expandedMenuItems.indexOf(menuItemKey) > -1;
  }

  function toggleExpandedMenuItem(menuItemKey: string): void {
    const expandedWithoutCurrent = expandedMenuItems.filter(m => m != menuItemKey);
    if (isExpanded(menuItemKey)) {
      setExpandedMenuItems(expandedWithoutCurrent);
    } else {
      setExpandedMenuItems([menuItemKey, ...expandedWithoutCurrent]);
    }
  }

  function _onLinkClick(ev: React.MouseEvent<HTMLElement>, item?: INavLink): void {
    if (!item) {
      return;
    }

    if (item.key.indexOf("SiteDesignEdition_") == 0) {
      const siteDesignId = item.key.replace("SiteDesignEdition_", "");
      const siteDesign = find(appState.allAvailableSiteDesigns, sd => sd.Id == siteDesignId);
      execute("EDIT_SITE_DESIGN", { siteDesign } as IEditSiteDesignActionArgs);
    } else if (item.key.indexOf("SiteScriptEdition_") == 0) {
      const siteScriptId = item.key.replace("SiteScriptEdition_", "");
      const siteScript = find(appState.allAvailableSiteScripts, ssc => ssc.Id == siteScriptId);
      execute("EDIT_SITE_SCRIPT", { siteScript } as IEditSiteScriptActionArgs);
    } else {
      execute("GO_TO", { page: item.key } as IGoToActionArgs);
    }
  }

  function _onExpandClick(ev: React.MouseEvent<HTMLElement>, item?: INavLink): void {
    if (!item) {
      return;
    }

    if (["SiteScriptsList", "SiteDesignsList"].indexOf(item.key) < 0) {
      return;
    }

    toggleExpandedMenuItem(item.key);
  }

  function _getSiteDesignNavLinks(): INavLink[] {
    if (!appState.allAvailableSiteDesigns) {
      return [];
    }

    return appState.allAvailableSiteDesigns.map(siteDesign => (
      {
        name: siteDesign.Title,
        url: '',
        expandAriaLabel: `Site Design: ${siteDesign.Title}`,
        key: `SiteDesignEdition_${siteDesign.Id}`
      }));
  }

  function _getSiteScriptNavLinks(): INavLink[] {
    if (!appState.allAvailableSiteScripts) {
      return [];
    }

    return appState.allAvailableSiteScripts.map(siteScript => (
      {
        name: siteScript.Title,
        url: '',
        expandAriaLabel: `Site Script: ${siteScript.Title}`,
        key: `SiteScriptEdition_${siteScript.Id}`
      }));
  }

  const getSelectedPage = () => {
    const page = appState.page;
    const argId = appState.page == "SiteDesignsList"
      ? appState.currentSiteDesign && appState.currentSiteDesign.Id
      : appState.page == "SiteScriptsList"
        ? appState.currentSiteScript && appState.currentSiteScript.Id
        : null;

    return `${page}${argId ? `_${argId}` : ''}`;
  };

  return (
    <div className={styles.siteDesignsStudioV2}>
      <div className={styles.layout}>
        <div className={styles.navBar}>
          <Nav
            onLinkClick={_onLinkClick}
            onLinkExpandClick={_onExpandClick}
            selectedKey={getSelectedPage()}
            styles={{
              root: {
                boxSizing: 'border-box',
                border: '1px solid #eee',
                overflowY: 'auto',
                maxWidth: "330px",
                height: "85vh"
              },
            }}
            groups={[
              {
                links: [
                  {
                    name: 'Home',
                    url: '',
                    expandAriaLabel: 'Home',
                    key: "Home",
                  },
                  {
                    name: 'Site Designs',
                    url: '',
                    key: 'SiteDesignsList',
                    isExpanded: isExpanded("SiteDesignsList"),
                    links: _getSiteDesignNavLinks()
                  },
                  {
                    name: 'Site Scripts',
                    url: '',
                    key: 'SiteScriptsList',
                    isExpanded: isExpanded("SiteScriptsList"),
                    links: _getSiteScriptNavLinks()
                  }
                ],
              },
            ]}
          />
        </div>
        <div className={styles.pageContent}>
          {props.children}
        </div>
      </div>
    </div>
  );
};


const AppPage = (props: any) => {
  useTraceUpdate("AppPage", props);
  const [appContext, action] = useAppContext<IApplicationState, ActionType>();

  const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);
  const userMessageTimeoutHandleRef = useRef(null);

  useEffect(() => {
    if (!appContext.allAvailableSiteDesigns || appContext.allAvailableSiteDesigns.length == 0) {
      siteDesignsService.getSiteDesigns().then(siteDesigns => {
        action("SET_ALL_AVAILABLE_SITE_DESIGNS", { siteDesigns } as ISetAllAvailableSiteDesigns);
      });
    }

    if (!appContext.allAvailableSiteScripts || appContext.allAvailableSiteScripts.length == 0) {
      siteDesignsService.getSiteScripts().then(siteScripts => {
        action("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts } as ISetAllAvailableSiteScripts);
      });
    }
  }, []);

  let content = null;
  switch (appContext.page) {
    case "SiteDesignEdition":
      content = <SiteDesignEditor siteDesign={appContext.currentSiteDesign } />;
      break;
    case "SiteDesignsList":
      content = <SiteDesignsListInContext />;
      break;
    case "SiteScriptEdition":
      content = <SiteScriptEditor siteScript={appContext.currentSiteScript} />;
      break;
    case "SiteScriptsList":
      content = <SiteScriptsList />;
      break;
    case "Home":
    default:
      content = <div>
        <h2><Link onClick={() => action("GO_TO", { page: "SiteDesignsList" })}>Site Designs</Link></h2>
        <SiteDesignsListInContext preview />
        <h2><Link onClick={() => action("GO_TO", { page: "SiteScriptsList" })}>Site Scripts</Link></h2>
        <SiteScriptsList preview />
      </div>;
  }

  if (window.location.search.indexOf('debug=1') >= 0) {
    return <Debugger />;
  }


  if (!userMessageTimeoutHandleRef.current && appContext.userMessage) {
    userMessageTimeoutHandleRef.current = setTimeout(() => {
      action("SET_USER_MESSAGE", { userMessage: null });
      userMessageTimeoutHandleRef.current = null;
    }, 5000);
  }

  return <>
    {appContext.userMessage && <MessageBar messageBarType={appContext.userMessage.messageType}>
      {appContext.userMessage.message}
    </MessageBar>}
    {content}
  </>;
};



export interface ISiteDesignsStudioProps {
  description: string;
  applicationState: IApplicationState;
}

export default (props: ISiteDesignsStudioProps) => {
  return <App applicationState={props.applicationState} reducers={Reducers}>
    <AppLayout>
      <AppPage />
    </AppLayout>
  </App>;
};
