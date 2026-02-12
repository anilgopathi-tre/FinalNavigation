import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { INavConfig, DEFAULT_NAV_CONFIG } from './INavConfig';

const LOG_SOURCE: string = 'HubNavigationApplicationCustomizer';

/** Throttle delay for DOM observer (ms) */
const OBSERVER_THROTTLE_MS = 50;

/** CSS selectors for hub navigation links */
const NAV_SELECTORS = [
  '[data-automationid="HubNav"] a',
  '[class*="hubNav"] a',
  '[class*="HubNav"] a',
  '[class*="megaMenu"] a',
  '[class*="MegaMenu"] a',
  '[class*="topNav"] a',
  '[class*="TopNav"] a',
  '[class*="CompositeHeader"] a',
  'nav a[href*="/sites/"]',
  '[role="navigation"] a[href*="/sites/"]'
];

export interface IHubNavigationApplicationCustomizerProperties {
  /** Optional: Override config file path (default: SiteAssets/hub-nav-config.json) */
  configPath?: string;
}

/** Application Customizer to highlight current site in hub navigation */
export default class HubNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IHubNavigationApplicationCustomizerProperties> {

  private _styleElement: HTMLStyleElement | null = null;
  private _config: INavConfig = DEFAULT_NAV_CONFIG;
  private _observer: MutationObserver | null = null;
  private _throttleTimer: number | null = null;

  public async onInit(): Promise<void> {
    try {

      
// const style = `
//     .ms-HubNav .ms-HubNav-linkButton.is-header {
//         color: #008000 !important;
//         font-weight: 600 !important;
//     }
//   `;
//   const head = document.head || document.getElementsByTagName("head")[0];
//   const styleTag = document.createElement("style");
//   styleTag.innerHTML = style;
//   head.appendChild(styleTag);



      // 🔒 Exit early if site is not allowed
    // if (!this._isAllowedSite()) {
    //   Log.info(LOG_SOURCE, 'Hub Navigation Customizer skipped – site not in allowed list');
    //   return;
    // }
      Log.info(LOG_SOURCE, 'Initialized');

      // Load configuration from Site Assets
      await this._loadConfig();

      // Inject CSS styles based on config
      this._injectStyles();
      this._watchMegaMenu();
      //this._styleMegaMenuLabels();

      // Apply highlighting
      this._applyHighlighting();

      // Re-apply on navigation events
      this.context.application.navigatedEvent.add(this, () => {
        try {
          setTimeout(() => this._applyHighlighting(), 50);
        } catch (error) {
          Log.error(LOG_SOURCE, new Error(`Navigation event handler failed: ${error}`));
        }
      });

      // Watch for DOM changes (mega menu opens) with throttling
      this._observeDOM();

    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Initialization failed: ${error}`));
    }

    return Promise.resolve();
  }

  /**
 * Check whether the current site URL is allowed to run this extension
 */
// private _isAllowedSite(): boolean {
//   const siteUrl = this.context.pageContext.web.serverRelativeUrl
//     .toLowerCase()
//     .replace(/\/$/, '');

//   const allowedPatterns: RegExp[] = [
//     /^\/sites\/int$/,            // /sites/int
//     /^\/sites\/int-.+/,          // /sites/int-*
//     /^\/sites\/uat-int$/,        // /sites/uat-int
//     /^\/sites\/uat-int-.+/       // /sites/uat-int-*
//   ];

//   return allowedPatterns.some(pattern => pattern.test(siteUrl));
// }

  /**
   * Load navigation config from Site Assets JSON file
   */
  private async _loadConfig(): Promise<void> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const configPath = this.properties.configPath || 'SiteAssets/hub-nav-config.json';
      const configUrl = `${siteUrl}/${configPath}`;

      const response = await this.context.spHttpClient.get(
        configUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        try {
          const json = await response.json();
          this._config = {
            currentSiteColor: json.currentSiteColor || DEFAULT_NAV_CONFIG.currentSiteColor,
            currentSiteFontWeight: json.currentSiteFontWeight || DEFAULT_NAV_CONFIG.currentSiteFontWeight,
            otherSiteColor: json.otherSiteColor || DEFAULT_NAV_CONFIG.otherSiteColor,
            otherSiteFontWeight: json.otherSiteFontWeight || DEFAULT_NAV_CONFIG.otherSiteFontWeight
          };
          Log.info(LOG_SOURCE, 'Config loaded from Site Assets');
        } catch (parseError) {
          Log.error(LOG_SOURCE, new Error(`Failed to parse config JSON: ${parseError}`));
        }
      } else {
        Log.warn(LOG_SOURCE, `Config file not found at ${configUrl}, using defaults`);
      }
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to load config: ${error}`));
    }
  }

  /**
   * Inject CSS styles for navigation highlighting
   */
  private _injectStyles(): void {
    try {
      if (this._styleElement) {
        this._styleElement.remove();
      }

      this._styleElement = document.createElement('style');
      this._styleElement.setAttribute('data-hub-nav-customizer', 'true');
      this._styleElement.innerHTML = `
        .hub-nav-current-site,
        .hub-nav-current-site span,
        .hub-nav-current-site button,
        a.hub-nav-current-site {
          color: ${this._config.currentSiteColor} !important;
          font-weight: ${this._config.currentSiteFontWeight} !important;
        }
        
        .hub-nav-other-site,
        .hub-nav-other-site span,
        .hub-nav-other-site button,
        a.hub-nav-other-site {
          color: ${this._config.otherSiteColor} !important;
          font-weight: ${this._config.otherSiteFontWeight} !important;
        }
          /* 🔴 Hub Navigation LABELS (non-clickable text) */
        /* ===============================
         Hub NAV LABELS ONLY (NO LINKS)
      =============================== */
      .ms-HorizontalNavItem-label[data-navigationcomponent="HubNav"]
        .ms-HorizontalNavItem-linkText {
        color: #d13438 !important;
        font-weight: 600;
      }

      /* Optional hover */
      .ms-HorizontalNavItem-label[data-navigationcomponent="HubNav"]:hover
        .ms-HorizontalNavItem-linkText {
        color: #a4262c !important;
      }
        
      
      

      
      `;
      document.head.appendChild(this._styleElement);
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to inject styles: ${error}`));
    }
  }

  private _watchMegaMenu(): void {

  const processCallout = (callout: HTMLElement) => {

    const items = callout.querySelectorAll<HTMLAnchorElement>(
      '.ms-MegaMenu-gridLayout a'
    );

    const sections = callout.querySelectorAll<HTMLElement>(
      '.ms-MegaMenu-gridLayout .ms-Menu-section[aria-label]'
    );

    sections.forEach(section => {
      //const labelText=Array.from(section.classList).some(cls=>cls.startsWith('spfx-mega-label'));
      //if(labelText)return;  
      if(section.classList.contains('spfx-mega-label'))
	  {
		  section.classList.add('spfx-mega-label');
		  section.style.color='#d13438';  
		  section.style.fontWeight='600'; 
		  section.style.pointerEvents='none';
		  section.style.cursor='default'; 
	  }
      // const labelText = section.getAttribute('aria-label');
      // if (!labelText) return;
      // const label = document.createElement('div');
      // label.className = 'spfx-mega-label';
      // label.textContent = labelText;
      // section.parentElement?.insertBefore(label, section);
	  //return;
    });


    items.forEach(item => {

      const isLabel =
        Array.from(item.classList).some(cls =>
          cls.startsWith('itemLinkMenuHeading')
        );

      if (!isLabel) return;

      if (item.classList.contains('spfx-label-highlight')) return;

      item.classList.add('spfx-label-highlight');
    });
  };

  const observer = new MutationObserver((mutations: MutationRecord[]) => {
    mutations.forEach(mutation => {
      mutation.addedNodes.forEach((node: Node) => {
        if (
          node instanceof HTMLElement &&
          node.classList.contains('ms-Callout')
        ) {
          processCallout(node);
        }
      });
    });
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
}
  /**
   * Watch for mega menu openings to style labels
   */
   
//   private _watchMegaMenu(): void {

//   // const applyMegaMenuLabelStyling = (callout: HTMLElement) => {

//   //   const sections = callout.querySelectorAll<HTMLElement>(
//   //     '.ms-MegaMenu-gridLayout .ms-Menu-section[aria-label]'
//   //   );

//   //   sections.forEach(section => {

//   //     const labelText = section.getAttribute('aria-label');
//   //     if (!labelText) return;

//   //     // Fluent UI renders header using pseudo element
//   //     // We add a class to section itself
//   //     if (!section.classList.contains('spfx-mega-header')) {
//   //       section.classList.add('spfx-mega-header');
//   //     }
//   //   });
//   // };
// //  const selector = ['div[data-automationid="TopNav"][role="menuitem"] span',
// //   'div[data-automationid="TopNav"][role="menuitem"] a',
// //   'div[data-automationid="HubTopNav"][role="menuitem"] button'
// //  ].join(', ');


//   const applyMegaMenuLabelStyling = (callout: HTMLElement) => {

//     // const sections = callout.querySelectorAll<HTMLElement>(
//     //   '.ms-MegaMenu-gridLayout .ms-Menu-section[aria-label]'
//     // );
//     const sections = callout.querySelectorAll<HTMLElement>(selector);

//     sections.forEach(section => {
//       const label=(section.textContent || '').trim() ;  
//       if (!label) return;
    

//       //const labelText = section.getAttribute('aria-label');
//       //if (!labelText) return;

//       // Fluent UI renders header using pseudo element
//       // We add a class to section itself
//       if (!section.classList.contains('spfx-mega-header')) {
//        section.classList.add('spfx-mega-header');
//       }
//     });
//   }

//   const observer = new MutationObserver((mutations: MutationRecord[])=> {
//     mutations.forEach(mutation => {
//       mutation.addedNodes.forEach((node: Node) => {
//         if (
//           node instanceof HTMLElement &&
//           node.classList.contains('ms-Callout')
//         ) {
//           applyMegaMenuLabelStyling(node);
//         }
//       });
//     });
//   });

//   observer.observe(document.body, {
//     childList: true,
//     subtree: true
//   });
// }




// private _styleMegaMenuLabels(): void {

//   const processCallout = (callout: HTMLElement) => {

//     // Remove old injected labels
//     callout.querySelectorAll('.spfx-mega-label').forEach(e => e.remove());

//     const sections = callout.querySelectorAll<HTMLElement>(
//       '.ms-MegaMenu-gridLayout .ms-Menu-section[aria-label]'
//     );

//     sections.forEach(section => {
//       const labelText = section.getAttribute('aria-label');
//       if (!labelText) return;

//       const label = document.createElement('div');
//       label.className = 'spfx-mega-label';
//       label.textContent = labelText;

//       section.parentElement?.insertBefore(label, section);
//     });
//   };

//   const observer = new MutationObserver((mutations: MutationRecord[]) => {
//     mutations.forEach((mutation: MutationRecord) => {
//       mutation.addedNodes.forEach((node: Node) => {
//         if (
//           node instanceof HTMLElement &&
//           node.classList.contains('ms-Callout')
//         ) {
//           processCallout(node);
//         }
//       });
//     });
//   });

//   observer.observe(document.body, {
//     childList: true,
//     subtree: true
//   });
// }



  /**
   * Apply CSS classes to navigation links based on current site
   */
  private _applyHighlighting(): void {
    try {
      const currentSiteUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '').toLowerCase();
      const currentSiteName = this._extractSiteName(currentSiteUrl);

      if (!currentSiteName) return;

      const allLinks = document.querySelectorAll(NAV_SELECTORS.join(', '));

      allLinks.forEach((link: Element) => {
        try {
          const href = link.getAttribute('href') || '';
          const linkSiteName = this._extractSiteName(href.toLowerCase());

          // Remove existing classes
          link.classList.remove('hub-nav-current-site', 'hub-nav-other-site');

          // Apply appropriate class
          if (linkSiteName && linkSiteName === currentSiteName) {
            link.classList.add('hub-nav-current-site');
          } else if (href.indexOf('/sites/') > -1) {
            link.classList.add('hub-nav-other-site');
          }
        } catch (linkError) {
          Log.warn(LOG_SOURCE, `Failed to process link: ${linkError}`);
        }
      });

    //   // =====================================
    // // 2️⃣ NEW: MEGA MENU LABEL LOGIC
    // // =====================================
    // const megaSections = document.querySelectorAll<HTMLElement>(
    //   '.ms-Callout .ms-MegaMenu-gridLayout .ms-Menu-section[aria-label]'
    // );
    // megaSections.forEach(section => {
    //   // Avoid reprocessing
    //   if (section.classList.contains('spfx-mega-header')) return;
    //   section.classList.add('spfx-mega-header');
    // });


    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to apply highlighting: ${error}`));
    }
  }

  /**
   * Extract site name from URL (e.g., "mysite" from "/sites/mysite/pages")
   */
  private _extractSiteName(url: string): string {
    try {
      const match = url.split('/sites/')[1];
      return match ? match.split('/')[0] : '';
    } catch (error) {
      Log.warn(LOG_SOURCE, `Failed to extract site name from URL: ${error}`);
      return '';
    }
  }

  /**
   * Observe DOM changes with throttling to handle mega menu
   */
  private _observeDOM(): void {
    try {
      if (this._observer) return;

      this._observer = new MutationObserver(() => {
        try {
          // Throttle to prevent excessive calls
          if (this._throttleTimer) return;

          this._throttleTimer = window.setTimeout(() => {
            try {
              this._applyHighlighting();
            } catch (error) {
              Log.error(LOG_SOURCE, new Error(`Observer callback failed: ${error}`));
            } finally {
              this._throttleTimer = null;
            }
          }, OBSERVER_THROTTLE_MS);
        } catch (error) {
          Log.error(LOG_SOURCE, new Error(`Observer throttle failed: ${error}`));
        }
      });

      this._observer.observe(document.body, {
        childList: true,
        subtree: true
      });
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to setup DOM observer: ${error}`));
    }
  }

  protected onDispose(): void {
    try {
      if (this._observer) {
        this._observer.disconnect();
        this._observer = null;
      }
    } catch (error) {
      Log.warn(LOG_SOURCE, `Failed to disconnect observer: ${error}`);
    }

    try {
      if (this._styleElement) {
        this._styleElement.remove();
        this._styleElement = null;
      }
    } catch (error) {
      Log.warn(LOG_SOURCE, `Failed to remove style element: ${error}`);
    }

    try {
      if (this._throttleTimer) {
        clearTimeout(this._throttleTimer);
        this._throttleTimer = null;
      }
    } catch (error) {
      Log.warn(LOG_SOURCE, `Failed to clear throttle timer: ${error}`);
    }
  }
}
