// import { Log } from '@microsoft/sp-core-library';
// import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
// import { SPHttpClient } from '@microsoft/sp-http';
// import { INavConfig, DEFAULT_NAV_CONFIG } from './INavConfig';

// const LOG_SOURCE: string = 'HubNavigationApplicationCustomizer';

// /** Throttle delay for DOM observer (ms) */
// const OBSERVER_THROTTLE_MS = 300;

// /** CSS selectors for hub navigation links */
// const NAV_SELECTORS = [
//   '[data-automationid="HubNav"] a',
//   '[class*="hubNav"] a',
//   '[class*="HubNav"] a',
//   '[class*="megaMenu"] a',
//   '[class*="MegaMenu"] a',
//   '[class*="topNav"] a',
//   '[class*="TopNav"] a',
//   '[class*="CompositeHeader"] a',
//   'nav a[href*="/sites/"]',
//   '[role="navigation"] a[href*="/sites/"]'
// ];

// export interface IHubNavigationApplicationCustomizerProperties {
//   /** Optional: Override config file path (default: SiteAssets/hub-nav-config.json) */
//   configPath?: string;
// }

// /** Application Customizer to highlight current site in hub navigation */
// export default class HubNavigationApplicationCustomizer
//   extends BaseApplicationCustomizer<IHubNavigationApplicationCustomizerProperties> {

//   private _styleElement: HTMLStyleElement | null = null;
//   private _config: INavConfig = DEFAULT_NAV_CONFIG;
//   private _observer: MutationObserver | null = null;
//   private _throttleTimer: number | null = null;

//   public async onInit(): Promise<void> {
//     try {
//       Log.info(LOG_SOURCE, 'Initialized');

//       // Load configuration from Site Assets
//       await this._loadConfig();

//       // Inject CSS styles based on config
//       this._injectStyles();

//       // Apply highlighting
//       this._applyHighlighting();

//       // Re-apply on navigation events
//       this.context.application.navigatedEvent.add(this, () => {
//         try {
//           setTimeout(() => this._applyHighlighting(), 500);
//         } catch (error) {
//           Log.error(LOG_SOURCE, new Error(`Navigation event handler failed: ${error}`));
//         }
//       });

//       // Watch for DOM changes (mega menu opens) with throttling
//       this._observeDOM();

//     } catch (error) {
//       Log.error(LOG_SOURCE, new Error(`Initialization failed: ${error}`));
//     }

//     return Promise.resolve();
//   }

//   /**
//    * Load navigation config from Site Assets JSON file
//    */
//   private async _loadConfig(): Promise<void> {
//     try {
//       const siteUrl = this.context.pageContext.web.absoluteUrl;
//       const configPath = this.properties.configPath || 'SiteAssets/hub-nav-config.json';
//       const configUrl = `${siteUrl}/${configPath}`;

//       const response = await this.context.spHttpClient.get(
//         configUrl,
//         SPHttpClient.configurations.v1
//       );

//       if (response.ok) {
//         try {
//           const json = await response.json();
//           this._config = {
//             currentSiteColor: json.currentSiteColor || DEFAULT_NAV_CONFIG.currentSiteColor,
//             currentSiteFontWeight: json.currentSiteFontWeight || DEFAULT_NAV_CONFIG.currentSiteFontWeight,
//             otherSiteColor: json.otherSiteColor || DEFAULT_NAV_CONFIG.otherSiteColor,
//             otherSiteFontWeight: json.otherSiteFontWeight || DEFAULT_NAV_CONFIG.otherSiteFontWeight
//           };
//           Log.info(LOG_SOURCE, 'Config loaded from Site Assets');
//         } catch (parseError) {
//           Log.error(LOG_SOURCE, new Error(`Failed to parse config JSON: ${parseError}`));
//         }
//       } else {
//         Log.warn(LOG_SOURCE, `Config file not found at ${configUrl}, using defaults`);
//       }
//     } catch (error) {
//       Log.error(LOG_SOURCE, new Error(`Failed to load config: ${error}`));
//     }
//   }

//   /**
//    * Inject CSS styles for navigation highlighting
//    */
//   private _injectStyles(): void {
//     try {
//       if (this._styleElement) {
//         this._styleElement.remove();
//       }

//       this._styleElement = document.createElement('style');
//       this._styleElement.setAttribute('data-hub-nav-customizer', 'true');
//       this._styleElement.innerHTML = `
//         .hub-nav-current-site,
//         .hub-nav-current-site span,
//         .hub-nav-current-site button,
//         a.hub-nav-current-site {
//           color: ${this._config.currentSiteColor} !important;
//           font-weight: ${this._config.currentSiteFontWeight} !important;
//         }
        
//         .hub-nav-other-site,
//         .hub-nav-other-site span,
//         .hub-nav-other-site button,
//         a.hub-nav-other-site {
//           color: ${this._config.otherSiteColor} !important;
//           font-weight: ${this._config.otherSiteFontWeight} !important;
//         }
//       `;
//       document.head.appendChild(this._styleElement);
//     } catch (error) {
//       Log.error(LOG_SOURCE, new Error(`Failed to inject styles: ${error}`));
//     }
//   }

//   /**
//    * Apply CSS classes to navigation links based on current site
//    */
//   private _applyHighlighting(): void {
//     try {
//       const currentSiteUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '').toLowerCase();
//       const currentSiteName = this._extractSiteName(currentSiteUrl);

//       if (!currentSiteName) return;

//       const allLinks = document.querySelectorAll(NAV_SELECTORS.join(', '));

//       allLinks.forEach((link: Element) => {
//         try {
//           const href = link.getAttribute('href') || '';
//           const linkSiteName = this._extractSiteName(href.toLowerCase());

//           // Remove existing classes
//           link.classList.remove('hub-nav-current-site', 'hub-nav-other-site');

//           // Apply appropriate class
//           if (linkSiteName && linkSiteName === currentSiteName) {
//             link.classList.add('hub-nav-current-site');
//           } else if (href.indexOf('/sites/') > -1) {
//             link.classList.add('hub-nav-other-site');
//           }
//         } catch (linkError) {
//           Log.warn(LOG_SOURCE, `Failed to process link: ${linkError}`);
//         }
//       });
//     } catch (error) {
//       Log.error(LOG_SOURCE, new Error(`Failed to apply highlighting: ${error}`));
//     }
//   }

//   /**
//    * Extract site name from URL (e.g., "mysite" from "/sites/mysite/pages")
//    */
//   private _extractSiteName(url: string): string {
//     try {
//       const match = url.split('/sites/')[1];
//       return match ? match.split('/')[0] : '';
//     } catch (error) {
//       Log.warn(LOG_SOURCE, `Failed to extract site name from URL: ${error}`);
//       return '';
//     }
//   }

//   /**
//    * Observe DOM changes with throttling to handle mega menu
//    */
//   private _observeDOM(): void {
//     try {
//       if (this._observer) return;

//       this._observer = new MutationObserver(() => {
//         try {
//           // Throttle to prevent excessive calls
//           if (this._throttleTimer) return;

//           this._throttleTimer = window.setTimeout(() => {
//             try {
//               this._applyHighlighting();
//             } catch (error) {
//               Log.error(LOG_SOURCE, new Error(`Observer callback failed: ${error}`));
//             } finally {
//               this._throttleTimer = null;
//             }
//           }, OBSERVER_THROTTLE_MS);
//         } catch (error) {
//           Log.error(LOG_SOURCE, new Error(`Observer throttle failed: ${error}`));
//         }
//       });

//       this._observer.observe(document.body, {
//         childList: true,
//         subtree: true
//       });
//     } catch (error) {
//       Log.error(LOG_SOURCE, new Error(`Failed to setup DOM observer: ${error}`));
//     }
//   }

//   protected onDispose(): void {
//     try {
//       if (this._observer) {
//         this._observer.disconnect();
//         this._observer = null;
//       }
//     } catch (error) {
//       Log.warn(LOG_SOURCE, `Failed to disconnect observer: ${error}`);
//     }

//     try {
//       if (this._styleElement) {
//         this._styleElement.remove();
//         this._styleElement = null;
//       }
//     } catch (error) {
//       Log.warn(LOG_SOURCE, `Failed to remove style element: ${error}`);
//     }

//     try {
//       if (this._throttleTimer) {
//         clearTimeout(this._throttleTimer);
//         this._throttleTimer = null;
//       }
//     } catch (error) {
//       Log.warn(LOG_SOURCE, `Failed to clear throttle timer: ${error}`);
//     }
//   }
// }

import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { INavConfig, DEFAULT_NAV_CONFIG } from './INavConfig';

const LOG_SOURCE = 'HubNavigationApplicationCustomizer';
const OBSERVER_THROTTLE_MS = 60;

const NAV_SELECTORS: string[] = [
  '[data-automationid="HubNav"] a',
  '[data-automationid="TopNav"] a',
  '[class*="hubNav"] a',
  '[class*="HubNav"] a',
  '.ms-Callout .ms-MegaMenu-gridLayout a',
  'nav a'
];

export interface IHubNavigationApplicationCustomizerProperties {
  configPath?: string;
}

export default class HubNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IHubNavigationApplicationCustomizerProperties> {

  private _styleElement: HTMLStyleElement | null = null;
  private _config: INavConfig = DEFAULT_NAV_CONFIG;
  private _observer: MutationObserver | null = null;
  private _throttleTimer: number | null = null;

  // =====================================================
  // INIT
  // =====================================================
  public async onInit(): Promise<void> {
    try {
      Log.info(LOG_SOURCE, 'Initialized');

      await this._loadConfig();
      this._injectStyles();
      this._applyHighlighting();
      this._wireNavigation();
      this._observeDOM();

    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Initialization failed: ${error}`));
    }

    return Promise.resolve();
  }

  // =====================================================
  // LOAD CONFIG
  // =====================================================
  private async _loadConfig(): Promise<void> {
    try {
      const siteUrl = this.context?.pageContext?.web?.absoluteUrl?.replace(/\/$/, '');
      if (!siteUrl) return;

      const configPath = this.properties?.configPath || 'SiteAssets/hub-nav-config.json';
      const configUrl = `${siteUrl}/${configPath}`;

      const response = await this.context.spHttpClient.get(
        configUrl,
        SPHttpClient.configurations.v1
      );

      if (response?.ok) {
        const json = await response.json();
        this._config = { ...DEFAULT_NAV_CONFIG, ...json };
      } else {
        this._config = { ...DEFAULT_NAV_CONFIG };
      }

    } catch (error) {
      Log.warn(LOG_SOURCE, `Config load failed. Using defaults. ${error}`);
      this._config = { ...DEFAULT_NAV_CONFIG };
    }
  }

  // =====================================================
  // CSS
  // =====================================================
  private _injectStyles(): void {
    try {

      if (this._styleElement?.parentNode) {
        this._styleElement.parentNode.removeChild(this._styleElement);
      }

      this._styleElement = document.createElement('style');

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
        a.hub-nav-other-site  {
          color: ${this._config.otherSiteColor} !important;
          font-weight: ${this._config.otherSiteFontWeight} !important;
        }

        .hub-nav-parent,
        .hub-nav-parent span,
        .hub-nav-parent button{
          color: ${this._config.parentColor} !important;
          font-weight: ${this._config.parentFontWeight} !important;
        }

        .hub-nav-label {
          color: ${this._config.labelColor} !important;
          font-weight: ${this._config.labelFontWeight} !important;
        }

        .hub-nav-external,
        a.hub-nav-external {
          color: ${this._config.externalColor} !important;
          font-weight: ${this._config.externalFontWeight} !important;
        }
      /* ===============================
         Hub NAV LABELS ONLY (NO LINKS)
        =============================== */
        .ms-HorizontalNavItem-label[data-navigationcomponent="HubNav"]
          .ms-HorizontalNavItem-linkText {
          color: ${this._config.labelColor} !important;
          font-weight: ${this._config.labelFontWeight} !important;
        }
		
		    /* Optional hover */
        .ms-HorizontalNavItem-label[data-navigationcomponent="HubNav"]:hover
          .ms-HorizontalNavItem-linkText {
          color: ${this._config.labelColor} !important;
          font-weight: ${this._config.labelColor} !important;
        }
      `;

      document.head.appendChild(this._styleElement);

    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Style injection failed: ${error}`));
    }
  }

  // =====================================================
  // MAIN LOGIC
  // =====================================================
  private _applyHighlighting(): void {
    try {

      const currentAbsolute =
        this.context?.pageContext?.web?.absoluteUrl
          ?.toLowerCase()
          .replace(/\/$/, '');

      if (!currentAbsolute) return;

      const items = document.querySelectorAll<HTMLElement>(
        NAV_SELECTORS.join(', ')
      );

      items.forEach((el) => {
        try {

          el.classList.remove(
            'hub-nav-current-site',
            'hub-nav-other-site',
            'hub-nav-parent',
            'hub-nav-label',
            'hub-nav-external'
          );

          const tag = el.tagName?.toLowerCase();
          const href = (el.getAttribute('href') || '').trim();
          const hasPopup = el.getAttribute('aria-haspopup');

          const isAnchor = tag === 'a';
          const isButton = tag === 'button';

          if (isButton || hasPopup) {
            el.classList.add('hub-nav-parent');
            return;
          }

          if (!href || href === '#' || href.startsWith('javascript')) {
            el.classList.add('hub-nav-label');
            return;
          }

          if (!isAnchor) return;

          let resolved: URL;
          try {
            resolved = new URL(href, window.location.origin);
          } catch {
            el.classList.add('hub-nav-label');
            return;
          }

          const targetAbsolute =
            resolved.href.toLowerCase().replace(/\/$/, '');

          const currentHost = window.location.host.toLowerCase();

          if (resolved.host.toLowerCase() !== currentHost) {
            el.classList.add('hub-nav-external');
            return;
          }

          if (targetAbsolute.startsWith(currentAbsolute)) {
            el.classList.add('hub-nav-current-site');
            return;
          }

          el.classList.add('hub-nav-other-site');

        } catch (innerError) {
          Log.warn(LOG_SOURCE, `Element processing failed: ${innerError}`);
        }
      });

      this._applySectionHeaderLabels();

    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Highlighting failed: ${error}`));
    }
  }

  // =====================================================
  // SECTION HEADERS & NESTED LABELS
  // =====================================================
  private _applySectionHeaderLabels(): void {
    try {

      const sections = document.querySelectorAll<HTMLElement>(
        '.ms-Callout .ms-MegaMenu-gridLayout .ms-Menu-section'
      );

      sections.forEach(section => {
        try {

          const header = section.querySelector<HTMLElement>('.ms-Menu-heading');
          if (!header) return;

          header.classList.add('hub-nav-label');

          const nodes = section.querySelectorAll<HTMLElement>('span, a, button');

          nodes.forEach(node => {
            try {

              if (node === header) return;

              const anchorParent = node.closest<HTMLAnchorElement>('a[href]');

              if (anchorParent) {
                const href = anchorParent.getAttribute('href')?.trim() || '';

                const isRealLink =
                  href &&
                  href !== '#' &&
                  !href.startsWith('javascript');

                if (isRealLink) {
                  node.classList.remove('hub-nav-label');
                  return;
                }
              }

              if (
                node.tagName.toLowerCase() === 'button' &&
                node.getAttribute('aria-haspopup')
              ) {
                return;
              }

              const href = node.getAttribute('href')?.trim() || '';

              if (!href || href === '#' || href.startsWith('javascript')) {
                node.classList.add('hub-nav-label');
              }

            } catch (nodeError) {
              Log.warn(LOG_SOURCE, `Nested label processing failed: ${nodeError}`);
            }
          });

        } catch (sectionError) {
          Log.warn(LOG_SOURCE, `Section processing failed: ${sectionError}`);
        }
      });

    } catch (error) {
      Log.warn(LOG_SOURCE, `Section header styling failed: ${error}`);
    }
  }

  // =====================================================
  // SPA NAVIGATION
  // =====================================================
  private _wireNavigation(): void {
    try {
      this.context.application.navigatedEvent.add(this, () => {
        try {
          setTimeout(() => this._applyHighlighting(), 50);
        } catch (navError) {
          Log.warn(LOG_SOURCE, `Navigation event failed: ${navError}`);
        }
      });
    } catch (error) {
      Log.warn(LOG_SOURCE, `Navigation wiring failed: ${error}`);
    }
  }

  // =====================================================
  // DOM OBSERVER
  // =====================================================
  private _observeDOM(): void {
    try {

      if (this._observer) return;

      this._observer = new MutationObserver(() => {
        try {

          if (this._throttleTimer) return;

          this._throttleTimer = window.setTimeout(() => {
            try {
              this._applyHighlighting();
            } catch (observerError) {
              Log.warn(LOG_SOURCE, `Observer highlight failed: ${observerError}`);
            } finally {
              this._throttleTimer = null;
            }
          }, OBSERVER_THROTTLE_MS);

        } catch (mutationError) {
          Log.warn(LOG_SOURCE, `Mutation handling failed: ${mutationError}`);
        }
      });

      this._observer.observe(document.body, {
        childList: true,
        subtree: true
      });

    } catch (error) {
      Log.warn(LOG_SOURCE, `Observer setup failed: ${error}`);
    }
  }

  // =====================================================
  // CLEANUP
  // =====================================================
  protected onDispose(): void {
    try {

      if (this._observer) {
        this._observer.disconnect();
        this._observer = null;
      }

      if (this._styleElement?.parentNode) {
        this._styleElement.parentNode.removeChild(this._styleElement);
        this._styleElement = null;
      }

      if (this._throttleTimer) {
        clearTimeout(this._throttleTimer);
        this._throttleTimer = null;
      }

    } catch (error) {
      Log.warn(LOG_SOURCE, `Dispose failed: ${error}`);
    }
  }
}
