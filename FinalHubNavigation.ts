import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { INavConfig, DEFAULT_NAV_CONFIG } from './INavConfig';

const LOG_SOURCE = 'HubNavigationApplicationCustomizer';
const OBSERVER_THROTTLE_MS = 60;

// Broad but safe selectors for hub/top/mega nav
const BASE_SELECTORS: string[] = [
  // Hub/Top nav anchors
  '[data-automationid="HubNav"] a',
  '[data-automationid="TopNav"] a',
  '[class*="hubNav"] a',
  '[class*="HubNav"] a',
  '[role="navigation"] a[href*="/sites/"]',

  // Hub/Top nav buttons (openers/parents)
  '[data-automationid="HubNav"] button[role="menuitem"]',
  '[data-automationid="TopNav"] button[role="menuitem"]',

  // Mega menu (anchors)
  '.ms-Callout .ms-MegaMenu-gridLayout a',

  // Fallback
  'nav a[href*="/sites/"]'
];

// Mega menu “labels” are rendered as headings or spans inside sections
const MEGA_LABEL_QUERY =
  '.ms-Callout .ms-MegaMenu-gridLayout .ms-Menu-section .ms-Menu-heading:not(a), ' +
  '.ms-Callout .ms-MegaMenu-gridLayout .ms-Menu-section .ms-Menu-heading span';

export interface IHubNavigationApplicationCustomizerProperties {
  /** Optional: Override config file path (default: SiteAssets/hub-nav-config.json) */
  configPath?: string;
}

export default class HubNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IHubNavigationApplicationCustomizerProperties> {

  private _styleElement: HTMLStyleElement | null = null;
  private _config: INavConfig = DEFAULT_NAV_CONFIG;
  private _observer: MutationObserver | null = null;
  private _throttleTimer: number | null = null;

  public async onInit(): Promise<void> {
    try {
      Log.info(LOG_SOURCE, 'Initialized');

      await this._loadConfig();

      // Optional: uncomment if you want to restrict by site
      // if (!this._isAllowedSite()) {
      //   Log.info(LOG_SOURCE, 'Customizer skipped – site not allowed by config');
      //   return;
      // }

      this._injectStyles();
      this._applyHighlighting();      // initial pass
      this._styleMegaMenuSubLabels();

      this._wireNavigation();
      this._observeDOM();             // react to future changes
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Initialization failed: ${error}`));
    }

    return Promise.resolve();
  }

  // --------------------------------------------
  // Config
  // --------------------------------------------
  private async _loadConfig(): Promise<void> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const configPath = this.properties.configPath || 'SiteAssets/hub-nav-config.json';
      const configUrl = `${siteUrl}/${configPath}`;

      const response = await this.context.spHttpClient.get(
        configUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const json = await response.json();
        this._config = {
          ...DEFAULT_NAV_CONFIG,
          ...json
        };
        Log.info(LOG_SOURCE, 'Config loaded from Site Assets');
      } else {
        Log.warn(LOG_SOURCE, `Config not found (${configUrl}). Using defaults.`);
        this._config = { ...DEFAULT_NAV_CONFIG };
      }
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to load config: ${error}`));
      this._config = { ...DEFAULT_NAV_CONFIG };
    }
  }

  // private _isAllowedSite(): boolean {
  //   try {
  //     const serverRel = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
  //     const patterns = this._config.allowedSitePatterns || [];
  //     return patterns.length === 0 ||
  //       patterns.some(p => new RegExp(p).test(serverRel));
  //   } catch {
  //     return true;
  //   }
  // }

  // --------------------------------------------
  // CSS Injection
  // --------------------------------------------
  private _injectStyles(): void {
    try {
      if (this._styleElement) this._styleElement.remove();
      this._styleElement = document.createElement('style');
      this._styleElement.setAttribute('data-hub-nav-customizer', 'true');

      const css = `
        /* Links that point to the current site */
        .hub-nav-current-site,
        .hub-nav-current-site span,
        .hub-nav-current-site button,
        a.hub-nav-current-site {
          color: ${this._config.currentSiteColor} !important;
          font-weight: ${this._config.currentSiteFontWeight} !important;
        }

        /* Other internal site links */
        .hub-nav-other-site,
        .hub-nav-other-site span,
        .hub-nav-other-site button,
        a.hub-nav-other-site {
          color: ${this._config.otherSiteColor} !important;
          font-weight: ${this._config.otherSiteFontWeight} !important;
        }

        /* Parent/openers (buttons opening submenus) */
        .hub-nav-parent,
        .hub-nav-parent span,
        .hub-nav-parent button {
          color: ${this._config.parentColor} !important;
          font-weight: ${this._config.parentFontWeight} !important;
        }

        /* Non-clickable labels (true headings) */
        .hub-nav-label,
        .hub-nav-label span,
        .hub-nav-label button {
          color: ${this._config.labelColor} !important;
          font-weight: ${this._config.labelFontWeight} !important;
        }

        /* Sub-labels (heading-like links under section header) */
        .hub-nav-sublabel,
        .hub-nav-sublabel span,
        .hub-nav-sublabel a {
          color: ${this._config.labelColor} !important;
          font-weight: ${this._config.labelFontWeight} !important;
        }

        /* External links */
        .hub-nav-external,
        a.hub-nav-external {
          color: ${this._config.externalColor} !important;
          font-weight: ${this._config.externalFontWeight} !important;
        }
           /* 🔴 Hub Navigation LABELS (non-clickable text) */
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
          font-weight: ${this._config.labelFontWeight} !important;
        }
      `;

      this._styleElement.innerHTML = css;
      document.head.appendChild(this._styleElement);
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to inject styles: ${error}`));
    }
  }

  // --------------------------------------------
  // Event/Wire-up
  // --------------------------------------------
  private _wireNavigation(): void {
    try {
      // SPFx client-side routing
      this.context.application.navigatedEvent.add(this, () => {
        try {
          setTimeout(() => this._applyHighlighting(), 50);
        } catch (error) {
          Log.error(LOG_SOURCE, new Error(`Navigation event handler failed: ${error}`));
        }
      });
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to wire navigation: ${error}`));
    }
  }
private _observeDOM(): void {
  try {
    if (this._observer) return;

    this._observer = new MutationObserver((mutations) => {
      if (this._throttleTimer) return;
      this._throttleTimer = window.setTimeout(() => {
        try {
          this._applyHighlighting();

          // If a callout was added, style labels + sublabels inside it
          mutations.forEach(m => {
            m.addedNodes.forEach(n => {
              if (n instanceof HTMLElement && n.classList.contains('ms-Callout')) {
                this._styleMegaMenuLabels();
                this._styleMegaMenuSubLabels(n); // <-- new
              }
            });
          });
        } finally {
          this._throttleTimer = null;
        }
      }, OBSERVER_THROTTLE_MS);
    });

    this._observer.observe(document.body, { childList: true, subtree: true });

    // Initial pass (covers already-open menu in rare cases)
    this._styleMegaMenuSubLabels();
  } catch (error) {
    // log if you want; safe to ignore
  }
}

  // private _observeDOM(): void {
  //   try {
  //     if (this._observer) return;

  //     this._observer = new MutationObserver((mutations) => {
  //       // Throttled re-apply
  //       if (this._throttleTimer) return;
  //       this._throttleTimer = window.setTimeout(() => {
  //         try {
  //           // Re-classify links
  //           this._applyHighlighting();
            

  //           // Style mega menu labels if a callout appeared
  //           const calloutAdded = mutations.some(m =>
  //             Array.from(m.addedNodes).some(n =>
  //               n instanceof HTMLElement && n.classList.contains('ms-Callout')
  //             )
  //           );
  //           if (calloutAdded) {
  //             this._styleMegaMenuLabels();
  //           }
  //         } finally {
  //           this._throttleTimer = null;
  //         }
  //       }, OBSERVER_THROTTLE_MS);
  //     });

  //     this._observer.observe(document.body, { childList: true, subtree: true });
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, new Error(`Failed to setup DOM observer: ${error}`));
  //   }
  // }

  // --------------------------------------------
  // Core logic
  // --------------------------------------------
  private _applyHighlighting(): void {
    try {
      const currentSiteUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '').toLowerCase();
      const currentSiteName = this._extractSiteName(currentSiteUrl);
      if (!currentSiteName) return;

      const selectorPool = [...BASE_SELECTORS, ...(this._config.extraSelectors || [])];
      const allItems = document.querySelectorAll(selectorPool.join(', '));

      allItems.forEach((el: Element) => {
        try {
          // Clear previous
          el.classList.remove(
            'hub-nav-current-site',
            'hub-nav-other-site',
            'hub-nav-parent',
            'hub-nav-label',
            'hub-nav-external'
          );

          const role = (el.getAttribute('role') || '').toLowerCase();
          const tag = el.tagName.toLowerCase();
          const href = (el.getAttribute('href') || '').trim();
          const isAnchor = tag === 'a';
          const isButton = tag === 'button' || role === 'button' || role === 'menuitem';
          const isInCallout = !!el.closest('.ms-Callout');

          // Mega labels are non-clickable headings inside callout sections
          // We’ll process them separately, but ensure we don’t misclassify here
          if (isInCallout && this._isMegaLabelNode(el)) {
            el.classList.add('hub-nav-label');
            return;
          }

          // Parent/openers (buttons that open sub-menus)
          if (!isAnchor && isButton) {
            el.classList.add('hub-nav-parent');
            return;
          }

          // Anchors classification
          if (isAnchor) {
            const lowerHref = href.toLowerCase();

            // External?
            if (this._isExternalUrl(lowerHref)) {
              el.classList.add('hub-nav-external');
              return;
            }

            // Internal site link?
            const linkSiteName = this._extractSiteName(lowerHref);

            if (linkSiteName && linkSiteName === currentSiteName) {
              el.classList.add('hub-nav-current-site');
            } else if (lowerHref.indexOf('/sites/') > -1) {
              el.classList.add('hub-nav-other-site');
            } else {
              // Not a sites/ link; leave it alone unless you want a default
            }
          }
        } catch (e) {
          Log.warn(LOG_SOURCE, `Failed to process nav element: ${e}`);
        }
      });

      // Handle mega menu labels (non-clickable)
      this._styleMegaMenuLabels();
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to apply highlighting: ${error}`));
    }
  }

  private _isMegaLabelNode(el: Element): boolean {
    // Heuristics: heading containers inside MegaMenu sections that are not anchors
    if (!(el instanceof HTMLElement)) return false;
    const isHeading = el.matches('.ms-Menu-heading, .ms-Menu-heading span');
    const isAnchor = el.tagName.toLowerCase() === 'a';
    return isHeading && !isAnchor;
  }

  private _styleMegaMenuLabels(): void {
    try {
      const labelNodes = document.querySelectorAll<HTMLElement>(MEGA_LABEL_QUERY);
      labelNodes.forEach(node => {
        // Ensure a consistent class on non-clickable labels
        node.classList.add('hub-nav-label');
        // Defensive inline styles (optional, class CSS already covers it)
        node.style.setProperty('color', this._config.labelColor, 'important');
        node.style.setProperty('font-weight', String(this._config.labelFontWeight), 'important');
      });
    } catch (error) {
      Log.warn(LOG_SOURCE, `Failed to style mega menu labels: ${error}`);
    }
  }

  /**
 * Style "sub labels" (heading-like links under section header) inside Mega Menu callouts
 */
private _styleMegaMenuSubLabels(root?: ParentNode): void {
  try {
    const scope: ParentNode = root || document;

    // Heuristics to capture common Fluent/rollout variants:
    const subLabelCandidates = scope.querySelectorAll<HTMLElement>([
      // Often seen: special heading-like anchor class
      '.ms-Callout .ms-MegaMenu-gridLayout a[class*="itemLinkMenuHeading"]',

      // Some tenants flag the menu-item as header-ish with "is-header"
      '.ms-Callout .ms-MegaMenu-gridLayout .ms-Menu-item.is-header a',

      // Fallbacks: anchors that visually sit right under a section heading
      // - immediate anchors in a Menu-section that are *not* nested items
      '.ms-Callout .ms-MegaMenu-gridLayout .ms-Menu-section > a',

      // - anchors inside a container marked as a heading
      '.ms-Callout .ms-MegaMenu-gridLayout .ms-Menu-section .ms-Menu-heading a'
    ].join(','));

    subLabelCandidates.forEach(a => {
      // Avoid class bloat
      a.classList.add('hub-nav-sublabel');

      // Optional defensive inline style (class already covers it)
      a.style.setProperty('color', this._config.labelColor, 'important');
      a.style.setProperty('font-weight', String(this._config.labelFontWeight), 'important');
    });
  } catch (e) {
    // non-fatal
  }
}

  private _extractSiteName(url: string): string {
    try {
      const idx = url.indexOf('/sites/');
      if (idx === -1) return '';
      const after = url.substring(idx + 7); // skip "/sites/"
      return after.split('/')[0] || '';
    } catch {
      return '';
    }
  }

  private _isExternalUrl(href: string): boolean {
    try {
      if (!href || href.startsWith('/') || href.startsWith('#')) return false;
      const current = new URL(this.context.pageContext.web.absoluteUrl);
      const target = new URL(href, current.origin);
      return current.host.toLowerCase() !== target.host.toLowerCase();
    } catch {
      return false;
    }
  }

  // --------------------------------------------
  // Cleanup
  // --------------------------------------------
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
