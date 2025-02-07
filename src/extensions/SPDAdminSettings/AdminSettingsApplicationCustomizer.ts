import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";

import * as React from "react";
import * as ReactDOM from "react-dom";
import { IAdminPanelProps } from "./components/IAdminPanelProps";
import AdminPanel from "./components/AdminPanel";
import { ClientsideText, sp, Web } from "@pnp/sp/presets/all";

export interface IAdminSettingsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string | null;
  listID: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AdminSettingsApplicationCustomizer extends BaseApplicationCustomizer<IAdminSettingsApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private loginUser: string = "";
  private groupUsers: string[] = [];
  private handleNavigation: () => void;

  public async onInit(): Promise<void> {
    try {
      await super.onInit();
      sp.setup({
        spfxContext: this.context as any,
      });

      this.handleNavigation = this.onNavigated.bind(this);
      this.context.application.navigatedEvent.add(this, this.handleNavigation);

      await this.getAssociatedGroupUsers();
      await this.renderAdminPanelIfNeeded();
    } catch (error) {
      console.error("Initialization failed:", error);
    }

    return Promise.resolve();
  }
  private async onNavigated(): Promise<void> {
    try {
      await this.getAssociatedGroupUsers();
      await this.renderAdminPanelIfNeeded();
    } catch (error) {
      console.error("Error during navigation handling:", error);
    }
  }
  private async renderAdminPanelIfNeeded(): Promise<void> {
    if (
      this.groupUsers.includes(this.loginUser) ||
      this.context.pageContext.legacyPageContext?.isSiteAdmin
    ) {
      this.context.placeholderProvider.changedEvent.add(
        this,
        this._renderPlaceHolders
      );
      await this._renderPlaceHolders();
      await this.createAdminPage(); // Create admin page if necessary
    }
  }
  private async _renderPlaceHolders(): Promise<void> {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      if (this.properties) {
        let topString: string | null = this.properties.testMessage;

        if (!topString) {
          topString = "(Top property was not defined.)";
        }

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = ``;

          const element: React.ReactElement<IAdminPanelProps> =
            React.createElement(AdminPanel, {
              context: this.context,
              topString: topString,
            });
          ReactDOM.render(element, this._topPlaceholder.domElement);
        }
      }
    } else {
      if (this.properties) {
        let topString: string | null = new Date().toString();

        if (!topString) {
          topString = "(Top property was not defined.)";
        }

        if (this._topPlaceholder?.domElement) {
          this._topPlaceholder.domElement.innerHTML = ``;

          const element: React.ReactElement<IAdminPanelProps> =
            React.createElement(AdminPanel, {
              context: this.context,
              topString: topString,
            });
          ReactDOM.render(element, this._topPlaceholder.domElement);
        }
      }
    }
  }
  private async getAssociatedGroupUsers(): Promise<void> {
    try {
      const siteUrl = this.context.pageContext.site.absoluteUrl;

      const web = Web(siteUrl);

      const ownerGroup = await web.associatedOwnerGroup.users();
      const memberGroup = await web.associatedMemberGroup.users();

      this.groupUsers = [
        ...ownerGroup.map((user) => user.Email),
        ...memberGroup.map((user) => user.Email),
      ];
      this.loginUser = this.context.pageContext.legacyPageContext["userEmail"];
    } catch (error) {
      console.error("Error fetching group users:", error);
    }
  }

  public onDispose(): void {
    if (this.handleNavigation) {
      this.context.application.navigatedEvent.remove(
        this,
        this.handleNavigation
      );
    }
  }
  private _onDispose(): void {
    console.log("[AdminViewExtensionApplicationCustomizer] Disposed.");
  }
  /**
   * Checks for an existing Admin Page and creates one if it doesn't exist.
   */
  private async createAdminPage(): Promise<void> {
    try {
      const siteUrl = this.context.pageContext.site.absoluteUrl;
      const web = Web(siteUrl);
      const pageName = "Admin-Page.aspx";
      const sitePagesLib = web.lists.getByTitle("Site Pages");

      // Check if the page already exists
      const items = await sitePagesLib.items
        .filter(`FileLeafRef eq '${pageName}'`)
        .top(1)();

      if (items.length === 0) {
        // Create a new modern site page
        const page = await web.addClientsidePage(pageName, "Admin Page");

        // Add a section and text content to the page
        const section = page.addSection();
        section.addControl(new ClientsideText("Welcome to Admin Page"));

        await page.save();
        console.log("Admin Page created successfully.");
      } else {
        // console.log("Admin Page already exists.");
      }
    } catch (error) {
      console.error("Error creating or checking for the Admin Page:", error);
    }
  }
}
