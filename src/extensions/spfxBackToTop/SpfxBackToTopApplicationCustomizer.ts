import * as React from "react";
import * as ReactDOM from "react-dom";

import { override } from "@microsoft/decorators";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import IBackToTopProps from "./BackToTop/IBackToTopProps";

import * as strings from "SpfxBackToTopApplicationCustomizerStrings";
import BackToTop from "./BackToTop/BackToTop";
import { SPEventArgs } from "@microsoft/sp-core-library";

export interface ISpfxBackToTopApplicationCustomizerProperties {}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxBackToTopApplicationCustomizer extends BaseApplicationCustomizer<ISpfxBackToTopApplicationCustomizerProperties> {
  private topPlaceholder: PlaceholderContent | undefined;

  private renderPlaceHolders(): void {
    if (!this.topPlaceholder) {
      this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top
      );
      this._renderControls(0);
    }
  }

  private _renderControls = (delay: number) => {
    // The event is getting called before the page navigation happens
    // Due to this, the onScroll event that we are adding (for BackToTop)
    // is getting reset when the partial page load is finished
    // There is an open issue regarding this - https://github.com/SharePoint/sp-dev-docs/issues/5321
    // So, I have added a 3 seconds delay for the child components to load.
    // This is a temporary fix untill the actual issue in SPFx is resolved.
    setTimeout(() => {
      if (this.topPlaceholder) {
        if (this.topPlaceholder.domElement) {
          console.log("back to top")
          const element: React.ReactElement<IBackToTopProps> =
            React.createElement(BackToTop, {
              currentUrl: this.context.pageContext.web.absoluteUrl,
            });
          ReactDOM.render(element, this.topPlaceholder.domElement);
        }
      } else {
        this.renderPlaceHolders();
      }
    }, delay);
  };

  @override
  public onDispose(): Promise<void> {
    this.context.placeholderProvider.changedEvent.remove(
      this,
      this.renderPlaceHolders
    );
    return Promise.resolve();
  }

  private navigatedEventHandler(args: SPEventArgs): void {
    this._renderControls(3000);
  }

  @override
  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(
      this,
      this.renderPlaceHolders
    );
    this.context.application.navigatedEvent.add(
      this,
      this.navigatedEventHandler
    );
    return Promise.resolve();
  }
}
