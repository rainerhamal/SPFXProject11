import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  //! Lsn 5.3.8 Update the header and footer placeholders: Update the list of imports to add the following references: PlaceholderContent and PlaceholderName.
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HelloAppCustomizerApplicationCustomizerStrings';

//! Lsn 5.3.7 Add the following import statements to the top of the file after the existing import statements:
import styles from './HelloAppCustomizerApplicationCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

const LOG_SOURCE: string = 'HelloAppCustomizerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
//! Lsn 5.3.2 Update application customizer to add placeholders to the page. Update the customizer to have two public settable properties. Edit this to have only two properties:
export interface IHelloAppCustomizerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  //? testMessage: string;
  header: string;
  footer: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HelloAppCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloAppCustomizerApplicationCustomizerProperties> {

  //! Lsn 5.3.8 add the following two private members:
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  //! Lsn 5.3.9 This method is used when the placeholders are disposed.
  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }

  //! Lsn 5.3.10 This method will be called when the placeholders are rendered:
  private _renderPlaceHolders(): void {
    console.log('Available application customizer placeholders: ',
      this.context.placeholderProvider.placeholderNames.map((name) => PlaceholderName[name]).join(', '));

      //! Lsn 5.3.11 This code will obtain a handle to the top placeholder on the page. It will then add some markup to the placeholder using the message defined in the public property:
      if (!this._topPlaceholder) {
        this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top, 
          {onDispose:this._onDispose}
        );
        if (!this._topPlaceholder) {
          console.error('The expected placeholder (Top) was not found.');
          return;
        }
        if (this.properties) {
          let headerMessage: string = this.properties.header;
          if (!headerMessage) {
            headerMessage = '(header property was not defined.)';
          }
          if (this._topPlaceholder.domElement) {
            this._topPlaceholder.domElement.innerHTML = `
              <div class="${styles.app}">
                <div class="${styles.top}">
                  <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(headerMessage)}
                </div>
              </div>`;
          }
        }
      }
      //! Lsn 5.3.12 Add the following code to the _renderPlaceHolders() to update the bottom placeholder: 
      if (!this._bottomPlaceholder) {
        this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose }
        );
      
        if (!this._bottomPlaceholder) {
          console.error('The expected placeholder (Bottom) was not found.');
          return;
        }
      
        if (this.properties) {
          let footerMessage: string = this.properties.footer;
          if (!footerMessage) {
            footerMessage = '(footer property was not defined.)';
          }
      
          if (this._bottomPlaceholder.domElement) {
            this._bottomPlaceholder.domElement.innerHTML = `
              <div class="${styles.app}">
                <div class="${styles.bottom}">
                  <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(footerMessage)}
                </div>
              </div>`;
          }
        }
      }
  }

  
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    //! Lsn 5.3.13 Replace all the code in the onInit() method with the following code:

    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    return Promise.resolve();

    // let message: string = this.properties.testMessage;
    // if (!message) {
    //   message = '(No properties were provided.)';
    // }

    // Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);
  }
}
