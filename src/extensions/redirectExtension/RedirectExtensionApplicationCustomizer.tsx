import React, { Suspense, lazy } from 'react';
import ReactDOM from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from '@microsoft/sp-application-base';

import type { RJSFSchema } from '@rjsf/utils';

import * as strings from 'RedirectExtensionApplicationCustomizerStrings';

const LOG_SOURCE: string = 'RedirectExtensionApplicationCustomizer';

const Modal = lazy(
  () => import(/* webpackChunkName: 'redirectextension-modal' */ './components/Modal')
);
const CustomActionForm = lazy(
  () =>
    import(
      /* webpackChunkName: 'redirectextension-customactionform' */ './components/CustomActionForm'
    )
);
const Dialog = lazy(
  () => import(/* webpackChunkName: 'redirectextension-dialog' */ './components/Dialog')
);

export type RedirectExtensionApplicationCustomizerProperties = {
  title: string;
  message: string;
  button: string;
  rules: {
    source: string;
    destination: string;
  }[];
};

export const RedirectExtensionApplicationCustomizerPropertiesSchema: RJSFSchema = {
  title: strings.FormHeading,
  type: 'object',
  properties: {
    title: {
      type: 'string',
      title: strings.FormLabelDialogTitle,
    },
    message: {
      type: 'string',
      title: strings.FormLabelMessage,
    },
    button: {
      type: 'string',
      title: strings.FormLabelButton,
    },
    rules: {
      type: 'array',
      title: strings.FormLabelRules,
      items: {
        type: 'object',
        properties: {
          source: {
            type: 'string',
            title: strings.FormLabelSourceURL,
          },
          destination: {
            type: 'string',
            title: strings.FormLabelDestinationURL,
          },
        },
      },
    },
  },
};

export default class RedirectExtensionApplicationCustomizer extends BaseApplicationCustomizer<RedirectExtensionApplicationCustomizerProperties> {
  private _topPlaceholder?: PlaceholderContent;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._render);

    return Promise.resolve();
  }

  private _onDispose(): void {
    Log.info(LOG_SOURCE, 'Disposed custom top placeholders.');
  }

  private _getDestination(): string | undefined {
    const rule = this.properties.rules.find((r) => window.location.href.match(r.source));

    if (!rule) {
      return undefined;
    }

    let destination = rule.destination;

    // Perform the replacement using regular expressions
    window.location.href.replace(new RegExp(rule.source), function (...args) {
      // Get capture groups from the source regex
      const sourceGroups = args.slice(1, -2);

      // Replace capture groups in the destination regex
      let replacedDest = destination;

      sourceGroups.forEach(function (group, index) {
        replacedDest = replacedDest.replace(new RegExp('\\$' + (index + 1), 'g'), group);
      });

      // Return the replaced destination
      destination = replacedDest;

      return destination;
    });

    return destination;
  }

  private _render(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
    }

    const container = this._topPlaceholder?.domElement;

    if (!container) {
      return;
    }

    if (window.location.search.includes('redirectconfig=true')) {
      ReactDOM.render(
        <Suspense fallback={<></>}>
          <Modal>
            <CustomActionForm
              schema={RedirectExtensionApplicationCustomizerPropertiesSchema}
              context={this.context}
              id={this.componentId}
            />
          </Modal>
        </Suspense>,
        container
      );

      return;
    }

    ReactDOM.render(
      <Suspense fallback={<></>}>
        <Dialog
          button={this.properties.button}
          href={this._getDestination()}
          subText={this.properties.message}
          title={this.properties.title}
        />
      </Suspense>,
      container
    );
  }
}
