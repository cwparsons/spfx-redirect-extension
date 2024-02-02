import React, { useEffect, useState } from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { DialogFooter } from '@fluentui/react/lib/Dialog';
import { Modal } from '@fluentui/react/lib/Modal';
import { spfi, SPFx } from '@pnp/sp';
import Form from '@rjsf/fluent-ui';
import validator from '@rjsf/validator-ajv8';

import type { RJSFSchema } from '@rjsf/utils';
import type { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

import '@pnp/sp/sites';
import '@pnp/sp/user-custom-actions';

import * as strings from 'RedirectExtensionApplicationCustomizerStrings';
import { Log } from '@microsoft/sp-core-library';

type CustomActionsProps = {
  context: ApplicationCustomizerContext;
  id: string;
};

const schema: RJSFSchema = {
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

export const CustomActionForm = ({ context, id }: CustomActionsProps): JSX.Element => {
  const [hidden, setHidden] = useState(false);
  const [properties, setProperties] = useState<object>();

  const fetchCustomActions = async (): Promise<void> => {
    const sp = spfi().using(SPFx(context));
    const siteUserCustomActions = await sp.site.userCustomActions();

    const customAction = siteUserCustomActions.find((a) => a.ClientSideComponentId === id);

    if (!customAction) {
      return;
    }

    const json = JSON.parse(customAction.ClientSideComponentProperties);

    setProperties(json);
  };

  const onSave = async (): Promise<void> => {
    const sp = spfi().using(SPFx(context));
    const siteUserCustomActions = await sp.site.userCustomActions();

    const customAction = siteUserCustomActions.find((a) => a.ClientSideComponentId === id);

    if (!customAction) {
      return;
    }

    const stringifiedProperties = JSON.stringify(properties);

    await sp.site.userCustomActions.getById(customAction?.Id).update({
      ClientSideComponentProperties: stringifiedProperties,
    });

    setHidden(true);
  };

  useEffect(() => {
    fetchCustomActions().catch((e) => Log.error('RedirectExtension', e));
  }, []);

  return (
    <Modal
      isOpen={!hidden}
      styles={{
        main: { padding: '2rem', minWidth: '600px', minHeight: '600px' },
        scrollableContent: { overflowX: 'hidden', height: '100%' },
      }}
    >
      <Form
        schema={schema}
        formData={properties}
        onChange={(e) => setProperties(e.formData)}
        validator={validator}
        uiSchema={{
          'ui:submitButtonOptions': {
            norender: true,
          },
        }}
      />

      <DialogFooter styles={{ actions: { clear: 'both', marginBlockStart: '1rem' } }}>
        <PrimaryButton onClick={onSave}>{strings.FormButtonSave}</PrimaryButton>
        <DefaultButton onClick={() => setHidden(true)} type="button">
          {strings.FormButtonCancel}
        </DefaultButton>
      </DialogFooter>
    </Modal>
  );
};

CustomActionForm.displayName = 'CustomActionForm';

export default CustomActionForm;
