import React, { useEffect, useState } from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { DialogFooter } from '@fluentui/react/lib/Dialog';
import { spfi, SPFx } from '@pnp/sp';
import Form from '@rjsf/fluent-ui';
import validator from '@rjsf/validator-ajv8';

import type { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import type { RJSFSchema } from '@rjsf/utils';

import '@pnp/sp/sites';
import '@pnp/sp/user-custom-actions';

import * as strings from 'RedirectExtensionApplicationCustomizerStrings';
import { Log } from '@microsoft/sp-core-library';

type CustomActionsProps = {
  context: ApplicationCustomizerContext;
  id: string;
  onDismiss?: () => void;
  schema: RJSFSchema;
};

export const CustomActionForm = ({
  context,
  id,
  schema,
  onDismiss,
}: CustomActionsProps): JSX.Element => {
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

    if (onDismiss) {
      onDismiss();
    }
  };

  useEffect(() => {
    fetchCustomActions().catch((e) => Log.error('RedirectExtension', e));
  }, []);

  return (
    <>
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

      <DialogFooter styles={{ actions: { clear: 'both', marginBlockStart: '3rem' } }}>
        <PrimaryButton onClick={onSave}>{strings.FormButtonSave}</PrimaryButton>
        {onDismiss && (
          <DefaultButton onClick={() => onDismiss()} type="button">
            {strings.FormButtonCancel}
          </DefaultButton>
        )}
      </DialogFooter>
    </>
  );
};

CustomActionForm.displayName = 'CustomActionForm';

export default CustomActionForm;
