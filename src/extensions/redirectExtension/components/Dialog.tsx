import React from 'react';

import { Dialog as FluentDialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { useBoolean } from '@fluentui/react-hooks';

type DialogProps = {
  href?: string;
  subText?: string;
  title?: string;
  button: string;
};

const modalPropsStyles = { main: { maxWidth: 450 } };
const dialogContentProps = {
  type: DialogType.normal,
};

export const Dialog = ({ button, href, subText, title }: DialogProps): JSX.Element => {
  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(false);

  return (
    <FluentDialog
      dialogContentProps={{
        ...dialogContentProps,
        subText,
        title,
      }}
      hidden={hideDialog}
      onDismiss={toggleHideDialog}
      modalProps={{
        isBlocking: true,
        styles: modalPropsStyles,
      }}
    >
      <DialogFooter>
        {href && (
          <PrimaryButton
            data-interception="off"
            href={href}
            onClick={() => toggleHideDialog()}
            text={button}
          />
        )}
      </DialogFooter>
    </FluentDialog>
  );
};

export default Dialog;
