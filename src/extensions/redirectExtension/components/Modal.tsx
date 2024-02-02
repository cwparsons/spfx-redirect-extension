import React, { useState } from 'react';
import { Modal as FluentModal, IModalProps } from '@fluentui/react/lib/Modal';

export interface IModalChild {
  onDismiss: () => void;
}

export const Modal = ({ children, ...props }: IModalProps): JSX.Element => {
  const [hidden, setHidden] = useState(false);

  return (
    <FluentModal
      {...props}
      isOpen={!hidden}
      styles={{
        main: { padding: '2rem', minWidth: '600px', minHeight: '600px' },
        scrollableContent: { overflow: 'hidden', height: '100%' },
      }}
    >
      {React.Children.map(children, (child) =>
        React.cloneElement(child as JSX.Element, {
          onDismiss: () => setHidden(true),
        })
      )}
    </FluentModal>
  );
};

Modal.displayName = 'Modal';

export default Modal;
