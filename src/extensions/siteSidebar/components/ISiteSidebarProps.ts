import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

export interface ISiteSidebarProps {
  context: ApplicationCustomizerContext;
  isOpen: boolean;
  onDismiss: () => void;
}
