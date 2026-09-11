import { BaseComponentContext } from '@microsoft/sp-component-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import '@pnp/sp/attachments';
import '@pnp/sp/site-users/web';

let _sp: SPFI | undefined = undefined;

/**
 * Returns a singleton instance of SPFI configured for SPFx context.
 *
 * @param context - BaseComponentContext optional if already initialized
 */
export const getSP = (context?: BaseComponentContext): SPFI => {
  if (_sp === undefined) {
    if (!context) {
      throw new Error('SPFx Context must be provided on initial getSP call.');
    }
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

