import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from './pnpjsConfig';
import { LIST_NAMES } from '../Constants/Constant';

export default class LoggingService {
  /**
   * Logs an application error to the SharePoint LogHistory list with fallback to browser console.
   */
  public static async logError(
    context: WebPartContext,
    functionName: string,
    pageName: string,
    error: unknown
  ): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const stackTrace = error instanceof Error && error.stack ? error.stack : '';

    console.error(`[${pageName} -> ${functionName}]`, errorMessage, error);

    try {
      const sp = getSP(context);
      await sp.web.lists.getByTitle(LIST_NAMES.LOG_HISTORY).items.add({
        Title: `${pageName} - ${functionName}`,
        FunctionName: functionName,
        PageName: pageName,
        ErrorMessage: errorMessage,
        StackTrace: stackTrace
      });
    } catch (loggingError) {
      // Fallback if LogHistory list does not exist in the tenant
      console.warn('Logging to LogHistory list failed or list does not exist:', loggingError);
    }
  }
}
