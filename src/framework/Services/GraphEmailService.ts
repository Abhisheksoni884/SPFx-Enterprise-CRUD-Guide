import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IEmployeeLeaveRequest } from '../../webparts/crudOperation/models/IEmployeeLeaveRequest';
import { formatDate } from '../Utilities/Utilities';
import LoggingService from './LoggingService';

export default class GraphEmailService {
  /**
   * Sends an email notification to the manager via Microsoft Graph API v3.
   */
  public static async sendLeaveNotificationToManager(
    context: WebPartContext,
    managerEmail: string,
    managerName: string,
    request: IEmployeeLeaveRequest
  ): Promise<void> {
    if (!managerEmail) {
      console.warn('GraphEmailService: Manager email is empty. Skipping email notification.');
      return;
    }

    try {
      const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');

      const pageUrl = context.pageContext.site.absoluteUrl;

      const emailSubject = `New Leave Request Pending Approval: ${request.EmployeeName}`;
      
      const emailBodyHtml = `
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background-color: #f4f6f9; color: #333;">
          <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; border: 1px solid #e0e0e0; overflow: hidden;">
            <div style="background-color: #0078d4; color: #ffffff; padding: 20px; text-align: center;">
              <h2 style="margin: 0;">Leave Application Request</h2>
            </div>
            <div style="padding: 24px;">
              <p style="font-size: 16px;">Hello <strong>${managerName}</strong>,</p>
              <p style="font-size: 14px; color: #555;">A new leave request has been submitted by <strong>${request.EmployeeName}</strong> and requires your review.</p>
              
              <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px;">
                <tr style="background-color: #f9f9f9;">
                  <td style="padding: 10px; font-weight: bold; border-bottom: 1px solid #eee; width: 35%;">Employee Name:</td>
                  <td style="padding: 10px; border-bottom: 1px solid #eee;">${request.EmployeeName}</td>
                </tr>
                <tr>
                  <td style="padding: 10px; font-weight: bold; border-bottom: 1px solid #eee;">Leave Type:</td>
                  <td style="padding: 10px; border-bottom: 1px solid #eee;">${request.LeaveType}</td>
                </tr>
                <tr style="background-color: #f9f9f9;">
                  <td style="padding: 10px; font-weight: bold; border-bottom: 1px solid #eee;">Start Date:</td>
                  <td style="padding: 10px; border-bottom: 1px solid #eee;">${formatDate(request.StartDate)}</td>
                </tr>
                <tr>
                  <td style="padding: 10px; font-weight: bold; border-bottom: 1px solid #eee;">End Date:</td>
                  <td style="padding: 10px; border-bottom: 1px solid #eee;">${formatDate(request.EndDate)}</td>
                </tr>
                <tr style="background-color: #f9f9f9;">
                  <td style="padding: 10px; font-weight: bold; border-bottom: 1px solid #eee;">Reason:</td>
                  <td style="padding: 10px; border-bottom: 1px solid #eee;">${request.Reason || 'N/A'}</td>
                </tr>
              </table>

              <div style="text-align: center; margin-top: 30px;">
                <a href="${pageUrl}" style="background-color: #0078d4; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 4px; font-weight: bold; display: inline-block;">
                  Open Manager Dashboard
                </a>
              </div>
            </div>
            <div style="background-color: #f4f4f4; padding: 12px; text-align: center; font-size: 12px; color: #777;">
              Employee Leave Management System &bull; SharePoint Online
            </div>
          </div>
        </div>
      `;

      const mailPayload = {
        message: {
          subject: emailSubject,
          body: {
            contentType: 'HTML',
            content: emailBodyHtml
          },
          toRecipients: [
            {
              emailAddress: {
                address: managerEmail,
                name: managerName
              }
            }
          ]
        },
        saveToSentItems: 'true'
      };

      await client.api('/me/sendMail').post(mailPayload);
    } catch (error) {
      await LoggingService.logError(context, 'sendLeaveNotificationToManager', 'GraphEmailService', error);
      console.error('Failed to send email notification to manager via Graph API:', error);
    }
  }
}
