import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export const sendEmail = async (
  spHttpClient: SPHttpClient,
  from: string,
  to: string[],
  subject: string,
  body: string,
  webUrl: string
): Promise<string> => {
  const requestUrl = `${webUrl}/_api/SP.Utilities.Utility.SendEmail`;

  const emailProperties = {
    __metadata: { type: 'SP.Utilities.EmailProperties' },
    From: from,
    To: { results: to },
    Body: body,
    Subject: subject,
  };

  const requestHeaders = {
    Accept: 'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose',
    'X-RequestDigest': document.getElementById('__REQUESTDIGEST')?.getAttribute('value') || '',
  };

  const options = {
    method: 'POST',
    headers: requestHeaders,
    body: JSON.stringify({ properties: emailProperties }),
  };

  try {
    // Using SPHttpClient to send email through SharePoint API
    const response: SPHttpClientResponse = await spHttpClient.fetch(requestUrl, SPHttpClient.configurations.v1, options);

    if (response.ok) {
      return 'Email sent successfully!';
    } else {
      return `Failed to send email. Status: ${response.statusText}`;
    }
  } catch (error) {
    console.error('Error sending email:', error);
    return 'An error occurred while sending the email.';
  }
};
