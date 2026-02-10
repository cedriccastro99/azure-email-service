import { Client } from "@microsoft/microsoft-graph-client";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { TAzureSendMail, AzureEmailServiceConfig } from "./types";

export class AzureEmailService {
  private clientApp: ConfidentialClientApplication;
  private maxRetries: number;

  constructor(config: AzureEmailServiceConfig) {
    const { clientId, clientSecret, tenantId, maxRetries = 2 } = config;

    this.clientApp = new ConfidentialClientApplication({
      auth: {
        clientId,
        clientSecret,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
    });

    this.maxRetries = maxRetries;
    console.log("AzureEmailService initialized");
  }

  private async getAccessToken(): Promise<string> {
    try {
      const clientCredentialRequest = {
        scopes: ["https://graph.microsoft.com/.default"],
      };

      const response = await this.clientApp.acquireTokenByClientCredential(
        clientCredentialRequest,
      );

      if (!response?.accessToken) {
        throw new Error("Failed to acquire access token");
      }

      return response.accessToken;
    } catch (error) {
      console.error("Error acquiring access token:", error);
      throw error;
    }
  }

  async sendEmail(params: TAzureSendMail): Promise<boolean> {
    const {
      senderName = "",
      body,
      fromEmail,
      subject = "Email from Azure Email Service",
      toEmail = [],
      attachments = [],
      cc = [],
      bcc = [],
      replyTo = "",
    } = params;

    // Validate required fields
    if (!fromEmail) {
      throw new Error("fromEmail is required");
    }

    if (!toEmail || toEmail.length === 0) {
      throw new Error("At least one recipient email is required");
    }

    if (!body) {
      throw new Error("Email body is required");
    }

    let lastError: Error | undefined;

    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      try {
        const accessToken = await this.getAccessToken();
        const graphClient = Client.init({
          authProvider: (done) => {
            done(null, accessToken);
          },
        });

        const mail = {
          message: {
            subject: subject,
            body: {
              contentType: "Html",
              content: body,
            },
            toRecipients: toEmail.map((email) => ({
              emailAddress: { address: email },
            })),
            attachments: attachments,
            ccRecipients: cc.length
              ? cc.map((email) => ({ emailAddress: { address: email } }))
              : [],
            bccRecipients: bcc.length
              ? bcc.map((email) => ({ emailAddress: { address: email } }))
              : [],
            replyTo: replyTo ? [{ emailAddress: { address: replyTo } }] : [],
            ...(senderName
              ? {
                  from: {
                    emailAddress: {
                      name: senderName,
                      address: fromEmail,
                    },
                  },
                }
              : {}),
          },
        };

        await graphClient.api(`/users/${fromEmail}/sendMail`).post(mail);

        for (const recipient of toEmail) {
          console.log(`Email sent to ${recipient}`);
        }

        console.log("Email sent successfully.");
        return true;
      } catch (error: any) {
        lastError = error;
        console.log(`Attempt ${attempt + 1} failed: ${error.message}`);

        if (attempt < this.maxRetries) {
          await this.delay(1000 * (attempt + 1));
          continue;
        }

        break;
      }
    }

    throw lastError;
  }

  private async delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
