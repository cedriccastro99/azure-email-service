export type AzureEmailAttachment = {
  '@odata.type': '#microsoft.graph.fileAttachment',
  name: string,
  contentType: string,
  contentBytes: string
}

export interface TAzureSendMail {
  senderName?: string;
  body: string;
  fromEmail: string;
  subject?: string;
  toEmail: string[];
  attachments?: AzureEmailAttachment[];
  cc?: string[];
  bcc?: string[];
  replyTo?: string;
}

export interface AzureEmailServiceConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  maxRetries?: number;
}