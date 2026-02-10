# Azure Email Service

A simple and lightweight TypeScript/JavaScript wrapper for sending emails via Microsoft Graph API using Azure AD authentication.

## Installation

```bash
npm install @cedric-castro/azure-email-service
```

## Prerequisites

You need to have:
1. An Azure AD application with client credentials
2. Microsoft Graph API permissions: `Mail.Send`
3. A user account with a mailbox in your tenant

## Usage

### TypeScript

```typescript
import { AzureEmailService } from '@cedric-castro/azure-email-service';

const emailService = new AzureEmailService({
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  tenantId: 'your-tenant-id',
  maxRetries: 3 // optional, default is 2
});

// Send email
await emailService.sendEmail({
  fromEmail: 'sender@yourdomain.com',
  toEmail: ['recipient@example.com'],
  subject: 'Hello from Azure',
  body: '<h1>Hello!</h1><p>This is a test email.</p>',
  senderName: 'John Doe', // optional
  cc: ['cc@example.com'], // optional
  bcc: ['bcc@example.com'], // optional
  replyTo: 'reply@example.com', // optional
  attachments: [] // optional
});
```

### JavaScript

```javascript
const { AzureEmailService } = require('@cedric-castro/azure-email-service');

const emailService = new AzureEmailService({
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  tenantId: 'your-tenant-id'
});

emailService.sendEmail({
  fromEmail: 'sender@yourdomain.com',
  toEmail: ['recipient@example.com'],
  subject: 'Hello',
  body: '<p>Test email</p>'
})
.then(() => console.log('Email sent!'))
.catch(err => console.error('Error:', err));
```

## API

### Constructor

```typescript
new AzureEmailService(config: AzureEmailServiceConfig)
```

**Config options:**
| Parameter | Type | Required | Description |
| :--- | :--- | :---: | :--- |
| **`clientId`** | `string` | ✅ Yes | Azure AD application client ID |
| **`clientSecret`** | `string` | ✅ Yes | Azure AD application client secret |
| **`tenantId`** | `string` | ✅ Yes | Azure AD tenant ID |
| **`maxRetries`** | `number` | ❌ No | Maximum retry attempts (Default: **2**) |

### sendEmail

```typescript
sendEmail(params: TAzureSendMail): Promise<boolean>
```

**Parameters:**
| Parameter | Type | Required | Description |
| :--- | :--- | :---: | :--- |
| **`fromEmail`** | `string` | ✅ Yes | Sender email address |
| **`toEmail`** | `string[]` | ✅ Yes | Array of recipient email addresses |
| **`body`** | `string` | ✅ Yes | Email body in HTML format |
| **`subject`** | `string` | ❌ No | Email subject |
| **`senderName`** | `string` | ❌ No | Display name for sender |
| **`cc`** | `string[]` | ❌ No | CC recipients |
| **`bcc`** | `string[]` | ❌ No | BCC recipients |
| **`replyTo`** | `string` | ❌ No | Reply-to address |
| **`attachments`** | `AzureEmailAttachment[]` | ❌ No | Email attachments |

**Returns:** `Promise<boolean>` - Returns `true` if email sent successfully

**Throws:** Error if email fails to send after all retry attempts

## Azure Setup

1. Register an application in Azure AD
2. Add API permissions: `Microsoft Graph > Application permissions > Mail.Send`
3. Grant admin consent for the permissions
4. Create a client secret
5. Note down: Client ID, Client Secret, and Tenant ID

## Error Handling

```typescript
try {
  await emailService.sendEmail({
    fromEmail: 'sender@domain.com',
    toEmail: ['recipient@example.com'],
    subject: 'Test',
    body: '<p>Hello</p>'
  });
  console.log('Success!');
} catch (error) {
  console.error('Failed to send email:', error.message);
}
```

## Features

| Feature | Description | Status |
| :--- | :--- | :---: |
| **TypeScript Support** | Includes full type definitions for robust development. | ✅ |
| **Automatic Retry** | Built-in mechanism with **exponential backoff** for reliability. | ✅ |
| **HTML Support** | Capability to send rich-text emails using HTML bodies. | ✅ |
| **CC & BCC** | Support for Carbon Copy and Blind Carbon Copy recipients. | ✅ |
| **Attachments** | Support for sending files via `AzureEmailAttachment`. | ✅ |
| **Custom Sender** | Option to define a specific display name for the "From" field. | ✅ |
| **Reply-to** | Ability to specify a separate address for user responses. | ✅ |
| **Promise-based** | Modern `async/await` API for clean asynchronous flow. | ✅ |

## License

MIT

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## Support

For issues and questions, please open an issue on [GitHub](https://github.com/cedriccastro99/azure-email-service/issues).

