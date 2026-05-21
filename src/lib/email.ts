import { graphClient } from './sharepoint/client';

/**
 * Send mail via Microsoft Graph as the signed-in user.
 *
 * Uses the `Mail.Send` delegated permission (added to the app's scopes in
 * `auth/msalConfig.ts` and `sharepoint/client.ts`). The message lands in the
 * sender's actual Outlook Sent Items folder.
 *
 * Save-to-sent is enabled by default so the user has a copy in their own
 * mailbox — useful for any reply that comes back outside the app.
 */

export interface EmailRecipient {
  address: string;
  name?: string;
}

export interface EmailAttachment {
  filename: string;
  contentType: string;
  /** Plain base64 (no data: prefix). */
  contentBase64: string;
}

export interface SendEmailOptions {
  to: EmailRecipient[];
  cc?: EmailRecipient[];
  bcc?: EmailRecipient[];
  subject: string;
  /** Plain-text body. HTML is escaped on display in Outlook for safety. */
  bodyText: string;
  /** Optional small attachments (inline base64; cap ~3 MB total per Graph API). */
  attachments?: EmailAttachment[];
}

/**
 * Send a plain-text email. Throws with a useful message on failure (consent
 * not granted, network error, recipient validation, etc.) so the caller can
 * surface it in the compose modal.
 */
export async function sendEmail(options: SendEmailOptions): Promise<void> {
  const { to, cc, bcc, subject, bodyText, attachments } = options;

  if (to.length === 0) {
    throw new Error('At least one recipient is required.');
  }
  if (!subject.trim()) {
    throw new Error('Subject is required.');
  }

  const message: Record<string, unknown> = {
    subject,
    body: {
      contentType: 'Text',
      content: bodyText,
    },
    toRecipients: to.map(toGraphRecipient),
  };
  if (cc && cc.length > 0) message.ccRecipients = cc.map(toGraphRecipient);
  if (bcc && bcc.length > 0) message.bccRecipients = bcc.map(toGraphRecipient);

  if (attachments && attachments.length > 0) {
    message.attachments = attachments.map((a) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.filename,
      contentType: a.contentType,
      contentBytes: a.contentBase64,
    }));
  }

  try {
    await graphClient.api('/me/sendMail').post({
      message,
      saveToSentItems: true,
    });
  } catch (err) {
    // Surface a friendlier error for the common consent-missing case
    const raw = err instanceof Error ? err.message : String(err);
    if (raw.includes('Mail.Send') || raw.toLowerCase().includes('insufficient privileges') || raw.includes('AADSTS65001')) {
      throw new Error(
        "This app doesn't have permission to send mail yet. An Azure AD admin needs to grant the Mail.Send permission, then sign out and back in. (Original error: " + raw + ")",
      );
    }
    throw new Error(raw);
  }
}

function toGraphRecipient(r: EmailRecipient): Record<string, unknown> {
  return {
    emailAddress: r.name
      ? { address: r.address, name: r.name }
      : { address: r.address },
  };
}

// =============================================================================
// Variable substitution for templates
//
// Supports {{contact}}, {{contact_email}}, {{property}}, {{properties}},
// {{owner}}, {{date}}, {{user}}, {{user_email}}. Unknown variables pass
// through unchanged so the user notices and can fix the template.
// =============================================================================

export interface TemplateContext {
  contactName?: string;
  contactEmail?: string;
  /** Primary property name (first one when multiple are linked). */
  propertyName?: string;
  /** Comma-joined list when more than one property is linked. */
  propertiesList?: string;
  ownerName?: string;
  userName?: string;
  userEmail?: string;
}

/**
 * Replace {{tokens}} in a template string with values from `ctx`.
 * Whitespace inside the braces is tolerated: `{{ contact }}` works too.
 */
export function applyTemplateVars(text: string, ctx: TemplateContext): string {
  const today = new Date().toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric',
  });
  const map: Record<string, string | undefined> = {
    contact: ctx.contactName,
    contact_email: ctx.contactEmail,
    property: ctx.propertyName,
    properties: ctx.propertiesList,
    owner: ctx.ownerName,
    date: today,
    user: ctx.userName,
    user_email: ctx.userEmail,
  };
  return text.replace(/\{\{\s*([a-z_]+)\s*\}\}/gi, (match, key: string) => {
    const value = map[key.toLowerCase()];
    return value == null ? match : value;
  });
}
