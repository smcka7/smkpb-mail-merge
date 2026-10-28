/*
 * SMKPB Mail Merge settings.
 *
 * clientId: the "Application (client) ID" of the app you register in Microsoft Entra (see README.md).
 * You can paste it here before publishing, or leave it blank and paste it into the add-in's
 * Settings section the first time you open it (it's saved to your mailbox and follows you).
 */
window.MAIL_MERGE_CONFIG = {
  clientId: "",

  // Pause between emails. Exchange Online allows 30 messages a minute per mailbox;
  // 2.4 seconds keeps you at about 25 a minute.
  delayMs: 2400,

  // Largest file you can attach to every email. Microsoft Graph caps a sendMail request at 4 MB,
  // and attachments grow by a third when encoded.
  maxAttachmentBytes: 2.5 * 1024 * 1024,
};
