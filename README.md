Gmail Exit – Exporting Email Addresses and Communication Data

A practical guide for a controlled transition away from Gmail.

This project demonstrates how to extract email addresses from received and sent Gmail messages and organize them into a structured contact list.

The approach is intended for users who want to reduce their long-term dependency on Gmail and need an overview of who they have actually communicated with before migrating to another email provider.
Why Leave Gmail?

There is no single reason to move away from Gmail. Gmail is a mature, reliable, and feature-rich email service.

However, there are legitimate reasons why individuals and organizations may want to reduce their dependency on a large centralized provider.
1. Data Ownership

Email is one of the most important sources of personal and professional communication data.

When communication is hosted entirely by a single cloud provider, users become dependent on that provider's infrastructure, policies, terms of service, and future product decisions.

Using an independent provider or a self-managed setup can provide greater control over where and how communication data is stored.
2. Vendor Lock-In

A Gmail account is often connected to much more than email:

    Email
    Contacts
    Calendar
    Google Drive
    Google Photos
    YouTube
    Google Account authentication
    Third-party services using Google Sign-In

Over time, this can create significant dependency on a single ecosystem.

A planned migration can reduce that dependency.
3. Privacy and Metadata

Email privacy is not only about message content.

Communication can also generate metadata such as:

    sender
    recipient
    timestamp
    communication frequency
    subject
    technical delivery information
    relationships between communication partners

Even when message contents are protected during transmission, metadata can still reveal valuable information about communication patterns.
4. Long-Term Portability

A useful security principle is:

    Your data should remain portable.

Regularly exporting important data makes it easier to migrate to another provider in the future.

This is particularly important for email because changing an email address can affect dozens or hundreds of external services.
5. Account Dependency and Single Points of Failure

A single Google Account can become a central authentication point for many services.

If access to that account is lost, multiple connected services may be affected simultaneously.

Maintaining independent recovery options and backups reduces this risk.
Project Goal

The main question this project helps answer is:

    Which email addresses have I actually communicated with?

The resulting spreadsheet can contain information such as:
Name	Email Address	Received	Sent	Total	Type
Max Mustermann	max@example.com	42	8	50	Received + Sent
Example Company	info@example.com	27	0	27	Received Only
Anna Miller	anna@example.com	0	12	12	Sent Only

This provides a practical contact inventory that can be used during an email migration.
Security Model

The script is designed to process the data within the user's Google environment.

It does not intentionally transmit passwords, Gmail credentials, or exported contacts to an external server.

However, the script requires authorization to access Gmail.

Therefore:

    Only run code that you have reviewed and understand.

Never authorize an unknown or untrusted Apps Script project simply because it claims to provide an export function.

A script should never ask you for:

    your Gmail password
    your 2FA code
    recovery codes
    browser cookies
    SSH keys
    API keys unrelated to the task

Be especially careful with scripts containing code that uploads data to external services.
Requirements

You need:

    a Gmail account
    access to Google Apps Script
    Google Sheets
    a modern web browser

The script is designed to work with several thousand messages.

The example is configured for a maximum of 6,500 messages per export. This limit can be changed in the configuration section.
Installation
1. Open Google Apps Script

Open:

https://script.google.com/

Create a new Apps Script project.
2. Replace the Default Code

Delete the automatically generated example code.

Create or use the main .gs file and paste the script below.
Script

const MAX_MESSAGES = 6500;
const THREAD_BATCH = 500;
const TIME_LIMIT_MS = 270000;


/*
 * START THE EXPORT
 */
function startExport() {

  deleteProcessTriggers();

  const props =
    PropertiesService.getScriptProperties();

  props.deleteAllProperties();

  const spreadsheet =
    SpreadsheetApp.create(
      "Gmail Email Addresses Export"
    );

  const sheet =
    spreadsheet.getActiveSheet();

  sheet.setName("Addresses");

  sheet.getRange(1, 1, 1, 6)
    .setValues([[
      "Name",
      "Email Address",
      "Received",
      "Sent",
      "Total",
      "Type"
    ]]);

  sheet.setFrozenRows(1);

  sheet.getRange("H1")
    .setValue("STATUS");

  sheet.getRange("H2")
    .setValue("0 messages");

  sheet.getRange("H3")
    .setValue("0 addresses");

  sheet.getRange("H4")
    .setValue("Started...");

  props.setProperty(
    "SPREADSHEET_ID",
    spreadsheet.getId()
  );

  props.setProperty(
    "THREAD_OFFSET",
    "0"
  );

  props.setProperty(
    "PROCESSED",
    "0"
  );

  Logger.log(
    "Spreadsheet: " +
    spreadsheet.getUrl()
  );

  processEmails();
}


/*
 * MAIN PROCESSING FUNCTION
 */
function processEmails() {

  const startTime = Date.now();

  const props =
    PropertiesService.getScriptProperties();

  const spreadsheetId =
    props.getProperty("SPREADSHEET_ID");

  const sheet =
    SpreadsheetApp
      .openById(spreadsheetId)
      .getSheetByName("Addresses");

  let threadOffset =
    Number(
      props.getProperty(
        "THREAD_OFFSET"
      ) || 0
    );

  let processed =
    Number(
      props.getProperty(
        "PROCESSED"
      ) || 0
    );

  const ownEmail =
    Session.getEffectiveUser()
      .getEmail()
      .toLowerCase();

  const contacts =
    loadContacts(sheet);

  let finished = false;

  while (
    Date.now() - startTime <
      TIME_LIMIT_MS &&
    processed < MAX_MESSAGES
  ) {

    const threads =
      GmailApp.search(
        "in:anywhere",
        threadOffset,
        THREAD_BATCH
      );

    if (
      threads.length === 0
    ) {

      finished = true;
      break;
    }

    for (
      const thread of threads
    ) {

      const messages =
        thread.getMessages();

      for (
        const message of messages
      ) {

        if (
          processed >=
          MAX_MESSAGES
        ) {
          finished = true;
          break;
        }

        processMessage(
          message,
          contacts,
          ownEmail
        );

        processed++;

        if (
          Date.now() - startTime >=
          TIME_LIMIT_MS
        ) {
          break;
        }
      }

      if (
        finished ||
        Date.now() - startTime >=
          TIME_LIMIT_MS
      ) {
        break;
      }
    }

    threadOffset +=
      threads.length;

    if (
      threads.length <
      THREAD_BATCH
    ) {
      finished = true;
      break;
    }
  }

  props.setProperty(
    "THREAD_OFFSET",
    String(threadOffset)
  );

  props.setProperty(
    "PROCESSED",
    String(processed)
  );

  updateSheet(
    sheet,
    contacts
  );

  const addressCount =
    Object.keys(contacts).length;

  sheet.getRange("H2")
    .setValue(
      processed +
      " messages processed"
    );

  sheet.getRange("H3")
    .setValue(
      addressCount +
      " unique addresses"
    );

  sheet.getRange("H4")
    .setValue(
      Math.round(
        processed /
        MAX_MESSAGES *
        100
      ) +
      "% of maximum " +
      MAX_MESSAGES
    );

  SpreadsheetApp.flush();

  Logger.log(
    "Progress: " +
    processed +
    " messages"
  );

  Logger.log(
    "Unique addresses: " +
    addressCount
  );

  if (
    !finished &&
    processed < MAX_MESSAGES
  ) {

    deleteProcessTriggers();

    ScriptApp.newTrigger(
      "processEmails"
    )
    .timeBased()
    .after(3000)
    .create();

    sheet.getRange("H5")
      .setValue(
        "Next processing run scheduled..."
      );

  } else {

    finishExport(
      sheet,
      contacts
    );
  }
}


/*
 * LOAD EXISTING CONTACTS
 */
function loadContacts(sheet) {

  const contacts = {};

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    return contacts;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();

  values.forEach(row => {

    const name =
      String(row[0] || "");

    const email =
      String(row[1] || "")
        .toLowerCase()
        .trim();

    if (!email) {
      return;
    }

    contacts[email] = {

      name: name,

      received:
        Number(row[2] || 0),

      sent:
        Number(row[3] || 0)
    };
  });

  return contacts;
}


/*
 * PROCESS ONE MESSAGE
 */
function processMessage(
  message,
  contacts,
  ownEmail
) {

  const fromText =
    message.getFrom() || "";

  const toText =
    message.getTo() || "";

  const ccText =
    message.getCc() || "";

  const bccText =
    message.getBcc() || "";

  const from =
    extractEmails(fromText);

  const to =
    extractEmails(toText);

  const cc =
    extractEmails(ccText);

  const bcc =
    extractEmails(bccText);


  /*
   * RECEIVED
   */
  from.forEach(email => {

    if (
      email === ownEmail
    ) {
      return;
    }

    if (
      !contacts[email]
    ) {

      contacts[email] = {

        name:
          extractName(
            fromText,
            email
          ),

        received: 0,

        sent: 0
      };
    }

    contacts[email]
      .received++;
  });


  /*
   * SENT
   */
  const sent =
    from.some(
      email =>
        email === ownEmail
    );

  if (sent) {

    [
      ...to,
      ...cc,
      ...bcc
    ].forEach(email => {

      if (
        email === ownEmail
      ) {
        return;
      }

      if (
        !contacts[email]
      ) {

        contacts[email] = {

          name:
            extractName(
              toText +
              " " +
              ccText +
              " " +
              bccText,
              email
            ),

          received: 0,

          sent: 0
        };
      }

      contacts[email]
        .sent++;
    });
  }
}


/*
 * EXTRACT EMAIL ADDRESSES
 */
function extractEmails(text) {

  if (!text) {
    return [];
  }

  const matches =
    text.match(
      /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g
    );

  if (!matches) {
    return [];
  }

  return [
    ...new Set(
      matches.map(
        email =>
          email.toLowerCase()
      )
    )
  ];
}


/*
 * EXTRACT DISPLAY NAME
 */
function extractName(
  text,
  email
) {

  if (!text) {
    return "";
  }

  const escaped =
    email.replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&"
    );

  const regex =
    new RegExp(
      '([^<,]+?)\\s*<\\s*' +
      escaped +
      '\\s*>',
      "i"
    );

  const match =
    text.match(regex);

  if (match) {

    return match[1]
      .trim()
      .replace(
        /^["']|["']$/g,
        ""
      );
  }

  return "";
}


/*
 * UPDATE SPREADSHEET
 */
function updateSheet(
  sheet,
  contacts
) {

  if (
    sheet.getLastRow() > 1
  ) {

    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        6
      )
      .clearContent();
  }

  const rows = [];

  Object.keys(contacts)
    .sort()
    .forEach(email => {

      const contact =
        contacts[email];

      let type;

      if (
        contact.received > 0 &&
        contact.sent > 0
      ) {

        type =
          "Received + Sent";

      } else if (
        contact.received > 0
      ) {

        type =
          "Received Only";

      } else {

        type =
          "Sent Only";
      }

      rows.push([

        contact.name || "",

        email,

        contact.received,

        contact.sent,

        contact.received +
          contact.sent,

        type
      ]);
    });

  if (
    rows.length > 0
  ) {

    sheet
      .getRange(
        2,
        1,
        rows.length,
        6
      )
      .setValues(rows);
  }

  sheet.autoResizeColumns(
    1,
    6
  );
}


/*
 * FINISH EXPORT
 */
function finishExport(
  sheet,
  contacts
) {

  deleteProcessTriggers();

  const count =
    Object.keys(contacts).length;

  sheet.getRange("H2")
    .setValue("COMPLETE");

  sheet.getRange("H3")
    .setValue(
      count +
      " unique email addresses"
    );

  sheet.getRange("H4")
    .setValue(
      "Export completed"
    );

  sheet.getRange("H5")
    .setValue(
      "Messages: " +
      PropertiesService
        .getScriptProperties()
        .getProperty(
          "PROCESSED"
        )
    );

  SpreadsheetApp.flush();

  Logger.log(
    "EXPORT COMPLETE"
  );

  Logger.log(
    "Messages: " +
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "PROCESSED"
      )
  );

  Logger.log(
    "Unique addresses: " +
    count
  );
}


/*
 * DELETE AUTOMATIC TRIGGERS
 */
function deleteProcessTriggers() {

  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {

    if (
      trigger.getHandlerFunction() ===
      "processEmails"
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );
    }
  });
}

Running the Export

After saving the project:

    Select startExport from the function selector.
    Click Run.
    Review the permissions requested by Google.
    Grant the required permissions.
    Open the generated Google Sheet.
    Wait until the status changes to COMPLETE.

The script automatically continues in additional runs when required.
Duplicate Handling

Email addresses are normalized to lowercase and stored using the address as a unique key.

For example:

Max@example.com
max@example.com
MAX@example.com

are treated as the same address:

max@example.com

The address therefore appears only once in the final spreadsheet.

Communication counts are tracked separately.

For example:

max@example.com
Received: 42
Sent: 8
Total: 50

Output

A completed export may look like:

Name              Email Address       Received  Sent  Total
----------------------------------------------------------------
Max Mustermann    max@example.com     42        8     50
Anna Miller       anna@example.com    17        3     20
Example Company   info@example.com    31        0     31

The resulting Google Sheet can be downloaded as .xlsx or .csv.
Migration to a New Email Provider

Exporting addresses is only one part of a complete Gmail migration.

A controlled migration should generally include several stages.
1. Inventory

Before closing or abandoning an account, identify:

    email messages
    contacts
    calendars
    important attachments
    Google Drive data
    forwarding rules
    filters
    labels
    account recovery options
    services using Google authentication

Do not assume that an email-address export represents a complete account backup.
2. Select a New Provider

When evaluating a replacement provider, consider more than price and mailbox size.

Relevant criteria may include:

    privacy policy
    data location
    encryption
    two-factor authentication
    passkey support
    IMAP/POP support
    export capabilities
    custom domain support
    backup options
    transparent terms of service

One particularly important criterion is portability.

A provider should make it reasonably possible to leave the service later.
3. Consider Using Your Own Domain

For long-term independence, a personal domain can be useful.

For example:

firstname@yourdomain.com

instead of:

firstname@gmail.com

The main advantage is provider independence.

If the email provider changes, the public email address can remain the same.
4. Maintain a Transition Period

Do not immediately delete the old Gmail account.

A transition period of several months can help identify forgotten dependencies.

During this period:

    update important contacts
    change newsletter subscriptions
    update online accounts
    update financial services
    update government services
    monitor the old mailbox
    identify services still using the old address

Don't Forget Account Dependencies

One of the most important parts of an email migration is identifying services where the Gmail address is used as an account identifier.

Examples include:

Online stores
Banking
Insurance
Government services
Cloud services
Social media
Forums
Newsletters
Software services
Hosting
Domain registrars

The exported contact list can help identify communication relationships.

However, it is not a replacement for a complete account inventory.
Security After Migration

After the migration:

    enable 2FA
    preferably use passkeys where supported
    verify recovery email addresses
    verify recovery phone numbers
    remove unused app passwords
    review third-party OAuth applications
    review active sessions
    review forwarding rules
    review email filters
    create an independent backup

Backup Strategy

A migration should not be treated as a backup strategy by itself.

Important data should have at least one additional copy.

A simple model could look like:

                 Gmail
                   |
        +----------+----------+
        |          |          |
      Local      New Mail    Offline
      Backup     Provider    Archive

The objective is not to create unnecessary copies.

The objective is to avoid making a single provider the only place where important data exists.
Privacy Considerations

The exported contact list may contain personal information.

For example:

    names
    email addresses
    communication frequency

Treat the resulting files as sensitive personal data.

Do not upload real exports to a public GitHub repository.

Never commit files such as:

contacts.csv
contacts.xlsx
gmail-export.csv
mail-archive/

to a public repository.
.gitignore

If you use Git for this project, consider adding at least:

# Personal Gmail exports
*.csv
*.xlsx
*.xls

# Local exports
export/
exports/
backup/

# Credentials and configuration
*.json

# Temporary files
*.tmp
*.log

Review the rules according to your specific environment.
What This Project Does NOT Do

This project:

    does not crack accounts
    does not bypass Google security controls
    does not collect passwords
    does not read browser cookies
    does not intentionally upload contacts to an external server
    does not attempt to bypass access restrictions

It uses Google's Apps Script environment and permissions explicitly granted by the account owner.
Security Best Practices

Automation involving personal data should follow the principle of least privilege.

Before authorizing a script, review its complete source code.

Be especially cautious about code such as:

UrlFetchApp.fetch(...)

when it is used to send data to an unfamiliar external server.

Also review code that combines cloud storage access with external uploads.

A script should request only the permissions required for its intended purpose.
Why This Approach?

The objective of a Gmail exit should not be:

    "Delete Google immediately."

A safer objective is:

    Identify dependencies, secure your data, migrate communication, test the new environment, and only then reduce or remove the old dependency.

A planned migration is safer than an abrupt account change.
Migration Checklist
Before Migration

    Email data backed up
    Contacts exported
    Calendar exported
    Important Drive data backed up
    Important accounts identified
    New email provider selected
    New mailbox created
    2FA enabled
    Recovery options configured

During Migration

    Important contacts informed
    Online accounts updated
    Newsletter subscriptions updated
    Government/financial/insurance services updated
    Forwarding configured if appropriate
    Old Gmail account monitored

After Migration

    Transition period completed
    Old forwarding rules reviewed
    Remaining dependencies identified
    Final backup created
    Gmail data verified
    Old account only deactivated/deleted after verification

Conclusion

Moving away from Gmail is less about the technical act of changing email providers and more about managing dependencies and maintaining control over your data.

A structured approach is:

Inventory → Backup → Migrate → Test → Transition → Decommission

For long-term independence, consider prioritizing:

    data portability
    independent backups
    custom domains
    strong authentication
    recovery independence
    transparent providers
    the ability to migrate again in the future

The goal is not simply to leave one provider.

The goal is to make sure that you can leave any provider when you choose to.
Disclaimer

This repository is intended for personal data migration and administration of accounts you own or are authorized to manage.

Review the complete source code before executing it.

You are responsible for the processing, storage, and protection of any personal data contained in the resulting exports. :::
