/**
 * VIS Administration Inquiry Form — Auto-creator
 * Valley International School
 *
 * HOW TO USE:
 * 1. Go to https://script.google.com
 * 2. Click "New project"
 * 3. Delete any existing code in the editor
 * 4. Paste this entire script
 * 5. Click the ▶ Run button (or press Ctrl+Enter)
 * 
 * 6. Grant permissions when prompted (Google needs access to create Forms)
 * 7. When it finishes, check the Execution log (View → Logs) for the form link
 * 8. Copy the link from the log — that's your shareable form URL
 */

function createAdminInquiryForm() {

  // Create the form
  var form = FormApp.create('VIS Administration Inquiry');
  form.setDescription(
    'Valley International School — Administration Inquiry\n' +
    'Please fill in all details so our team can follow up with you promptly.'
  );
  form.setCollectEmail(false); // Email is captured as a field instead
  form.setProgressBar(false);
  form.setShuffleQuestions(false);
  form.setConfirmationMessage(
    'Thank you for your enquiry! A member of our team will be in touch with you shortly.'
  );

  // ── SECTION: Personal Details ──────────────────────────────────────────────

  form.addSectionHeaderItem()
    .setTitle('Personal Details');

  // Date
  form.addDateItem()
    .setTitle('Date')
    .setRequired(true);

  // Name of Parent / Guardian
  form.addTextItem()
    .setTitle('Name of Parent / Guardian')
    .setRequired(true);

  // Name of Child
  form.addTextItem()
    .setTitle('Name of Child')
    .setRequired(true);

  // Age of Child
  form.addTextItem()
    .setTitle('Age of Child')
    .setHelpText('Please enter the child\'s current age in years.')
    .setRequired(true);

  // Contact Number
  form.addTextItem()
    .setTitle('Contact Number')
    .setHelpText('e.g. 012-3456789')
    .setRequired(true);

  // Email Address
  form.addTextItem()
    .setTitle('Email Address')
    .setHelpText('We will send confirmation and follow-up details here.')
    .setRequired(true);

  // Previous School
  form.addTextItem()
    .setTitle('Previous School')
    .setHelpText('Enter "N/A" if this is your child\'s first school.')
    .setRequired(false);

  // ── SECTION: Enquiry Details ───────────────────────────────────────────────

  form.addSectionHeaderItem()
    .setTitle('Enquiry Details');

  // How did you hear about VIS?
  var referralItem = form.addMultipleChoiceItem();
  referralItem.setTitle('How did you hear about Valley International School?')
    .setRequired(true);

  // We need the "Other" option with a text box, and "Referral" with a text box.
  // Google Forms supports .createChoice() with open-ended responses for "Other".
  referralItem.setChoices([
    referralItem.createChoice('Facebook'),
    referralItem.createChoice('Instagram'),
    referralItem.createChoice('Website'),
    referralItem.createChoice('Referral (please state referral code below)'),
    referralItem.createChoice('Other (please describe below)')
  ]);
  referralItem.showOtherOption(false); // Using explicit text fields below instead

  // Referral code / Other details (shown when either of the last two is selected)
  form.addTextItem()
    .setTitle('Referral code / Additional details')
    .setHelpText(
      'If you selected "Referral", enter the referral code here.\n' +
      'If you selected "Other", please describe how you heard about us.'
    )
    .setRequired(false);

  // Programme Interested In
  var programmeItem = form.addCheckboxItem();
  programmeItem.setTitle('Programme(s) Interested In')
    .setHelpText('Select all that apply.')
    .setRequired(true)
    .setChoices([
      programmeItem.createChoice('Foundation'),
      programmeItem.createChoice('Primary'),
      programmeItem.createChoice('Secondary'),
      programmeItem.createChoice('Day Care'),
      programmeItem.createChoice('After School Programme')
    ]);

  // ── SECTION: Additional Notes ──────────────────────────────────────────────

  form.addSectionHeaderItem()
    .setTitle('Anything else?');

  form.addParagraphTextItem()
    .setTitle('Additional Notes or Questions')
    .setHelpText('Optional — any other information you\'d like to share.')
    .setRequired(false);

  // ── OUTPUT ─────────────────────────────────────────────────────────────────

  var formUrl = form.getPublishedUrl();
  var editUrl = form.getEditUrl();
  var shortUrl = form.shortenFormUrl(formUrl);

  Logger.log('===========================================');
  Logger.log('VIS Admin Inquiry Form created successfully!');
  Logger.log('');
  Logger.log('SHARE THIS LINK (for respondents):');
  Logger.log(shortUrl);
  Logger.log('');
  Logger.log('Full respondent URL: ' + formUrl);
  Logger.log('Edit URL (keep private): ' + editUrl);
  Logger.log('===========================================');

  // Also create a response spreadsheet
  var ss = SpreadsheetApp.create('VIS Admin Inquiries — Responses');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('Response spreadsheet created: ' + ss.getUrl());
  Logger.log('All responses will automatically appear there.');

  return shortUrl;
}
