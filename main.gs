/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on in the notification emails.
 */
var ADDON_TITLE = 'Subscribe User to a Mailing List';

/**
 * Adds a custom menu to the active form to show the add-on sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  FormApp.getUi()
      .createAddonMenu()
      .addItem('Configure Subscription Behavior', 'showSidebar')
      .addItem('About', 'showAbout')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the form containing the add-on's user interface for
 * configuring the notifications this add-on will produce.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Configure Behavior');
  FormApp.getUi().showSidebar(ui);
}

/**
 * Opens a purely-informational dialog in the form explaining details about
 * this add-on.
 */
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('About')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(420)
      .setHeight(270);
  FormApp.getUi().showModalDialog(ui, 'About Subscribe to Mailing List');
}

/**
 * Save sidebar settings to this form's Properties, and update the onFormSubmit
 * trigger as needed.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(settings) {
  PropertiesService.getDocumentProperties().setProperties(settings);
  adjustFormSubmitTrigger();
}

/**
 * Queries the User Properties and adds additional data required to populate
 * the sidebar UI elements.
 *
 * @return {Object} A collection of Property values and
 *     related data used to fill the configuration sidebar.
 */
function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();

  // Get multiple choice field items with 2 choices (Yes/No) in the form and compile a list
  // of their titles and IDs to gather
  var form = FormApp.getActiveForm();
  var choiceItems = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
  settings.choiceItems = [];
  for (var i = 0; i < choiceItems.length; i++) {
    var choices = choiceItems[i].asMultipleChoiceItem().getChoices(); 
    if (choices.length == 2) {
      if (choices[0].getValue() == 'Yes' && choices[1].getValue() == 'No') {
        settings.choiceItems.push({
          title: choiceItems[i].getTitle(),
          id: choiceItems[i].getId()
        });
      }
    }
  }
  
  // Get text field items in the form and compile a list
  //   of their titles and IDs.
  var textItems = form.getItems(FormApp.ItemType.TEXT);
  settings.textItems = [];
  for (var i = 0; i < textItems.length; i++) {
    settings.textItems.push({
      title: textItems[i].getTitle(),
      id: textItems[i].getId()
    });
  }
  
  return settings;
}

/**
 * Adjust the onFormSubmit trigger based on user's requests.
 */
function adjustFormSubmitTrigger() {
  var form = FormApp.getActiveForm();
  var triggers = ScriptApp.getUserTriggers(form);
  var settings = PropertiesService.getDocumentProperties();
  var triggerNeeded = settings.getProperty('subscribeUserToMailingList') == 'true';

  // Create a new trigger if required; delete existing trigger
  // if it is not needed.
  var existingTrigger = null;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
      existingTrigger = triggers[i];
      break;
    }
  }
  if (triggerNeeded && !existingTrigger) {
    var trigger = ScriptApp.newTrigger('respondToFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
  } else if (!triggerNeeded && existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }
}

/**
 * Responds to a form submission event if a onFormSubmit trigger has been
 * enabled.
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
function respondToFormSubmit(e) {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger require authorizations that have not
  // been supplied yet -- if so, warn the active user via email (if possible).
  // This check is required when using triggers with add-ons to maintain
  // functional triggers.
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to reauthorize; the normal trigger action is not
    // conducted, since it authorization needs to be provided first. Send at
    // most one 'Authorization Required' email a day, to avoid spamming users
    // of the add-on.
    sendReauthorizationRequest();
  } else {
    // All required authorizations has been granted, so continue to respond to
    // the trigger event.

    // Check if the form respondent needs to be subscribed to the mailing list; if so, construct and
    // send the notification. Be sure to respect the remaining email quota.
    if (settings.getProperty('subscribeUserToMailingList') == 'true' &&
        MailApp.getRemainingDailyQuota() > 0) {
      subscribeUserToMailingList(e.response);
    }
  }
}


/**
 * Called when the user needs to reauthorize. Sends the user of the
 * add-on an email explaining the need to reauthorize and provides
 * a link for the user to do so. Capped to send at most one email
 * a day to prevent spamming the users of the add-on.
 */
function sendReauthorizationRequest() {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
  var today = new Date().toDateString();
  if (lastAuthEmailDate != today) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template =
          HtmlService.createTemplateFromFile('AuthorizationEmail');
      template.url = authInfo.getAuthorizationUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
          'Authorization Required',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
    settings.setProperty('lastAuthEmailDate', today);
  }
}

/**
 * Subscribes the respondent to the mailing list.
 *
 * @param {FormResponse} response FormResponse object of the event
 *      that triggered this notification
 */
function subscribeUserToMailingList(response) {
  var form = FormApp.getActiveForm();
  var settings = PropertiesService.getDocumentProperties();
  
  // get the user's consent to add them to the list
  var consentId = settings.getProperty('respondentConsentItemId');
  var consentItem = form.getItemById(consentId);
  var consent = response.getResponseForItem(consentItem).getResponse();
  
  // if consent is Yes, then go on to subscribe them to the mailing list
  if (consent == "Yes") {
  
    // get the respondent's e-mail address
    var emailId = settings.getProperty('respondentEmailItemId');
    var emailItem = form.getItemById(parseInt(emailId));
    var respondentEmail = response.getResponseForItem(emailItem).getResponse();
    
    // get the group's e-mail address
    var groupEmail = settings.getProperty('groupEmail');
    
    // if both respondent email and group e-mail are provided, subscribe the respondent to the group!
    if (respondentEmail && groupEmail) {
      var memberResource = {
        email: respondentEmail,
        role: "MEMBER"
      };
      
      var member = AdminDirectory.Members.insert(memberResource, groupEmail);
      if (member) {
        //Logger.log('Member %s added to: %s', member.email, groupEmail);
      } else {
        //Logger.log('Unable to add member %s to %s.', respondentEmail, groupEmail);
        // send an e-mail to the effective user that there was an issue.
        var errorMail = "An error occurred subscribing " + respondentEmail + " to the group " + groupEmail + ".";
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
          'A problem occurred subscribing a user to a group',
          errorMail,
          {
            name: ADDON_TITLE
          });
      }
    }
  }
}
