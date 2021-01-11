/**
 * Logs the redirect URI to register.
 */
function logRedirectUri() {
  var service = getSmartService();
  Logger.log(service.getRedirectUri());
}

/**
 * Create the OAuth 2 service
 */
function getSmartService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth2.createService('smd')

      // Set the endpoint URLs, which are the same for all Google services.
      .setAuthorizationBaseUrl('https://nestservices.google.com/partnerconnections/' + PROJECT_ID + '/auth')
      .setTokenUrl('https://www.googleapis.com/oauth2/v4/token')

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(OAUTH_CLIENT_ID)
      .setClientSecret(OAUTH_CLIENT_SECRET)

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scopes to request (space-separated for Google services).
      .setScope('https://www.googleapis.com/auth/sdm.service')

      // Below are Google-specific OAuth2 parameters.

      // Sets the login hint, which will prevent the account chooser screen
      // from being shown to users logged in with multiple accounts.
      .setParam('login_hint', Session.getEffectiveUser().getEmail())

      // Requests offline access.
      .setParam('access_type', 'offline')

      // Consent prompt is required to ensure a refresh token is always
      // returned when requesting offline access.
      .setParam('prompt', 'consent');
}

/**
 * Direct the user to the authorization URL
 */
function showSidebar() {
  
  const smartService = getSmartService();
  
  if (!smartService.hasAccess()) {

    // App does not have access yet
    const authorizationUrl = smartService.getAuthorizationUrl();

    const template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    
    template.authorizationUrl = authorizationUrl;
    
    const page = template.evaluate();

    SpreadsheetApp.getUi().showSidebar(page);

  } else {
    // App has access
    console.log('App has access');
    
    // make the API request
    makeRequest();
  }
}

/**
 * Handle the callback
 */
function authCallback(request) {
  
  const smartService = getSmartService();
  
  const isAuthorized = smartService.handleCallback(request);
  
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}