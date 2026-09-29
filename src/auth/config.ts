// Fill in with the Entra ID app registration created for this plugin
// (public client, no secret, delegated Tasks.ReadWrite scope).
export const CLIENT_ID = 'REPLACE_WITH_ENTRA_APP_CLIENT_ID';
export const AUTHORITY = 'https://login.microsoftonline.com/common';
export const SCOPES = ['Tasks.ReadWrite'];
export const REDIRECT_URI = 'http://localhost:8080/callback';
