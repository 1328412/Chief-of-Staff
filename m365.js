/**
 * Executive Chief of Staff - Microsoft 365 & Graph API Approved Integration Module
 */

export class M365Service {
  constructor() {
    this.msalConfig = {
      auth: {
        clientId: "fafd3149-1381-4d6a-a8b7-6f864f719de4",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://1328412.github.io/Chief-of-Staff/"
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    };

    // Least-privilege delegated scopes for personal mailbox and personal OneDrive access.
    // NOTE: Organization-wide SharePoint scopes (e.g. Sites.Read.All) typically require admin consent.
    this.scopes = [
      "openid",
      "profile",
      "User.Read",
      "Mail.Read",
      "Files.Read"
    ];

    this.msalInstance = null;
    this.account = null;
    this.isLiveMode = false;
  }

  async init() {
    if (!window.msal) return null;

    this.msalInstance = new window.msal.PublicClientApplication(this.msalConfig);

    try {
      await this.msalInstance.handleRedirectPromise();
    } catch (error) {
      console.warn("MSAL redirect handling warning:", error);
    }

    const activeAccount = this.msalInstance.getActiveAccount();
    const accounts = this.msalInstance.getAllAccounts();
    this.account = activeAccount || accounts[0] || null;

    if (this.account) {
      this.msalInstance.setActiveAccount(this.account);
      this.isLiveMode = true;
    }

    return this.account;
  }

  async login() {
    if (!this.msalInstance) {
      alert("MSAL Library loading... Please ensure Azure Client ID is configured in settings.");
      return null;
    }

    try {
      // Clear stale interaction keys if present to prevent interaction_in_progress block
      for (let i = 0; i < sessionStorage.length; i++) {
        const key = sessionStorage.key(i);
        if (key && key.includes("msal.interaction.status")) {
          sessionStorage.removeItem(key);
        }
      }

      const loginResponse = await this.msalInstance.loginPopup({
        scopes: this.scopes,
        prompt: "select_account"
      });

      this.account = loginResponse.account;
      this.msalInstance.setActiveAccount(this.account);
      this.isLiveMode = true;
      console.log("Logged in to Microsoft 365 as:", this.account.username);

      return this.account;
    } catch (error) {
      console.error("Microsoft 365 Login Error:", error);

      const errorText = `${error?.errorCode || ""} ${error?.message || ""}`.toLowerCase();

      if (error.errorCode === "block_nested_popups") {
        // Some hosts run this app in a popup/webview. Redirect-based auth avoids nested popup restrictions.
        await this.msalInstance.loginRedirect({
          scopes: this.scopes,
          prompt: "select_account"
        });
        return null;
      }
      
      // If popup was blocked or interaction in progress, attempt reset
      if (error.errorCode === "interaction_in_progress" || error.errorCode === "popup_window_error") {
        alert("Login popup was blocked or interrupted. Clearing browser cache and resetting authentication state. Please click 'Connect Microsoft 365' again.");
        sessionStorage.clear();
        return null;
      }

      if (errorText.includes("admin") || errorText.includes("consent")) {
        const consentUrl = `https://login.microsoftonline.com/organizations/adminconsent?client_id=${encodeURIComponent(this.msalConfig.auth.clientId)}&redirect_uri=${encodeURIComponent(this.msalConfig.auth.redirectUri)}`;
        alert(
          `Admin consent is required by your organization policy.\n\n` +
          `App: Chief of Staff AI Agent\n` +
          `Client ID: ${this.msalConfig.auth.clientId}\n` +
          `Requested delegated scopes: ${this.scopes.join(", ")}\n\n` +
          `Ask your Microsoft 365 admin to grant consent using:\n${consentUrl}`
        );
        return null;
      }

      alert(`Microsoft 365 Login Notice: ${error.message}\n\nIf this app is running in a popup/embedded window, open it in a normal browser tab and sign in again.`);
      return null;
    }
  }

  async getAccessToken() {
    if (!this.account || !this.msalInstance) return null;

    try {
      const tokenResponse = await this.msalInstance.acquireTokenSilent({
        scopes: this.scopes,
        account: this.account
      });
      return tokenResponse.accessToken;
    } catch (e) {
      try {
        const tokenResponse = await this.msalInstance.acquireTokenPopup({
          scopes: this.scopes
        });
        return tokenResponse.accessToken;
      } catch (popupError) {
        if (popupError.errorCode === "block_nested_popups") {
          await this.msalInstance.acquireTokenRedirect({
            scopes: this.scopes,
            account: this.account
          });
          return null;
        }

        throw popupError;
      }
    }
  }

  async fetchGraphApi(endpoint) {
    const token = await this.getAccessToken();
    if (!token) return null;

    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      let detail = response.statusText;
      try {
        const payload = await response.json();
        detail = payload?.error?.message || payload?.error?.code || detail;
      } catch (_) {
        // Keep default status text when the error body is unavailable.
      }

      throw new Error(`Graph API ${response.status}: ${detail}`);
    }

    return await response.json();
  }

  // Live M365 Approved Data Fetchers
  async getLiveInboxMessages() {
    if (!this.isLiveMode) return null;
    const data = await this.fetchGraphApi('/me/messages?$top=15&$orderby=receivedDateTime desc');
    return data ? data.value : [];
  }

  async searchLiveSharePointAndOneDrive(query) {
    if (!this.isLiveMode) return null;
    const data = await this.fetchGraphApi(`/me/drive/root/search(q='${encodeURIComponent(query)}')`);
    return data ? data.value : [];
  }
}

export const m365Service = new M365Service();
