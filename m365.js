/**
 * Executive Chief of Staff - Microsoft 365 & Graph API Approved Integration Module
 */

export class M365Service {
  constructor() {
    this.msalConfig = {
      auth: {
        clientId: "fafd3149-1381-4d6a-a8b7-6f864f719de4",
authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://1328412.github.io/Chief-of-Staff/"      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    };

    // Approved Standard Read Connectors (Outlook Mail, OneDrive, SharePoint)
    this.scopes = [
      "User.Read",
      "Mail.Read",
      "Files.Read",
      "Files.Read.All",
      "Sites.Read.All"
    ];

    this.msalInstance = null;
    this.account = null;
    this.isLiveMode = false;
  }

  init() {
    if (window.msal) {
      this.msalInstance = new window.msal.PublicClientApplication(this.msalConfig);
    }
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
      this.isLiveMode = true;
      console.log("Logged in to Microsoft 365 as:", this.account.username);

      return this.account;
    } catch (error) {
      console.error("Microsoft 365 Login Error:", error);
      
      // If popup was blocked or interaction in progress, attempt reset
      if (error.errorCode === "interaction_in_progress" || error.errorCode === "popup_window_error") {
        alert("Login popup was blocked or interrupted. Clearing browser cache and resetting authentication state. Please click 'Connect Microsoft 365' again.");
        sessionStorage.clear();
        return null;
      }

      alert(`Microsoft 365 Login Notice: ${error.message}\n\nNote: To connect to live enterprise data, register an App in Azure Active Directory and set Client ID.`);
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
      const tokenResponse = await this.msalInstance.acquireTokenPopup({
        scopes: this.scopes
      });
      return tokenResponse.accessToken;
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
      throw new Error(`Graph API error: ${response.statusText}`);
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
}
