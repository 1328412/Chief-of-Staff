/**
 * Executive Chief of Staff - Microsoft 365 & Graph API Live Integration Module
 */

export class M365Service {
  constructor() {
    this.msalConfig = {
      auth: {
        clientId: "YOUR_AZURE_CLIENT_ID", // Replace with registered Azure App Client ID
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.href.split('#')[0]
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    };

    this.scopes = [
      "User.Read",
      "Calendars.ReadWrite",
      "Mail.ReadWrite",
      "Mail.Send",
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
      alert(`Microsoft 365 Login failed: ${error.message}\n\nNote: To connect to live enterprise data, register an App in Azure Active Directory and set Client ID.`);
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

  // Live M365 Data Fetchers
  async getLiveCalendarEvents() {
    if (!this.isLiveMode) return null;
    const data = await this.fetchGraphApi('/me/calendar/events?$top=10&$orderby=start/dateTime');
    return data ? data.value : [];
  }

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
