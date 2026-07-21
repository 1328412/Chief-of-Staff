import { M365Service } from './m365.js';
import { Triage } from './triage.js';
import { Schedule } from './schedule.js';
import { Briefing } from './briefing.js';
import { Orchestrator } from './orchestrator.js';

class App {
    constructor() {
        // Initialize the core services and UI components
        this.m365 = new M365Service();
        this.triage = new Triage();
        this.schedule = new Schedule();
        this.briefing = new Briefing();
        this.orchestrator = new Orchestrator();

        // DOM element references for interaction and updates
        this.connectButton = document.getElementById('connect-button');
        this.statusIndicator = document.getElementById('status-indicator');
        this.commandBar = document.getElementById('command-bar');
    }

    /**
     * Initializes the application, sets up event listeners, and renders the initial state.
     */
    async init() {
        this.setupEventListeners();

        // Attempt to initialize MSAL and check for an existing login session
        try {
            const account = await this.m365.init();
            if (account) {
                // If a user is already logged in, immediately switch to live mode
                console.log("User already logged in. Switching to live mode.");
                this.switchToLiveMode();
            } else {
                // Otherwise, render the app in demo mode
                console.log("No active session. Rendering in demo mode.");
                this.renderDemoMode();
            }
        } catch (error) {
            console.error("Initialization failed:", error);
            this.renderDemoMode(); // Fallback to demo mode on error
        }
    }

    /**
     * Sets up the primary event listener for the Microsoft 365 connect button.
     */
    setupEventListeners() {
        if (this.connectButton) {
            this.connectButton.addEventListener('click', () => this.login());
        }
    }

    /**
     * Handles the login process.
     */
    async login() {
        try {
            // This line triggers the MSAL login popup
            const account = await this.m365.login();

            // *** THE CRITICAL FIX IS HERE ***
            // If the login is successful and we get an account object back,
            // we switch the application to live mode.
            if (account) {
                console.log("Login successful. Switching to live mode.");
                this.switchToLiveMode();
            }
        
