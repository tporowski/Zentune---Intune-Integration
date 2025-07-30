document.addEventListener('DOMContentLoaded', () => {
    // Initialize the app after DOM is loaded
    initializeApp();
});

const client = ZAFClient.init();

const loginRequest = {
    scopes: ["openid", "profile", "User.Read"]
}

const tokenRequest = {
    scopes: ["User.Read", "DeviceManagementManagedDevices.Read.All"],
    forceRefresh: false
}

let appMetadata;
let myMSALObj;
let username = "";
let agent = "";
let tokenRefreshIntervalId = null;

// Main initialization function
async function initializeApp() {
    try {
        // First, get the metadata
        const metadata = await client.metadata();
        appMetadata = metadata.settings;

        // Now that we have metadata, we can initialize MSAL
        await initializeMSAL();

        // Set up event listeners after MSAL is initialized
        setupEventListeners();

    } catch (error) {
        console.error("App initialization error:", error);
        showError("Failed to initialize application: " + error.message);
    }
}

async function initializeMSAL() {
    const clientId = appMetadata.azure_client_id;
    const tenantId = appMetadata.azure_tenant_id;
    const subdomain = appMetadata.zendesk_subdomain;

    let redirectUri;
    if (subdomain.includes('localhost')) {
        redirectUri = `http://localhost:4567/0/assets/authRedirect.html`;
    } else {
        redirectUri = getRedirectUri();
    }

    const msalConfig = {
        auth: {
            clientId: clientId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
            redirectUri: redirectUri
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: false,
        }
    };

    // Initialize the MSAL instance
    myMSALObj = new msal.PublicClientApplication(msalConfig);

    // Initialize MSAL and handle authentication
    await myMSALObj.initialize();

    try {
        // Handle redirect promise
        const response = await myMSALObj.handleRedirectPromise();
        if (response) {
            handleResponse(response);
        } else {
            // Check if user is already logged in
            selectAccount();
        }
    } catch (error) {
        console.error("MSAL initialization error:", error);
        showError(error);
    }
}

function getRedirectUri() {
    const { origin, pathname } = window.location;
    const segments = pathname.split('/');
    segments.pop();

    segments.push('authRedirect.html');
    return origin + segments.join('/');
}

function setupEventListeners() {
    const signInBtn = document.getElementById('sign-in-btn');
    const logoutBtn = document.getElementById('logout-btn');
    const fetchDevicesBtn = document.getElementById('fetch-requester-devices');
    if (signInBtn) {
        signInBtn.addEventListener('click', signInPopup);
    }

    if (logoutBtn) {
        logoutBtn.addEventListener('click', signOut);
    }

    if (fetchDevicesBtn) {
        fetchDevicesBtn.addEventListener('click', fetchRequesterDevices);
    }

    // Removed swap accounts button and event listener

    window.addEventListener('unhandledrejection', ev => {
        console.error('Unhandled promise rejection', ev.reason);
        showError(`Unexpected error: ${ev.reason?.message || ev.reason}`);
    });
    window.addEventListener('error', ev => {
        console.error('Uncaught JS error', ev.error);
        showError(`Unexpected error: ${ev.error?.message || ev.message}`);
    });

}

/**
 * Calls getAllAccounts and determines the correct account to sign into, currently defaults to first account found in cache.
 */
async function selectAccount() {
    if (!myMSALObj) {
        console.error("MSAL not initialized");
        return;
    }

    const currentAccounts = myMSALObj.getAllAccounts();

    if (currentAccounts.length === 0) {
        updateUI("signed-out");
        return;
    } else if (currentAccounts.length > 1) {
        const selectedAccount = await showAccountSelectionPopup(currentAccounts);

        if (selectedAccount) {
            username = selectedAccount.username;
            updateUI("signed-in", selectedAccount);
            showStatus(`Switched to account: ${selectedAccount.name || selectedAccount.username}`);
        } else {
        }
    } else if (currentAccounts.length === 1) {
        let account = currentAccounts[0];
        username = account.username;
        updateUI("signed-in", currentAccounts[0]);
    }
}



/**
 * Sign in using popup
 */
function signInPopup() {
    if (!myMSALObj) {
        showError("MSAL not initialized. Please refresh the page.");
        return;
    }

    showStatus("Signing in with popup...");
    myMSALObj.loginPopup(loginRequest)
        .then(handleResponse)
        .catch((error) => {
            console.error("Login popup error:", error);
            showError(error);
        });
}

function signOut() {
    if (!myMSALObj) {
        showError("MSAL not initialized.");
        return;
    }

    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(username),
        postLogoutRedirectUri: window.location.origin
    };

    showStatus("Signing out...");

    // Use popup logout instead of redirect logout for iframe compatibility
    myMSALObj.logoutPopup(logoutRequest)
        .then(() => {
            username = "";
            
            // Reset fetch button state for next user
            resetFetchButton();
            
            updateUI("signed-out");
            showStatus("Signed out successfully");

            // Check for remaining accounts after successful logout
            const currentAccounts = myMSALObj.getAllAccounts();
            if (currentAccounts.length > 0) {
                selectAccount();
            }
        })
        .catch((error) => {
            console.error("Logout error:", error);
            // If user canceled logout or popup was closed, don't change UI state
            if (error.message && error.message.includes('user_cancelled')) {
                showStatus("Logout cancelled by user");
            } else {
                showError(error);
            }
        });
}

// Ensure proper UI updates based on authentication state
async function updateUI(state, account = null) {
    const title = document.getElementById("title");
    const statusElement = document.getElementById("status");
    const errorElement = document.getElementById("error");
    const msalElement = document.getElementById("msal");
    const logout = document.getElementById("logout-btn");
    let fetchBtn = document.getElementById("fetch-requester-devices");

    if (state === "signed-in" && account) {
        // Welcome user by first name
        let agent;
        if (account.name) {
            let cleanName = account.name.replace(/,/g, "");
            let split_name = cleanName.split(' ');
            if (split_name.length > 0 && split_name[0]) {
                if (!(account.username).includes(split_name[0].toLowerCase())) {
                    agent = split_name[0];
                } else if (split_name.length > 1 && split_name[1]) {
                    agent = split_name[1];
                } else {
                    agent = split_name[0];
                }
            } else {
                agent = account.username;
            }
        } else {
            agent = account.username;
        }
        if (title) title.textContent = `Welcome, ${agent}!`;
        if (errorElement) errorElement.classList.add("hidden");
        if (msalElement) msalElement.classList.add("hidden");
        if (logout) logout.classList.remove("hidden");
        if (fetchBtn) fetchBtn.classList.remove("hidden");
    } else {
        // Reset the title to the default app name on sign out
        if (title) title.textContent = "Intune Explorer";
        if (statusElement) statusElement.textContent = "Not signed in. Please sign in to continue.";
        // Remove the device list if it exists
        const devicesContainer = document.getElementById('devices-container');
        if (devicesContainer) {
            devicesContainer.remove();
        }
        try {
            if (errorElement) errorElement.classList.add("hidden");
            if (msalElement) msalElement.classList.remove("hidden");
            if (logout) logout.classList.add("hidden");
            if (fetchBtn) fetchBtn.classList.add('hidden');
        } catch (e) {
            console.error("Error updating UI:", e);
            showError("Failed to update UI: " + e.message);
        }
    }
}

function handleResponse(response) {
    if (response !== null) {
        username = response.account.username;
        updateUI("signed-in", response.account);
    } else {
        selectAccount();
    }
}

function signedIn(account = null) {
    updateUI("signed-in", account);
}

// Function to disable the fetch button when user lacks permissions
function disableFetchButton() {
    const fetchBtn = document.getElementById("fetch-requester-devices");
    if (fetchBtn) {
        fetchBtn.disabled = true;
        fetchBtn.title = "You don't have permission to access device information. Contact your administrator for DeviceManagementManagedDevices.Read.All permission.";

        // Add visual indicator
        if (!fetchBtn.textContent.includes('🔒')) {
            const originalText = fetchBtn.textContent;
            fetchBtn.textContent = `🔒 ${originalText}`;
        }
    }
}

// Function to reset the fetch button state when user logs out
function resetFetchButton() {
    const fetchBtn = document.getElementById("fetch-requester-devices");
    if (fetchBtn) {
        fetchBtn.disabled = false;
        fetchBtn.title = "Fetch devices for the ticket requester";
        
        // Remove lock icon if it exists
        if (fetchBtn.textContent.includes('🔒')) {
            fetchBtn.textContent = fetchBtn.textContent.replace('🔒 ', '');
        }
    }
}

// Modified getAccessToken to return the token
async function getAccessToken() {
    if (!myMSALObj) {
        throw new Error("MSAL not initialized.");
    }

    const account = myMSALObj.getAccountByUsername(username);

    if (!account) {
        throw new Error("No account found. Please sign in first.");
    }

    const accessTokenRequest = {
        ...tokenRequest,
        account: account
    };

    try {
        const response = await myMSALObj.acquireTokenSilent(accessTokenRequest);
        return response.accessToken;
    } catch (error) {
        // If silent acquisition fails, try popup
        try {
            const popupResponse = await myMSALObj.acquireTokenPopup(accessTokenRequest);
            return popupResponse.accessToken;
        } catch (popupError) {
            throw new Error(`Token acquisition failed: ${popupError.message}`);
        }
    }
}

// Token expiry monitoring
// Set up interval and clear it on app.willDestroy
if (typeof ZAFClient !== 'undefined' && client && typeof client.on === 'function') {
    tokenRefreshIntervalId = setInterval(async () => {
        try {
            await getAccessToken(); // This will refresh if needed
        } catch (error) {
            // Swallow error silently
        }
    }, 30 * 60 * 1000); // Check every 30 minutes

    client.on('app.willDestroy', () => {
        if (tokenRefreshIntervalId) {
            clearInterval(tokenRefreshIntervalId);
            tokenRefreshIntervalId = null;
        }
    });
}

// New function to fetch requester devices
async function fetchRequesterDevices() {
    // Check if button is already disabled from previous permission failure
    const fetchBtn = document.getElementById("fetch-requester-devices");
    if (fetchBtn && fetchBtn.disabled) {
        Swal.fire({
            title: 'Access Denied',
            text: 'You don\'t have the required permissions to fetch device information. Please contact your administrator.',
            icon: 'error',
            confirmButtonColor: '#0078d4'
        });
        return;
    }

    let accessToken;
    let endpoint;
    let devicesData;
    let ticketData;
    let requesterEmail;
    try {
        try {
            // First, get the requester's email from Zendesk
            ticketData = await client.get('ticket');
            requesterEmail = ticketData.ticket.requester.email;
            requesterName = ticketData.ticket.requester.name;

            // Get access token
            let account = myMSALObj.getAccountByUsername(username);

            if (!account) {
                console.error("No account found. Please sign in first.");
            }

            try {
                const response = await myMSALObj.acquireTokenSilent({ ...tokenRequest, account: account });
                accessToken = response.accessToken;
            } catch (error) {
                console.warn("Silent token acquisition failed, trying popup:", error);
                try {
                    const popupResponse = await myMSALObj.acquireTokenPopup(accessTokenRequest);
                    accessToken = popupResponse.accessToken;
                } catch (popupError) {
                    console.error(`Token acquisition failed: ${popupError.message}`);
                }
            }

            endpoint = `https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=userPrincipalName eq '${requesterEmail}'`;
        } finally {
            // Call Microsoft Graph
            devicesData = await callMSGraph(endpoint, accessToken);
            accessToken = null;
            document.getElementById("fetch-requester-devices").classList.add("hidden");
        }
        if (devicesData.value.length === 0) {
            // Show a warning when no devices are found
            Swal.fire({
                title: 'No Devices Found',
                text: `No managed devices were found for ${requesterName}. This user may not have any enrolled devices or may be using unmanaged devices.`,
                icon: 'info',
                confirmButtonColor: '#0078d4',
                confirmButtonText: 'OK'
            });
        } else {
            // Only show multiple devices warning if there are actually devices
            if (devicesData.value.length > 2) {
                Swal.fire({
                    toast: true,
                    position: 'top-start',
                    icon: 'warning',
                    title: 'Warning',
                    text: `Requester has ${(devicesData.value).length} devices, please investigate before resolving the ticket.`,
                    timer: null,
                    showConfirmButton: false
                });
            }
            // Only prompt to tattoo if there are devices to tattoo
            tattooDevices(devicesData.value, requesterName);
        }

        displayDevices(devicesData.value, requesterName);

    } catch (error) {
        console.error("Error fetching requester devices:", error);

        // Handle specific permission errors
        if (error.message && (
            error.message.includes('insufficient_scope') ||
            error.message.includes('AADSTS65001') ||
            error.message.includes('Forbidden') ||
            error.message.includes('401')
        )) {
            Swal.fire({
                title: 'Permission Denied',
                text: 'You don\'t have the required permissions to access device information. Please contact your administrator to request the "DeviceManagementManagedDevices.Read.All" permission.',
                icon: 'error',
                confirmButtonColor: '#0078d4',
                confirmButtonText: 'OK'
            }).then(() => {
                // Disable the button after showing the permission error
                disableFetchButton();
            });
        } else {
            showError(`Failed to fetch devices: ${error.message}`);
        }
    }
}

function tattooDevices(devices, name) {
    if (devices) {
        Swal.fire({
            title: 'Tattoo Device Information',
            text: 'Would you like to tattoo the device information into the ticket as an internal note?',
            icon: 'question',
            showCancelButton: true,
            confirmButtonText: 'Yes, tattoo it!',
            cancelButtonText: 'No, cancel',
            confirmButtonColor: '#0078d4',
            cancelButtonColor: '#6c757d'
        }).then(async (result) => {
            if (result.isConfirmed) {
                const deviceLinks = devices.map(device =>
                    `-\u00A0[${device.deviceName || 'Unknown Device'}](https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId/${device.id})\n\u00A0\u00A0\u00A0├─ ${device.model || 'Unknown Model'}\n\u00A0\u00A0\u00A0└─ ${device.serialNumber || 'Unknown Serial Number'}`
                ).join('\n');

                let ticketData = await client.get('ticket');

                const noteData = {
                    ticket: {
                        comment: {
                            body: `${name}'s Devices\n${deviceLinks}`,
                            public: false
                        }
                    }
                };

                try {
                    await client.request({
                        url: '/api/v2/tickets/' + ticketData.ticket.id,
                        type: 'PUT',
                        contentType: 'application/json',
                        data: JSON.stringify(noteData)
                    });
                    const Toast = Swal.mixin({
                        toast: true,
                        position: 'bottom-end',
                        showConfirmButton: false,
                        timer: 3000,
                        timerProgressBar: true,
                    });
                    Toast.fire({
                        icon: 'success',
                        title: 'Device information has been tattooed into the ticket.'
                    });
                } catch (error) {
                    console.error("Error updating ticket:", error);
                    showError("Failed to tattoo device information into the ticket.");
                }
            }
        });

    }
}

// Function to display devices in the UI
function displayDevices(devices, requesterEmail) {
    const existingCards = document.querySelectorAll('.ms-Card');
    existingCards.forEach(card => card.remove());

    // Create a container for devices if it doesn't exist
    let devicesContainer = document.getElementById('devices-container');
    if (!devicesContainer) {
        devicesContainer = document.createElement('div');
        devicesContainer.id = 'devices-container';
        devicesContainer.style.marginTop = '20px';
        document.body.appendChild(devicesContainer);
    }

    // Clear existing content
    devicesContainer.innerHTML = '';

    if (!devices || devices.length === 0) {
        showStatus(`No devices found for ${requesterName}`);

        // Add a visual warning card for no devices
        const noDevicesCard = document.createElement('div');
        noDevicesCard.className = 'ms-Card';
        noDevicesCard.style = 'min-width: 325px; max-width: 500px; margin: 12px auto; background: #fff4e6; border: 1px solid #ffb900; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);';

        noDevicesCard.innerHTML = `
            <div class="ms-Card-header" style="padding: 16px 20px 8px 20px; border-bottom: 1px solid #ffb900;">
                <div style="font-size: 18px; font-weight: 600; color: #8a8000; margin-bottom: 4px; display: flex; align-items: center;">
                    <span style="margin-right: 8px;">⚠️</span>
                    No Devices Found
                </div>
            </div>
            <div class="ms-Card-section" style="padding: 16px 20px;">
                <div style="font-size: 14px; color: #8a8000; line-height: 1.4;">
                    No managed devices were found for <strong>${requesterName}</strong>.<br>
                    This user may not have any enrolled devices or may be using unmanaged devices.
                </div>
            </div>
        `;

        devicesContainer.appendChild(noDevicesCard);
        return;
    }

    showStatus(`Found ${devices.length} device(s) for ${requesterName}  `);

    // Add a header
    const header = document.createElement('h3');
    header.textContent = `Devices for ${requesterEmail}:`;
    header.style.color = '#333';
    header.style.marginBottom = '15px';
    devicesContainer.appendChild(header);

    // Render each device
    devices.forEach(device => {
        renderDeviceCard(device, devicesContainer);
    });
}

function callMSGraph(endpoint, token) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(endpoint, options)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        });
}

function showStatus(message) {
    const statusElement = document.getElementById("status");
    if (statusElement) {
        statusElement.textContent = message;
    }
}

function showError(error) {
    const errorElement = document.getElementById("error");
    if (!errorElement) return;

    let errorMessage = "";

    if (typeof error === 'string') {
        errorMessage = error;
    } else if (error.message) {
        errorMessage = error.message;
    } else {
        errorMessage = JSON.stringify(error);
    }

    errorElement.textContent = `Error: ${errorMessage}. Try reloading the extension or force reloading the page (Ctrl+Shift+R).`;
    errorElement.classList.remove("hidden");
    console.error("MSAL Error:", error);

    if (error.message.includes('AADSTS65001')) {
        showError("Admin consent required. Please contact your IT administrator.");
    } else if (error.message.includes('insufficient_scope')) {
        showError("Missing required permissions. Please contact your IT administrator.");
    }
}

function clearError() {
    const errorElement = document.getElementById("error");
    if (errorElement) {
        errorElement.textContent = "";
        errorElement.classList.add("hidden");
    }
}

function renderDeviceCard(device, container = null) {
    const card = document.createElement('div');
    card.className = 'ms-Card';
    card.style = 'min-width: 325px; max-width: 500px; margin: 12px auto; background: #f8f9fa; border: 1px solid #e1e5e9; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);';

    // Format last sync date if available
    let lastSyncDate = 'Never';
    if (device.lastSyncDateTime) {
        lastSyncDate = new Date(device.lastSyncDateTime).toLocaleString();
    }

    // Determine compliance status color
    let complianceColor = '#605e5c';
    if (device.complianceState === 'compliant') {
        complianceColor = '#107c10';
    } else if (device.complianceState === 'noncompliant') {
        complianceColor = '#d13438';
    }

    card.innerHTML = `
        <div class="ms-Card-header" style="padding: 16px 20px 8px 20px; border-bottom: 1px solid #e1e5e9; position: relative;">
          <div style="font-size: 18px; font-weight: 600; color: #323130; margin-bottom: 4px;">
            ${device.deviceName || device.displayName || 'Unknown Device'}
          </div>
          <div style="font-size: 14px; color: #605e5c;">
            ${device.operatingSystem || 'Unknown OS'} • ${device.model || 'Unknown Model'}
          </div>
          <a href="https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId/${device.id}" 
             target="_blank" 
             style="position: absolute; top: 16px; right: 20px; padding: 6px 12px; background: #0078d4; color: white; text-decoration: none; border-radius: 4px; font-size: 12px; font-weight: 500; transition: background-color 0.2s ease;"
             onmouseover="this.style.backgroundColor='#106ebe'" 
             onmouseout="this.style.backgroundColor='#0078d4'">
            <span class="ms-fontColor-themeLighter ms-fontWeight-semibold ms-font-l">View in Intune</span>
          </a>
        </div>
        <div class="ms-Card-section" style="padding: 12px 20px;">
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; font-size: 14px;">
            <div>
              <div style="font-weight: 600; color: #323130; margin-bottom: 2px;">User</div>
              <div style="color: #605e5c;">${device.userDisplayName || 'N/A'}</div>
            </div>
            <div>
              <div style="font-weight: 600; color: #323130; margin-bottom: 2px;">Serial Number</div>
              <div style="color: #605e5c;">${device.serialNumber || 'N/A'}</div>
            </div>
            <div>
              <div style="font-weight: 600; color: #323130; margin-bottom: 2px;">Compliance</div>
              <div style="color: ${complianceColor}; font-weight: 500;">${device.complianceState || 'Unknown'}</div>
            </div>
            <div>
              <div style="font-weight: 600; color: #323130; margin-bottom: 2px;">Last Sync</div>
              <div style="color: #605e5c;">${lastSyncDate}</div>
            </div>
          </div>
        </div>
        <div class="ms-Card-section" style="padding: 8px 20px 16px 20px; border-top: 1px solid #f3f2f1;">
          <div style="font-size: 12px; color: #8a8886;">
            Enrolled: ${device.enrolledDateTime ? new Date(device.enrolledDateTime).toLocaleString() : 'N/A'}
          </div>
        </div>
      `;

    if (container) {
        container.appendChild(card);
    } else {
        document.body.appendChild(card);
    }
}

// Function to show account selection popup using SweetAlert
async function showAccountSelectionPopup(accounts) {
    if (!accounts || accounts.length === 0) {
        return null;
    }

    if (accounts.length === 1) {
        return accounts[0];
    }

    // Create options for the dropdown
    const options = {};
    accounts.forEach((account, index) => {
        const displayName = account.name || account.username;
        const email = account.username;
        options[index] = `${displayName} (${email})`;
    });

    try {
        const { value: selectedIndex } = await Swal.fire({
            title: 'Select your account',
            icon: 'question',
            input: 'select',
            inputOptions: options,
            inputPlaceholder: 'Choose an account...',
            showCancelButton: true,
            confirmButtonText: 'Switch Account',
            cancelButtonText: 'Cancel',
            confirmButtonColor: '#0078d4',
            cancelButtonColor: '#6c757d',
            inputValidator: (value) => {
                if (!value) {
                    return 'Please select an account!';
                }
            },
            customClass: {
                popup: 'swal-account-popup',
                title: 'swal-account-title',
                input: 'swal-account-select'
            },
            html: `
                <div style="margin-bottom: 15px; color: #605e5c; font-size: 14px;">
                    You have multiple accounts signed in. Please select which account you'd like to use.
                </div>
            `,
            didOpen: () => {
                // Style the select dropdown
                const select = Swal.getInput();
                if (select) {
                    select.style.fontSize = '14px';
                    select.style.padding = '8px 12px';
                    select.style.borderRadius = '4px';
                    select.style.border = '1px solid #d1d1d1';
                }
            }
        });

        if (selectedIndex !== undefined) {
            const selectedAccount = accounts[parseInt(selectedIndex)];
            return selectedAccount;
        }

        return null;

    } catch (error) {
        console.error('Error in account selection popup:', error);
        showError('Failed to show account selection dialog');
        return null;
    }
}