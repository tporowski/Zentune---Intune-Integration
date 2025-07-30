# Intune Explorer - Zendesk App

A powerful Zendesk application that integrates with Microsoft Intune to provide IT support agents with instant access to device information for ticket requesters.

## � What It Does

Intune Explorer is a Zendesk sidebar application that allows IT support agents to quickly view and analyze Microsoft Intune-managed devices for ticket requesters. The app integrates seamlessly with your Zendesk instance and uses Microsoft Graph API to fetch device information from Intune.

### Key Features

- **Device Discovery**: Automatically fetches all Intune-managed devices for the current ticket requester
- **Device Details**: Displays comprehensive device information including compliance status, last sync time, and hardware details
- **Quick Access**: Provides direct links to view devices in the Microsoft Intune admin center
- **Ticket Integration**: Optionally "tattoos" device information directly into tickets as internal notes
- **Multi-Account Support**: Handles multiple Microsoft accounts with account switching capabilities
- **Secure Authentication**: Uses Microsoft Authentication Library (MSAL) with proper OAuth 2.0 flows
- **Real-Time Data**: Always fetches the latest device information from Microsoft Graph

### How It Works

1. **Zendesk Integration**: The app runs as a sidebar widget in Zendesk Support tickets
2. **Authentication Flow**: Uses MSAL.js for secure OAuth 2.0 authentication with Microsoft Entra ID
3. **Data Retrieval**: Queries Microsoft Graph API to fetch device information from Intune
4. **UI Rendering**: Displays device cards with formatted information and direct action buttons

## 📦 How to Install

### Prerequisites

Before setting up the application, ensure you have:
- **Zendesk Admin Access**: Ability to install and configure private apps
- **Microsoft 365 Admin Access**: Permission to create app registrations in Entra ID
- **Intune License**: Active Microsoft Intune subscription
- **Required Permissions**: Ability to grant admin consent for Microsoft Graph permissions

### Step 1: Create Entra ID App Registration

1. Navigate to the [Azure Portal](https://portal.azure.com)
2. Go to **Entra ID** → **App registrations** → **New registration**
3. Fill in the registration details:
   - **Name**: `ZenTune`
   - **Supported account types**: `Accounts in this organizational directory only` (default)
   - **Redirect URI**: Leave blank (we'll configure this later)
4. Click **Register** to create the app

5. Once created, you'll be taken to the app overview page. **Keep in mind these values** for later:
   - **Application (client) ID**: Found on the Overview page
   - **Directory (tenant) ID**: Found on the Overview page

6. Go to **API permissions** → **Add a permission**
7. Select **Microsoft Graph** → **Delegated permissions**
8. Search for and add: `DeviceManagementManagedDevices.Read.All`
9. Click **Grant admin consent for [your tenant]** - this is crucial for the app to work

### Step 2: Install ZenTune from Zendesk Marketplace

1. In Zendesk Admin Center, go to **Apps and integrations** → **Apps** → **Zendesk Support apps**
2. Click **Browse marketplace** and search for "ZenTune"
3. Select the **ZenTune - Intune Integration** app
4. Click **Install** (pricing: $1/agent/month)
5. You'll be prompted for three configuration parameters:

   | Parameter | Where to Find It | Example |
   |-----------|------------------|---------|
   | **Zendesk Subdomain** | Your Zendesk URL: `https://{subdomain}.zendesk.com` | `mycompany` |
   | **Azure Client ID** | Application (client) ID from your app registration | `12345678-1234-1234-1234-123456789abc` |
   | **Azure Tenant ID** | Directory (tenant) ID from your app registration | `87654321-4321-4321-4321-cba987654321` |

6. Complete the installation and choose where you want the app to appear (typically ticket sidebar)

### Step 3: Configure the Redirect URI (Critical Step!)

1. **Open a Zendesk ticket** that has a requester with devices enrolled in Intune
2. **Verify the app appears** in the ticket sidebar (you should see "ZenTune" in the sidebar)
3. **Open browser developer tools** (F12 or Ctrl+Shift+I)
4. **Go to the Console tab** at the bottom of the developer tools
5. **Change the context** from "top" to "app_zentune" (or the exact name of your application) using the dropdown at the top of the console

   ![Step 1: Find the "top" dropdown](assets\step1.png)
   
   ![Step 2: Select "app_zentune" from dropdown](assets\step2.png)

6. **Type the following command** and press Enter:
   ```javascript
   getRedirectUri()
   ```
   
   ![Step 3: Enter getRedirectUri() command and copy the result](assets\step3.png)
   
   *Note: The redirect URI shown in the example image above is a developer URI and won't match your production URI. Your actual URI will look different.*

7. **Copy the returned URI** - it will look something like:
   ```
   https://123456.apps.zendesk.com/123456/assets/{APP GUID}/authRedirect.html
   ```

8. **Go back to your Azure app registration** → **Authentication**
9. **Click "Add a platform"** → **Single-page application**
10. **Paste the redirect URI** from step 7 into the redirect URI field
11. **Save** the configuration

### Step 4: Test the Application

1. **Reload your Zendesk page** (Ctrl+F5 or Cmd+Shift+R)
2. **Open a ticket** with a requester who has Intune-managed devices
3. **Look for the ZenTune app** in the ticket sidebar
4. **Click "Sign In"** to authenticate with your Microsoft account
5. **Click "Fetch User's Device(s)"** to test the integration

The app should now be fully functional! If you encounter any authentication errors, double-check that:
- The redirect URI in Azure exactly matches what was returned by `getRedirectUri()`
- Admin consent was granted for the Microsoft Graph permissions
- You're signed in with an account that has appropriate Intune permissions

---

**Author**: Tommy Porowski (tporowski17@gmail.com)  
**Version**: 1.0.0