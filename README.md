#  Microsoft Teams Presence Macro
 
This Webex Device Macro makes it possible to automatically set your MS Teams Presence depending on the call state of the device. It leverages the Microsoft Graph API to set the presence of the associated user and monitors the Webex Devices call state using the Cisco xAPIs.

![MS-Teams-Presence-Macro](https://user-images.githubusercontent.com/21026209/161542129-cae6671f-f50c-4fe0-9b6f-305a536c9987.png)

## Requirements
1. Cisco Webex Device running RoomOS/CE 9.6 or newer
2. Microsoft Graph App on your Microsoft Tenant

## Setup

1. Configure the Macro Javascript file with the four required parameters at the beginning
2. Upload the Macro through your Webex Devices web inteface. 

## Creating a MS Graph App
The Macro requires a Microsoft Graph App to perform its API calls
1. Follow this guide to [here](https://docs.microsoft.com/en-us/onedrive/developer/rest-api/getting-started/app-registration?view=odsp-graph-online) to setup your App
2. Take note of the following as you will need them for your Macro
* ``MFST_APPID`` - The App ID fo the Graph App
* ``MFST_SECRET`` - The Graph Client Secret of the Graph App
* ``MFST_TENANT`` - Your Microsoft Tenant domain, eg. exampleorg.onmicrosoft.com

## Getting the User ID
The user ID for these API calls isn't the email address of the user but instead a long alphanumeric string. The easiet way to find your own ID is to use the [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and click on the 'GET my Profile' link. The ID will be at the bottom of the reponse section on the right:

![image](https://user-images.githubusercontent.com/21026209/161545260-00c8861b-1e58-4052-b944-e88c29a77dc2.png)

Take note of this as you will need to include this to the Macro

* ``MFST_USERID`` - Your end user ID
