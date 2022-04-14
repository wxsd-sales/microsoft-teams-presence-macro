#  Microsoft Teams Presence Macro
 
This Webex Device Macro makes it possible to automatically set your Microsoft Teams Presence depending on the call state of your device. It leverages the Microsoft Graph API to set the presence of the associated user and monitors the Webex Devices call state using the Cisco xAPIs.

![MS-Teams-Presence-Macro](https://user-images.githubusercontent.com/21026209/161542129-cae6671f-f50c-4fe0-9b6f-305a536c9987.png)

## Requirements
1. Cisco Webex Device running RoomOS/CE 9.6 or newer
2. Microsoft Tenant and a Microsoft Graph App

## Setup

1. Configure the Macro Javascript file ``ms-teams-setpresence.js`` with the four required parameters at the beginning, additional instructions on where to find these are blow:
* ``MFST_TENANT`` - This is your Microsoft Tenants domain, eg. exampleorg.onmicrosoft.com
* ``MFST_APPID`` - This is the App ID of your Graph App (see instructions below)
* ``MFST_SECRET`` - This is the Graph Client Secret of the Graph App (see instructions below)
* ``MFST_USERID`` - Your end user ID (see instructions below)
3. Upload the Macro through your Webex Devices web interface. 


## Creating a Microsoft Graph App
The Macro requires a Microsoft Graph App to perform its API calls. Follow the steps below to create one:
1. Follow this guide [here](https://docs.microsoft.com/en-us/onedrive/developer/rest-api/getting-started/app-registration?view=odsp-graph-online) to setup your App. For mine, I called it Presence App.
2. After you create your App, on the overview page, you will find you ``Application (client) ID``, use this for the ``MFST_APPID`` parameter in the Macro.

![image](https://user-images.githubusercontent.com/21026209/161583122-922273dc-e20f-4fa0-9f48-27995eb0c26b.png)

3. Lastly, you will need to give the App permission to write the presence for your users, go to the A

## Give the App Presence Permissions
By default the App doesn't have the persmissions to modify the presence for users. You will need to add this permission:
1. Go to the API Permissions tab and delete the dafault User.Read permission.

![image](https://user-images.githubusercontent.com/21026209/163387262-3d9a7881-f84a-437a-94bc-ae00f55a4811.png)

2. Then click 'Add a permission', select the 'Microsoft Graph' option.

![image](https://user-images.githubusercontent.com/21026209/163387868-2fbfb1e5-52d2-4b10-b5d5-028270480344.png)

3. Search for the 'presence' permission and select the Presence.ReadWrite option and then click 'Add permissions'.

![image](https://user-images.githubusercontent.com/21026209/163388133-fd4dca9a-7ee5-47f0-bcbc-2e153eb6887f.png)



## Creating a Client Secret
In order to use the Graph App, we will need to generate a Client Secret:
1. Go to the 'Certicates & secrets' section of your app and click 'New client secret'

![image](https://user-images.githubusercontent.com/21026209/161586075-b964c883-af2a-4f5a-95fa-66a191e916cf.png)

2. Give it a Description and Expiration date and click add.

<img width="578" alt="image" src="https://user-images.githubusercontent.com/21026209/161584987-08593a53-1e6f-4757-978d-6fb51a226ddd.png">

3. The key in the Value column is our Client Secret, take note of this as it will vanish after you leave this page. Use this for your ``MFST_SECRET`` parameter in the Macro.

![image](https://user-images.githubusercontent.com/21026209/161585532-4e62555b-c945-47f5-a56a-a9ef8332a5b9.png)




## Getting the User ID
The user ID for these API calls isn't the email address of the user but instead a long alphanumeric string. The easiet way to find your own ID is to use the [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and click on the 'GET my Profile' link. The ID will be at the bottom of the reponse section on the right:

![image](https://user-images.githubusercontent.com/21026209/161586596-33cbc311-e5c9-41d2-b835-818ad1581805.png)


Take note of this as you will need it for the ``MFST_USERID`` parameter in the Macro.


## Contact

wxsd@cisco.external.com
