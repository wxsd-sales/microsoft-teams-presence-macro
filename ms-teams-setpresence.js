import xapi from 'xapi';

//SET USER SPECIFIC VARIABLES
let MSFT_TENANT = "PLEASE SET";                   //*REQUIRED* Set to your MS Tenant URL
let MSFT_USERID = "PLEASE SET";     //*REQUIRED* Set to the End User ID from MS Graph
const MSFT_APPID = "PLEASE SET";    //*REQUIRED* Application ID of your MS Graph App 
const MSFT_SECRET = "PLEASE SET";  //*REQUIRED* Graph Client Secret of your MS Graph App

//FIXED ATTRIBUTES *DO NOT CHANGE*
const MSFT_PRES_URL = "https://graph.microsoft.com/beta/users/" + MSFT_USERID + "/presence/setPresence";  //MS GRAPH API FOR PRESENCE
var MSFT_PRES_HEADER = '';                                                                                //GRAPH API HEADERS
var MSFT_TOKEN = '';                                                                                      //PLACEHOLDER FOR TOKEN
var MSFT_PRES_PAYLOAD = '';                                                                               //PLACEHOLDER FOR PAYLOAD

//MSFT OAUTH CODE - BASED ON WORK BY CHRIS HARTLEY - THANKS :D
async function authMSFT()
{
  var oauthurl = "https://login.microsoftonline.com/" + MSFT_TENANT + "/oauth2/v2.0/token";  
  var loginData = "client_id=" + MSFT_APPID + "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret=" + MSFT_SECRET + "&grant_type=client_credentials";
  let headers = ['Content-Type: application/x-www-form-urlencoded'];
  
  //TROUBLESHOOTING - LOG TO CONSOLE
  //console.log(oauthurl);
  //console.log(loginData);

  //SUBMIT THE REQUEST AND CHECK FOR TOKEN RESPONSE    
  try 
  {
    const get_token = await xapi.command('HttpClient Post', {Url: oauthurl, Header: headers, AllowInsecureHTTPS: 'True', ResultBody: 'PlainText'}, loginData);
    
    if (get_token.StatusCode == 200)
    {      
      let token_json = JSON.parse(get_token.Body);

      //TROUBLESHOOTING - LOG TO CONSOLE
      //console.log(get_token.Body);
      
      try
      {
        MSFT_TOKEN = token_json.access_token;

        //TROUBLESHOOTING - LOG TO CONSOLE
        console.log(MSFT_TOKEN);

        MSFT_PRES_HEADER = ['Content-Type: application/json', 'Authorization: Bearer ' + MSFT_TOKEN];
      } 
      catch(error) 
      {
        console.error("Failed to get the token");
        console.error(error);          
      }        
    } 
    else
    {
      console.log("Failed to connect to " + oauthurl);
    }    
  } 
  catch(error)
  {
    console.error("Error in HTTPClient POST");
    console.error(error);
  }
}

//PERFORM AUTH AND UPDATE ON CALL START
xapi.event.on('OutgoingCallIndication', (event) => {

  //AUTHORIZE WITH GRAPH API
  authMSFT().then(
    (result) => {
      
      //LOG PRESENCE CHANGE
      console.log("Setting MSFT Presence to Busy:InAConferenceCall");

      //SET MS GRAPH PAYLOAD
      MSFT_PRES_PAYLOAD = '{"sessionId":"' + MSFT_APPID + '","availability":"Busy","activity":"InAConferenceCall","expirationDuration":"PT1H"}';

      //SEND API UPDATE TO MS GRAPH
      xapi.command('HttpClient Post', {Url: MSFT_PRES_URL, Header: MSFT_PRES_HEADER}, MSFT_PRES_PAYLOAD).then((result) => {
          console.log("MS Graph SetPresence Success:" + result.StatusCode);
        })
        .catch((err) => {
            console.log("MS Graph SetPresence Failed:" + err.message);
        });
  })
});

//UPDATE STATUS ON CALL DISCONNECT
xapi.event.on('CallDisconnect', (event) => {

  //LOG ACTIVITY
  console.log("Setting MSFT Presence to Available:Available");

  //SET MS GRAPH PAYLOAD
  MSFT_PRES_PAYLOAD = '{"sessionId":"' + MSFT_APPID + '","availability":"Available","activity":"Available","expirationDuration":"PT1H"}';

  //SEND API UPDATE TO MS GRAPH
  xapi.command('HttpClient Post', {Url: MSFT_PRES_URL, Header: MSFT_PRES_HEADER}, MSFT_PRES_PAYLOAD).then((result) => {
      console.log("MS Graph SetPresence Success:" + result.StatusCode);
    })
    .catch((err) => {
        console.log("MS Graph SetPresence Failed:" + err.message);
    });
});