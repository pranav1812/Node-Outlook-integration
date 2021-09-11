// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');


module.exports = {
  getUserDetails: async (msalClient, userId)=> {
    const client = getAuthenticatedClient(msalClient, userId);

    const user = await client
      .api('/me')
      .select('displayName,mail,mailboxSettings,userPrincipalName')
      .get();
    return user;
  },
  getUserMails: async function(msalClient, userId, mailId= null, subject= null) {
    var filterString = ``;

    if (mailId && subject) {
      filterString= `(sender/emailAddress/address) eq '${mailId}' and contains(subject, '${subject}')`
    }else if (mailId) {
      filterString= `(sender/emailAddress/address) eq '${mailId}'`
    }else if (subject) {
      filterString= `contains(subject, '${subject}')`
    }
    console.log(filterString);

    const client = getAuthenticatedClient(msalClient, userId);
    console.log(3.0, mailId)
    // working filters:
    // "(sender/emailAddress/address) eq 'pranavvohra2000@gmail.com' and contains(subject, 'hello')"
    const mess = client.api('/me/messages')
    .select('subject,sender,toRecipients,receivedDateTime')
    .filter(filterString)
    .get();
    return mess;
  },
  sendMail: async(msalClient, userId, mailId= null, subject= null, body= null)=> {
      try {
        if(mailId){
          const client = getAuthenticatedClient(msalClient, userId);
          const emailObject= {
            "message": {
                "subject": subject || "no subject!",
                "body": {
                    "contentType": "Text",
                    "content": body || "no email body!"
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": mailId
                        }
                    }
                ]
            }    
          };
          const res= await client.api('me/sendMail').post(emailObject);
          return 'sent';
      }else{
        return 'no mail id';
      }
    } catch (error) {
        console.log(error);
      }

  } 


};

const getAuthenticatedClient= (msalClient, userId)=> {
  if (!msalClient || !userId) {
    throw new Error(
      `Invalid MSAL state. Client: ${msalClient ? 'present' : 'missing'}, User ID: ${userId ? 'present' : 'missing'}`);
  }

  // Initialize Graph client
  const client = graph.Client.init({
    // Implement an auth provider that gets a token
    // from the app's MSAL instance
    authProvider: async (done) => {
      try {
        // Get the user's account
        const account = await msalClient
          .getTokenCache()
          .getAccountByHomeId(userId);

        if (account) {
          // Attempt to get the token silently
          // This method uses the token cache and
          // refreshes expired tokens as needed
          const response = await msalClient.acquireTokenSilent({
            scopes: process.env.OAUTH_SCOPES.split(','),
            redirectUri: process.env.OAUTH_REDIRECT_URI,
            account: account
          });

          // First param to callback is the error,
          // Set to null in success case
          done(null, response.accessToken);
        }
      } catch (err) {
        console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        done(err, null);
      }
    }
  });

  return client;
}
