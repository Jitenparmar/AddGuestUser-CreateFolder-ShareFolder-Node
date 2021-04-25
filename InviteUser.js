var adal = require('adal-node');
var request = require('request');

const TENANT = "tenant.onmicrosoft.com";
const GRAPH_URL = "https://graph.microsoft.com";
const CLIENT_ID = "b0d1ad67-d9a8-433c-b3ca-be199d1acbac";
const CLIENT_SECRET = "u0cGZv_bmASdbkS3wpI~G0-70N~VQ.bJNK";
const GROUP_ID = "e1c614ec-e2e1-4eeb-90c6-0f6d38dcec16";

function getToken() {
    return new Promise((resolve, reject) => {
        const authContext = new adal.AuthenticationContext(`https://login.microsoftonline.com/a9363b1a-c633-4337-ac61-b75acb815ff2/oauth2/v2.0/token`);
        authContext.acquireTokenWithClientCredentials(GRAPH_URL, CLIENT_ID, CLIENT_SECRET, (err, tokenRes) => {
            if (err) {
                reject(err);
            }
            var accesstoken = tokenRes.accessToken;
            resolve(accesstoken);
        });
    });
}


getToken().then(token => {
    /* INVITE A USER TO YOUR TENANT */
    var options = {
        method: 'POST',
        url: 'https://graph.microsoft.com/v1.0/invitations',
        headers: {
            'Authorization': 'Bearer ' + token,
            'content-type': 'application/json'
        },
        body: JSON.stringify({
            "invitedUserDisplayName": "Jay Doshi",
			"invitedUserEmailAddress": "jaydoshi.com@gmail.com",
			"inviteRedirectUrl": "https://tenant.sharepoint.com/",
			"sendInvitationMessage": false,
			"invitedUserMessageInfo": {
				"customizedMessageBody": "Hi Jiten, You are invited to the SharePoint site."
			}
        })
    };

    request(options, (error, response, body) => {
        if (!error && response.statusCode == 201) {
            var result = JSON.parse(body);
            // Log all the keys and values
            for (var key in result) {
                console.log(`${key}: ${JSON.stringify(result[key])}`);
            }

            /* ADD USER TO A GROUP */
            var options = {
                method: 'POST',
                url: 'https://graph.microsoft.com/v1.0/groups/${GROUP_ID}/members/$ref',
                headers: {
                    'Authorization': 'Bearer ' + token,
                    'content-type': 'application/json'
                },
                body: JSON.stringify({
                    "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${result.invitedUser.id}`
                })
            };

            request(options, (error, response, body) => {
                console.log(body);
                if (!error && response.statusCode == 204) {
                    console.log('OK');
                } else {
                    console.log('NOK');
                }
            });
        }
    });
});