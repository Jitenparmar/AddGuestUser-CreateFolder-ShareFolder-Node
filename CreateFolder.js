var request = require('request');

const TENANT = "spo365devfactory.sharepoint.com";
const CLIENT_ID = "db7b843b-d7a1-4969-b665-61dcafec2444";
const CLIENT_SECRET = "isoxum99YPdk/K+0R8ZHOlTdrIdeyxcg6HKi3+j3Blk=";
const TENANT_ID = "a9363b1a-c633-4337-ac61-b75acb815ff2"

function ShareFolder() {
	var options = {
	  'method': 'POST',
	  'url': `https://accounts.accesscontrol.windows.net/${TENANT_ID}/tokens/OAuth/2`,
	  'headers': {
		'Content-Type': 'application/x-www-form-urlencoded',
		'Cookie': 'esctx=AQABAAAAAAD--DLA3VO7QrddgJg7Wevr1-nIB4V8sstvtkLoiQBnob_zeIkFOCaMiGMl-vobOQTl_1EcgJLcuKidaTW5Lw1DLRTMYrFL5cNxOFRX_Z_YAlAH6P92inQqj6sgsu-ElFwajONksYYt6406iOP1TcalmUiL73Rtl1giIeAVKPeVi3rB658jS3-3hvcKBx0YZD0gAA; fpc=AloiXTUgdMdNhpiG8s7qVDJG1hG6AQAAAGxXF9gOAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd'
	  },
	  form: {
		'grant_type': 'client_credentials',
		'client_id': CLIENT_ID+'@'+TENANT_ID,
		'client_secret': CLIENT_SECRET,
		'Resource': `00000003-0000-0ff1-ce00-000000000000/${TENANT}@${TENANT_ID}`
	  }
	};
	
    return new Promise((resolve, reject) => {
        request(options, function (error, response,body) {
		  if (error) throw new Error(error);
		  var result = JSON.parse(body);
		  resolve(result['access_token']);
		});
    });
}

/*createFolder().then(token => {
	//console.log("Final Token"+token);
	// Folder Create in Documents doc library
	var documentLibraryName = "Shared%20Documents";
	var folderName = "test1Demo";
	var SiteURL = "https://spo365devfactory.sharepoint.com";
	
	var options = {
	  'method': 'GET',
	  'url': SiteURL+`/_api/web/GetFolderByServerRelativeUrl(\'${documentLibraryName}/${folderName}\')/Exists`,
	  'headers': {
		'Accept': 'application/json;odata=verbose', 
		'Authorization': 'Bearer '+token
	  }
	};
	request(options, function (error, response,body) {
	  if (error) throw new Error(error);
	  var result = JSON.parse(body);
	  console.log(result['d']['Exists']);
	  if(!result['d']['Exists']){
		  var options = {
		  'method': 'POST',
		  'url': SiteURL+`/_api/Web/Folders/add(\'${documentLibraryName}/${folderName}\')`,
		  'headers': {
			'Authorization': 'Bearer '+token
		  }
		};
		request(options, function (error, response,body) {
		  if (error) throw new Error(error);
		  console.log("Folder Created!");
		});
	  }
	  //console.log(response.body);
	});

	
	
});*/

ShareFolder().then(token => {
	// Check and Resolve the User
		var SiteURL = "https://spo365devfactory.sharepoint.com";
		var UserEmail = "deval@spo365devfactory.onmicrosoft.com";
		var documentLibraryName = "Shared%20Documents";
		var folderName = "test1Demo";
		var options = {
		  'method': 'POST',
		  'url': SiteURL+'/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerResolveUser',
		  'headers': {
			'Authorization': 'Bearer '+token,
			'accept': 'application/json;odata=verbose',
			'Content-Type': 'application/json'
		  },
		  body: JSON.stringify({
			"queryParams": {
			  "AllowEmailAddresses": false,
			  "AllowMultipleEntities": true,
			  "AllowOnlyEmailAddresses": false,
			  "AllUrlZones": true,
			  "MaximumEntitySuggestions": 50,
			  "PrincipalSource": 15,
			  "PrincipalType": 13,
			  "QueryString": UserEmail
			}
		  })
		};
		request(options, function (error, response,body) {
		  if (error) throw new Error(error);
		  var result = JSON.parse(body);
		  //console.log("User Result::"+ result['d']['ClientPeoplePickerResolveUser']);
		  var newTemp = (result['d']['ClientPeoplePickerResolveUser']).replace(/"/g, "'");
		  //console.log(response.body);
		  
		  
		  // Share the Folder with User
			var options = {
			  'method': 'POST',
			  'url': SiteURL+'/_api/SP.Web.ShareObject',
			  'headers': {
				'contentType': 'application/json;odata=verbose',
				'Authorization': 'Bearer '+token,
				'accept': 'application/json;odata=verbose',
				'Content-Type': 'application/json'
			  },
			  body: JSON.stringify({
				"url": SiteURL+"/"+documentLibraryName+"/"+folderName, //"https://spo365devfactory.sharepoint.com/Shared%20Documents/Test113",
				"peoplePickerInput": "["+newTemp+"]",
				"roleValue": "",
				"propagateAcl": false,
				"sendEmail": true,
				"includeAnonymousLinkInEmail": true,
				"emailSubject": "A document folder has been shared to you",
				"emailBody": "Email Body Desc"
			  })

			};
			request(options, function (error, response) {
			  if (error) throw new Error(error);
			  console.log("Permission Granted for that folder");
			});

		});
});



