## Model Driven App WebApi Samples
The model-driven-app TypeScript samples are designed to run inside a Model Driven App Webresource and show how to:

- CRUD using Xrm.WebApi
- Calling Bound and Unbound functions using Xrm.WebApi
- Calling Bound and Unbound actions using Xrm.WebApi
- Associating and Disassociating using HTTP WebApi requests (currently not supported by Xrm.WebApi) 
- Special cases for Customer fields and ActivityParties

## Design Principle
The samples have been written with the minimal dependencies and aim to be as self-contained as possible. This means that all typings are in-line rather than as spearate interface definitions.
The exception to this is ```WebApiRequest.ts``` which provides functions to send WebApi HTTP requests for associate/disassociate which is not supported by the ```Xrm.WebApi``` at this time.

## How to build samples
1. Open the folder ```ts-model-driven-app``` in VS Code
2. Select **Terminal** -> **Run Task...** -> **npm: install**
3. Select **Terminal** -> **Run Build Task**

## How to run samples
You do not need to deploy the webresource, but instead use fiddler's autoresponder to load the webresource.
Learn about this technique at https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/streamline-javascript-development-fiddler-autoresponder

1. Install Fiddler
2. Turn on HTTPS Decryption **Tools**->**Options**->**HTTP**->**Decrypt HTTPS traffic**
3. In the **AutoResponder** tab, **Add Rule**

Match:
```
regex:(?insx).+/WebResources/powerappsamples_/(?'fname'[^?]*).*
```
Responds with:
```
C:\Repos\PowerApps-Samples\cds\webapi\ts-model-driven-app\${fname}
```

4. Open your Model Driven App and find the Appid from the url:
```
https://<org>/main.aspx?appid=<APPID>&pagetype=dashboard&id=d201a642-6283-4f1d-81b7-da4b1685e698&type=system&_canOverride=true
```
5. Navigate to the test webresource using the url:
```
https://<org>/main.aspx?appid=<APPID>&pagetype=webresource&webresourceName=powerappsamples_/model-driven-webapi.html&cmdbar=false&navbar=off
```
6. The tests will run. You can press F12 and add breakpoints to examine how the tests work.


The tests use Mocha and Chai for the test framework.