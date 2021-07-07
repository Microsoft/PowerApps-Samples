import * as PowerAppsAuthoringSdk from "@microsoft/powerappsauthoringsdk";

let makerSession: PowerAppsAuthoringSdk.MakerSession = {} as PowerAppsAuthoringSdk.MakerSession;
let appId: string = "";

export function initSdkAndCreateApp() {
    // Check if SDK is initialized then play the App else initialize the SDK and play App
    if (PowerAppsAuthoringSdk.isSdkInitialized()) {
        createApp();
    }
    else {
        initSdk().then(function () {
            createApp();
        });
    }
}

export function initSdkAndEditApp() {
    // Check if SDK is initialized then play the App else initialize the SDK and play App
    if (PowerAppsAuthoringSdk.isSdkInitialized()) {
        editApp();
    }
    else {
        initSdk().then(function () {
            editApp();
        });
    }
}

// Initialze Authoring  SDK
async function initSdk(): Promise<void> {
    let options: PowerAppsAuthoringSdk.SdkInitializer = {
        hostName: "Authoring SDK sample"
    }
    return PowerAppsAuthoringSdk.initAsync(options).then(function () {
        return Promise.resolve();
    }).catch((error: Error) => {
        return Promise.reject(error.message);
    });
}

function createApp(): void {
    // Initializing the options that can be passed to the App
    // Other parameters like solutionId and Template can also be set while creating an App
    const options: PowerAppsAuthoringSdk.CreateAppOptions = {
        displayName: "Authoring SDK sample",
        formFactor: PowerAppsAuthoringSdk.FormFactor.Tablet
    }

    // Creating an App
    PowerAppsAuthoringSdk.MakerSession.createAppAsync(options).then(function (session) {
        makerSession = session;
        // Set schema on the App
        setSchema();

        // Set data on the App being edited
        setData();

        // Set host definition on the App being edited
        setHostMethods();

        // Subscribe to the events
        subscribeToMakerSessionEvents();
    })
}

function subscribeToMakerSessionEvents() {
    makerSession.appSaved.subscribe((appInfo) => {
        appId = appInfo.appId;
        console.log("App is saved. AppId: " + appInfo.appId);
    });

    // This event can be used to play the App using the Player SDK using the AppId or LogicalName   
    makerSession.appPublished.subscribe((appInfo) => {
        appId = appInfo.appId;
        console.log("App is published. AppId: " + appInfo.appId);
    });

    // App signalled to the host via a host definition WADL set
    // @ts-ignore TODO: remove once upstream Authoring SDK has event merged
    makerSession.appSignalled.subscribe((appInfo) => {
	console.log("App signalled host. Info: " + JSON.stringify(appInfo));
    });
}

function editApp(): void {
    const appIdPrefix: string = "/providers/Microsoft.PowerApps/apps/";
    appId = appId ? decodeURIComponent(appId).replace(appIdPrefix, "") : appId;
    // Initializing the options that can be passed to the App    
    const options: PowerAppsAuthoringSdk.EditAppOptions = {
        appId: appId
    }

    // Edit an App
    PowerAppsAuthoringSdk.MakerSession.editAppAsync(options).then(function (session) {
        makerSession = session;
        subscribeToMakerSessionEvents();
    }).catch((error) => {
        console.log(error);
    });
}

function setSchema() {
    // Initializing schema
    const defaultSchema: PowerAppsAuthoringSdk.PowerAppsSchema = {
        Name: PowerAppsAuthoringSdk.DataType.String,
        Age: PowerAppsAuthoringSdk.DataType.Number
    };

    // Set schema on the App that is being edited
    // This will override the previous schema that was set
    // Use "hostName.Data" to access this object in the app. 
    // hostName is set while initializing the SDK
    makerSession.setSchemaAsync(defaultSchema).then(function () {
        console.log("Schema is set on the App")
    });
}

function setData() {
    // This allows setting the data on the App that is being edited
    // This should be called after set schema and should conform to the schema which was set
    // This will override the data that was set previously provided
    // This data is not saved with the App. 
    // Use setData on PowerApps Player SDK to set the data on the app while playing it.
    const data: PowerAppsAuthoringSdk.PowerAppsData = [
        {
            Name: "Test User 1",
            Age: 10
        },
        {
            Name: "Test User 2",
            Age: 30
        }]

    makerSession.setDataAsync(data).then(function () {
        console.log("Data is set on the App")
    });
}

// Set schema by passing in a WADL XML. 
function setHostMethods() {
    // id ="ChangeHostBackground" indicated the host function
    // User can call this function in the App like : Authoring SDK sample.ChangeHostBackground(param)
    // Title indicates the intellisense to be displayed to the user while editing the App
    // Refer to the link - https://www.w3.org/Submission/wadl/ to understand more about WADL
    const callback = "<application xmlns:xml=\"http://www.w3.org/XML/1998/namespace\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:service=\"https://example.com/apim\" xmlns:siena=\"http://schemas.microsoft.com/MicrosoftProjectSiena/WADL/2014/11\" siena:serviceId=\"Authoring SDK sample\" xmlns=\"http://wadl.dev.java.net/2009/02\">\r\n    <doc title=\"Authoring SDK Sample Button\" /> \r\n    <resources base=\"https://example.com/apim\"> \r\n        <resource path=\"/ChangeHostBackground/{hexColor}\"> \r\n            <param style=\"template\" name=\"hexColor\" type=\"xs:string\" required=\"true\" /> \r\n            <method siena:requiresAuthentication=\"false\" name=\"POST\" id=\"ChangeHostBackground\" siena:isHostFunction=\"true\"> \r\n                <doc title=\"Updates background color of Authoring SDK sample.\" /> \r\n                <request> \r\n                    <representation mediaType=\"application/json\" /> \r\n                </request> \r\n            </method> \r\n        </resource> \r\n    </resources> \r\n</application>";

    makerSession.setHostMethodDefinitionAsync(callback).then(function () {
        console.log("Host definition is set on the App")
    });;
}
