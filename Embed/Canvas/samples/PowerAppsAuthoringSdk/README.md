# PowerApps Embedding SDK 

This sample demonstrates the capabilities of PowerApps embedding SDKs and shows you how PowerApps can help customers extend applications with experiences ranging from creating simple custom forms to adding a feature-rich screens. 

## What this sample does

This sample uses the PowerApps embedding authoring SDK to create or edit a new canvas app from the host.

The PowerApps Authoring SDK enables developers to:
- Launch PowerApps Studio to allow the user to create or edit canvas apps.
- Specify the form-factor (tablet or mobile) for a new application along with the predefined templates.
- Subscribe to events from the Studio or Maker (like Studio Launched, App Saved, App Published, App Signalled) that describe the user’s interaction with the app and provides information (i.e., the application id) about the created app.
- Set schema for the data which can be passed to the application during authoring/runtime.
- Set data on the application during authoring time, this data shows up during authoring time but is not saved with application.
- Set host method definition which allows the application to callback into the host.

## How this sample works

To simulate creating or editing a canvas app from the host via the sample, you need to do the following steps:

### Setup

1. Download or clone the repository so that you have a local copy.

## How to run this sample

1. Open Command Prompt and navigate to the folder where the sample code is cloned.
2. Run `npm install`
3. Run `npm run build`
4. Run `npm start`
5. Navigate to [http://localhost:9000/](https://localhost:9000)
  
  
