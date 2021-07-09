# Power Platform Build Tools Hands-on Lab
## Module 2: Automating Solution Deployment Using Azure DevOps
Last updated: February 15, 2021
Authors: Per Mikkelsen, Shan McArthur, Evan Chaki, Amin Espinoza

## Lab Scenario
In this hands-on lab, you will create an Azure DevOps Project, setup permissions and create the required pipelines to automatically export your app (as unmanaged solution) from a development environment, generate a build artifact (managed solution) and finally deply the app into production. The lab will also introduce you to Microsoft Power Platform Build Tools.

## Create an Azure DevOps Project
We are going to use Azure DevOps for both source code storage, build and deployment automation. You can use any automated source control and build automation tools using the same principles. Your configuration data should be exported from a development environment, processed by Package Deployer, and checked into source control.

1. Log into [dev.azure.com](https://dev.azure.com/) with your credentials and click on your organization in the left panel. Follow the instructions to create a project.

![DevOps](./assets/module2/img0.jpg)

2. When the project is completed, you will need to create a Repo to hold the source code. Click on the **Repos** link in the left navigation.

![Repos](./assets/module2/img1.jpg)

3. Initialize the repo with the default README by clicking the **Initialize** button.

![README](./assets/module2/img2.jpg)

### Install and enable the Azure DevOps Extensions for Power Platform

4. On the upper right corner close to your profile icon, you can see a **bag icon**, click it and then select the option **Browse marketplace**.

![Extensions](./assets/module2/img3.jpg)

5. This will redirect you to the Azure DevOps marketplace. Search for **Power Platform** and select the **Power Platform Build Tools**. Note that there are great community options too.

![Marketplace](./assets/module2/img4.jpg)

6. Click **Get it free**.

![Install](./assets/module2/img5.jpg)

7. Click **Install**.

![Tools](./assets/module2/img6.jpg)

8. Click **Proceed to organization**.

![Organization](./assets/module2/img7.jpg)

### Configure Azure DevOps Permissions for Build Service Account

The build pipelines that we set up later will be exporting files from an org and checking them into your source code repo. This is not a default permission of Azure DevOps, so we need to configure the appropriate permissions for the pipeline to function properly.

9. Click the **Project Settings** icon on the lower left of the screen and click the **Repositories** link int the fly out menu. Select the repository in your project and go to the **Permissions** tab. Type in **Project Collection Build Service** and select it (select the project collection with the username appended).

![Permissions](./assets/module2/img8.jpg)

10. The setting for the user is displayed. Select Allow for **Contribute** and ensure that the green checkbox is displayed.

![Contributor](./assets/module2/img9.jpg)

### Build Pipeline 1: Create Export from Dev

The first pipeline you wil create will export your solution from your development environment as an unmanaged solution, unpack it and check it into source control (your repo).

11. Click the **Create Pipeline** button.

![Pipeline](./assets/module2/img10.jpg)

12. Click **use the classic editor** link to create a pipeline.

![Pipeline](./assets/module2/img11.jpg)

13. Leave default values as is and click **Continue**.

![Values](./assets/module2/img12.jpg)

14. Select **Empty job** to create an empty job.

![Job](./assets/module2/img13.jpg)

15. Select **Agent job 1** and enable the **Allow scripts to access the OAuth token** option.

![Job](./assets/module2/img14.jpg)

16. Name the pipeline **Create Export from Dev** and click Save. Changes to your pipeline are also checked into source control so you will be prompted to select the folder and type a comment before saving.

![Name](./assets/module2/img15.jpg)

17. You are now ready to start using the Power Platform Build Tools tasks for Azure DevOps in your pipeline. Select **Add a task to Agent job 1**.

![Agent](./assets/module2/img16.jpg)

18. The very first task you will create is a task that installs the required tools in the agent. Search for **Power Platform Tool**

![Install](./assets/module2/img17.jpg)

19. Select **Add** under the Power Platform Tool Installer task.

![Add](./assets/module2/img18.jpg)

20. This will add the Tool Installer task to the Pipeline. No additional configuration is required for this task.

![Task](./assets/module2/img19.jpg)

### Export Solution as Unmanaged

Your unmanaged solution is your source code for your configuration. You should export your solution as unmanaged to check it into source control. 

21. The next task you will add is a task that exports the unmanaged solution from development environment. Search for **Power Platform Export Solution**.

![Task](./assets/module2/img20.jpg)

22. Select the task and click **Add**.

![Export](./assets/module2/img21.jpg)

23. After adding the Power Platform Export Solution task, you will notice that additional configuration is required. Click on the Power Platform Export Solution (in the pipeline view).

![Export](./assets/module2/img22.jpg)

24. This will open the task configuration page:

* **Display name**: Is inherited from Build task itself.
* **Application type**: Two types of authentications are available:
    * **Username/password**: Simple to setup but does not support multi factor authentication (MFA). This is what we will use for this lab.
    * **Service Principal/client secret**: Recommended and supports MFA. This is harder to setup as it requires creation of the Service principal and client secret in the Azure Portal as well as creation of the application user in the Power Platform environment. Not used in this lab but anyone familiar with setting this up can use this option instead throughout the rest of the lab.
* **Service Connection**: This is the connection to the environment that you want to export the solution from. We will define this in the next step.
* **Solution name**: This is the mane of the solution you want to export.
* **Solution output file**: This specifies the path and filename of the generated solution zip.
* **Export as Managed solution**: By default, solutions are exported as unmanaged (for development purposes). Setting this flag exports the solution as Managed (used fro deployment to any downstream environment such as Test, Pre-prod, and Production).

25. Select **Username/password** under **Authentication type**. Click on **Manage** next to Service connection.

![Export](./assets/module2/img23.jpg)

26. This will bring up the connection configuration required to export the solution from the development environment (user-xx-dev). Click **Create service connection**.

![Service](./assets/module2/img24.jpg)

27. This will bring up the New Service connection screen. Select **Generic** from the list of connections and click **Next**.

![Generic](./assets/module2/img25.jpg)

28. Fill the required details:

    * **Server URL**: https://<environment-url>.crm.dynamics.com. This should be the URL to your development environment.
    * **Username**: Username of a user with admin access to the environment.
    * **Password**: Associated password for the user.
    * **Service Connection Name**: This name will be used to identify which environment to connect to the Build tasks.
    * **Description**: Give the connection a description that helps you identify the environment the service connection connects to.

A sample connection is shown below.

![Service](./assets/module2/img26.jpg)

29. Click **Save**. You will be in the pipeline service connection area in your project.
30. Close the browser tab and go back to the previous tab where tou were building the pipeline. Select the Service Connection that created in the previous step.

![Manage](./assets/module2/img27.jpg)

31. Fill out the remaining details:
    * Solution Name: **$(SolutionName)
      **Note**: This will use the input parameter that you specify when running (queuing) the build pipeline.
    * Solution Output File: **$(Build.ArtifactStagingDirectory)\$(SolutionName).zip**
      **Note** This will add the file to your repo and retain the existing solution name.

![Solution](./assets/module2/img28.jpg)

32. Save the Build Pipeline by clicking **Save & Queue**, then save from the top command bar and clicking **Save** on the dialog.

### Unpack Solution

The solution file that is exported from the server is a zip file with consolidated configuration files. These initial files are not suitable for source code management as they are not structured to make it feasible for source code management systems to properly do differencing on the files and capture the changes you want to commit to source control. You need to **unpack** the solution files to make then suitable for source control storage and processing.

33. Click **Ad a task**, then search for **Power Platform Unpack**.

![Unpack](./assets/module2/img29.jpg)

34. Add the Power Platform Unpack Solution task to the pipeline.

![Add](./assets/module2/img30.jpg)

35. Open the task to configure settings for the task with the following details:

* **Solution Input File**: $(Build.ArtifactStagingDirectory)\$(SolutionName).zip
* **Target folder to unpack solution**: $(Build.SourcesDirectory)\$(SolutionName)
* **Type of sulution**: Unmanaged.

![Unpack](./assets/module2/img31.jpg)

36. Save the updated pipeline.

![Save](./assets/module2/img32.jpg)

37. You can leave the comment blank.

![Comment](./assets/module2/img33.jpg)

### Commit Solution Files to Source Control

Next, we are going to add some scripts as a next pipeline step to commit the solution to the repo. The reason we allowed for scripts to access the OAuth token when we started building the pipeline was to allow for this next step.

38. Add a Command Line task to your pipeline by clicking the **+** button, searching for **command**, and adding it.

![Command](./assets/module2/img34.jpg)

39. Select the task, give the task a display name and copy the following script into the script textbox:

```bash
echo commit all changes
git config user.email "userXXX@wrkdevops.onmicrosoft.com"
git config user.name "Automatic Build"
git checkout main
git add --all
git commit -m "solution init"
echo push code to new repo
git -c http.extraheader="AUTHORIZATION: bearer $(System.AccessToken)" push origin main
```

40. Screenshot provided below:

![Script](./assets/module2/img35.jpg)

41. Save the task (Save & queue > Save).

### Test your Pipeline

42. Select **Queue** to execute your pipeline.

![Queue](./assets/module2/img36.jpg)

43. Leave the defaults and click **Run**.

![Run](./assets/module2/img37.jpg)

44. You can see your build has been queued. You can navigate to it directly in the notification on this screen or using the **Builds** area in the left navigation.

![Task](./assets/module2/img38.jpg)

45. The desired outcome would be a series of green checkboxes, but if you have some issues with your pipeline, you should see the errors in the log. You can click on those for more details if you like. In our case, we have not defined our solution name variable yet, so the export step will fail. This was done intentionally to show you what a failure looks like and how to look at the details.

![Fail](./assets/module2/img39.jpg)

46. Go back to your pipeline by selecting the arrow next to the pipeline name on the top of the screen.

![Back](./assets/module2/img40.jpg)

47. Select the three dots button and then **Edit pipeline**.

![Edit](./assets/module2/img41.jpg)

48. On the pipelines screen, switch to the Variables tab and click **Add** to add a new variable.

49. Use **SolutionName** for the name of the variable and add your unique solution name as a value. It should be your solution's name form Module1 (point 3).

![Edit](./assets/module2/img42.jpg)

50. Click **Save and queue** and observe the results. Fix any issues until you get a successful build.

![Success](./assets/module2/img43.jpg)

**Note** You can get more details about any issues found by setting the variable **debug** to **true**.

![Success](./assets/module2/img44.jpg)

51. Let's go look at what it added to your source control now. Click **Repos and Files** and notice that it added a new folder that contains your solution files.

![Repos](./assets/module2/img45.jpg)

## Pipeline 2: Build your Managed Solution

Your unmanaged solution is checked into source control, but you need to have a managed solution to deploy to other environments such as test and production. To obtain your managed solution to deploy to other environments such as test and production. To obtain your managed solution, you should ue a **Just-In-Time**(JIT) build environment where you would import your unmanaged solution files then export them as managed. These managed solution files will not be checked into source control buy will be stored as a build artifact in your pipeline. This will make them available to be deployed in your release pipeline.

52. Navigate to your pipelines and add a new build pipeline.

![Repos](./assets/module2/img46.jpg)

53. Click **Use the classic editor** link to build the pipeline without YAML. On the next screen use the defaults and click **Continue**. Finally, click on the **Empty job**link on the template selection page shown below.

![Job](./assets/module2/img47.jpg)

54. Name the pipeline **Build Managed Solution** and Save it. This will give you an empty pipeline.

![Job](./assets/module2/img48.jpg)

55. A small note on navigation. If you have problems seeing your newly created pipeline, you can see all the pipelines you have created in the **All** tab.

![Pipelines](./assets/module2/img49.jpg)

56. In your empty pipeline, add the Power Platform Tool Installer task.

![Add](./assets/module2/img50.jpg)

57. Add the Power Platform Pack Solution task.

![Add](./assets/module2/img51.jpg)

58. Configure the pack solution task with the following parameters:

* **Source folder of Solution to Pack**: $(Build.SourcesDirectory)\$(SolutionName)
* **Solution Output File**: $(Build.ArtifactStagingDirectory)\$(SolutionName).zip

![Add](./assets/module2/img52.jpg)

59. On the pipeline screen, switch to the variables tab in the pipeline and add the SolutionName variables with your unique solution name. This should be your solution name from Module 1. Click **Save** (through the **Save & Queue dropdown).

![Save](./assets/module2/img53.jpg)

60. The next task will be to import the solution into your build server, so we need to add a connection to that environment before we add and configure the task to do the import. Click the **Project settings** link in the lower left of the screen and navigate to **Service connections** then click **New service connection**.

![Connection](./assets/module2/img54.jpg)

61. Select **Generic** on the New Service connection page and click **Next**.

![Connection](./assets/module2/img55.jpg)

62. Fill out the required details:

* **Server URL**: https://<environment-url>.crm.dynamics.com. This should be the URL to your build environment.
* **Username**: Username of a user with administrator access to the environment.
* **Password**: Associated password for the user
* **Service connection name**: This name will be used to identify which environment to connect to in the Build tasks.
* **Description**: Give the connection a description that helps you identify the environment the service connection connects to.

The sample connection is shown below.

![Generic](./assets/module2/img56.jpg)

**Note**: If you forgot to record the server URL for your build environment, simple visit [https://admin.powerplatform.microsoft.com/](https://admin.powerplatform.microsoft.com/) click the environment. This will open another tab displaying the environment URL. Make sure it is your build environment.

![Environments](./assets/module2/img57.jpg)

![Environments](./assets/module2/img58.jpg)

63. Click **Save** You will be in the pipeline service connections area in your project.

64. Navigate back to your Build Management Solution pipeline (if you don't see the pipeline go to step 55 on how to view all the pipelines). Click **Edit**, and then add the **Power Platform Import Solution** task to take the solution file and import into your build environment.

![Import](./assets/module2/img59.jpg)

65. Select the connection you just added to your build server and provide the following for the solution input file: $(Build.ArtifactStagingDirectory)\$(SolutionName).zip.

![Select](./assets/module2/img60.jpg)

66. Add a task to export the solution file as a managed file.

![Add](./assets/module2/img61.jpg)

67. Select your build environment, check the **Export as managed solution** checkbox and enter the following details:

* **Solution name**: $(SolutionName)
* **Solution Output File**:  $(Build.ArtifactStagingDirectory)\$(SolutionName)_managed.zip

![Add](./assets/module2/img62.jpg)

68. The files at this stage will be in the build agent and we need to publish it as a build artifact. Add a new pipeline step **Publish Pipeline Artifact**.

![Publish](./assets/module2/img63.jpg)

69. Leave the predefined values as-is and click **Save and queue**:

![Publish](./assets/module2/img64.jpg)

70. Monitor it until it is successful.

![Monitor](./assets/module2/img65.jpg)

## Release Pipeline: Release to Production

71. The next task will be to import the solution into your production server, so we need to add a connection to that environment before we add and configure the task to do the import. Click the **Project seetings** link in the lower left of the screen and navigate to **Service connections**, then click *New Service Connection**.

![Connection](./assets/module2/img66.jpg)

72. Select **Generic** on the New service connection page and click **Next**.

![Connection](./assets/module2/img55.jpg)

73. Fill out the required details:

* **Server URL**: https://<environment-url>.crm.dynamics.com. This should be the URL to your build environment.
* **Username**: Username of a user with administrator access to the environment.
* **Password**: Associated password for the user
* **Service connection name**: This name will be used to identify which environment to connect to in the Build tasks.
* **Description**: Give the connection a description that helps you identify the environment the service connection connects to.

The sample connection is shown below.

![Connection](./assets/module2/img67.jpg)

74. Click **Save**. You will be in the pipeline service connections area in your project.

75. To deploy a build, you will configure a release pipeline. Navigate to the release pipelines. You likely won’t have one yet, so click on the **New pipeline** button.

![Release](./assets/module2/img68.jpg)

76. Click on *Empty job** when selecting a template.

![Empty](./assets/module2/img69.jpg)

77. Give it a stage name and then close the stage dialogue.

![Stage](./assets/module2/img70.jpg)

78. Next, we need to add the artifact that will be used in the deployment. Click the **Add** button in the Artifacts list. In the **Add an artifact** screen, select your **Build Managed Solution** from the source build pipeline. To make it easier to not have to deal with encoding spaces in filenames, change your source alias to something simple, like **Build**. Leave the rest as default. Click **Add** when completed.

![Artifact](./assets/module2/img71.jpg)

79. Switch to the **Tasks** view and click the plus button to add a new task.

![Task](./assets/module2/img72.jpg)

80. In the tasks view, click the **+** button to add a new task.

![Add](./assets/module2/img73.jpg)

81. Add the Power Platform Tool Installer task.

![Install](./assets/module2/img74.jpg)

82. Add the **Power Platform Import Solution** task to the pipeline.

![Import](./assets/module2/img75.jpg)

83. Configure the solution import to use the production environment connection and the following for solution input file: 

```
$(System.DefaultWorkingDirectory)/Build/drop/$(SolutionName)_managed.zip
```
Click **Save**.

![File](./assets/module2/img76.jpg)

84. Switch to the **Variables** tab and add a variable named **SolutionName** and the unique name of the solution to import. If you are using the solution created in lab 1, the solution name should be that. Click **Save** an then **Ok** in the next dialog.

![Save](./assets/module2/img77.jpg)

85. Click the **Create release** button. This functions like Queuing a build pipeline, only for release pipelines.

![Release](./assets/module2/img78.jpg)

86. Click **Create** on the popup panel.

![Create](./assets/module2/img79.jpg)

87. Monitor your release until it is succeeded, or troubleshoot if it is not successful.

![Succeed](./assets/module2/img80.jpg)

88. For final confirmation, log into your production system and see your application!

![App](./assets/module2/img81.jpg)

### Terms of use

© 2020 Microsoft Corporation. All rights reserved.
By using this demo/lab, you agree to the following terms: The technology/functionality described in this demo/lab is provided by Microsoft Corporation for purposes of obtaining your Power Platform ALM Workshop: Module 2 66 feedback and to provide you with a learning experience. You may only use the demo/lab to evaluate such technology features and functionality and provide feedback to Microsoft. You may not use it for any other purpose. You may not modify, copy, distribute, transmit, display, perform, reproduce, publish, license, create derivative works from, transfer, or sell this demo/lab or any portion thereof. COPYING OR REPRODUCTION OF THE DEMO/LAB (OR ANY PORTION OF IT) TO ANY OTHER SERVER OR LOCATION FOR FURTHER REPRODUCTION OR REDISTRIBUTION IS EXPRESSLY PROHIBITED. THIS DEMO/LAB PROVIDES CERTAIN SOFTWARE TECHNOLOGY/PRODUCT FEATURES AND FUNCTIONALITY, INCLUDING POTENTIAL NEW FEATURES AND CONCEPTS, IN A SIMULATED ENVIRONMENT WITHOUT COMPLEX SET-UP OR
INSTALLATION FOR THE PURPOSE DESCRIBED ABOVE. THE TECHNOLOGY/CONCEPTS REPRESENTED IN THIS DEMO/LAB MAY NOT REPRESENT FULL FEATURE FUNCTIONALITY AND MAY NOT WORK THE WAY A FINAL VERSION MAY WORK. WE ALSO MAY NOT RELEASE A FINAL VERSION OF SUCH FEATURES OR CONCEPTS. YOUR EXPERIENCE WITH USING SUCH FEATURES AND FUNCTIONALITY IN A PHYSICAL ENVIRONMENT MAY ALSO BE DIFFERENT.

### FEEDBACK

If you give feedback about the technology features, functionality and/or concepts described in this demo/lab to Microsoft, you give to Microsoft, without charge, the right to use, share and commercialize your feedback in any way and for any purpose. You also give to third parties, without charge, any patent rights needed for their products, technologies and services to use or interface with any specific parts of a Microsoft software or service that includes the feedback. You will not give feedback that is subject to a license that requires Microsoft to license its software or documentation to third parties because we include your feedback in them. These rights survive this agreement. MICROSOFT CORPORATION HEREBY DISCLAIMS ALL WARRANTIES AND CONDITIONS WITH REGARD TO THE DEMO/LAB, INCLUDING ALL WARRANTIES AND CONDITIONS OF MERCHANTABILITY, WHETHER EXPRESS, IMPLIED OR STATUTORY, FITNESS FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. MICROSOFT DOES NOT MAKE ANY ASSURANCES OR REPRESENTATIONS WITH REGARD TO THE ACCURACY OF THE RESULTS, OUTPUT THAT DERIVES FROM USE OF DEMO/ LAB, OR SUITABILITY OF THE INFORMATION CONTAINED IN THE DEMO/LAB FOR ANY PURPOSE.

### DISCLAIMER

This demo/lab contains only a portion of new features and enhancements in. Some of the features might change in future releases of the product. In this demo/lab, you will learn about some, but not all, new features.
