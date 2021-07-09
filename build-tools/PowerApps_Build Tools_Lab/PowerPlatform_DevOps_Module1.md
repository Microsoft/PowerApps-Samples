# Power Apps Hands-on Lab
## Module 1: Building Basic Model-driven Common Data Service Application for Deployment Using Azure DevOps
Last updated: February 15, 2021
Authors: Per Mikkelsen, Shan McArthur, Evan Chaki, Amin Espinoza

In this lab you will be creating a basic Model-driven Dataverse application to use with Module 2 of the lab to automate the application lifecycle management (ALM) of the application.

1. Navigate to [https://make.powerapps.com](https://make.powerapps.com) and sign in with your credentials. Click the environment selector dropdown the header and select your development environment.

![Environments](./assets/module1/img0.jpg)

2. Click the **Solutions** area in the left navigation then click the **New Solution** button to create a new solution.

![Solutions](./assets/module1/img1.jpg)

3. In the side panel that appears, enter a name for the application and click the dropdown for the publisher and click the **Add Publisher** option.  
**Note**: The solution publisher specifies who developed the app. For this reason, you should create a solution publisher name that's meaningful. Furthermore, the solution publisher includes a prefix, which helps you distinguish system components or components introduced by others and is also a mechanism to help avoid naming collisions. This allows for solutions from different publishers to be installed in an environment with minimal conflicts.

![Publisher](./assets/module1/img2.jpg)

4. A new window will pop up. If a window does not pop up, check your popup blocker settings. For the purpose of this lab, enter your **ALMLab** for the display name, name and prefix and click **Save and Close**.

![New publisher](./assets/module1/img3.jpg)

5. On the new solution panel, select the publisher that you just created, give the application a version number, and click **Create** to create a new unmanaged solution in the environment.

![New solution](./assets/module1/img4.jpg)

6. In the solutions list, select the solution you just created and click the **Edit** button.

![Edit](./assets/module1/img5.jpg)

7. Your new solution will be empty, and you need to add components to it. In this lab we will create a custom entity. Click the **New** dropdown from the top navigation and select Table.

![Table](./assets/module1/img6.jpg)

8. Enter a display name and plural name. The system will fill out the table name and other fields for you. Click **create** to create the table.

![New table](./assets/module1/img7.jpg)

9. Once your table is created, click the solution name again to go back to the solution view to add another component.

10. Click the New **dropdown**, then App, and **Model-driven app**.

![New app](./assets/module1/img8.jpg)

11. Enter an application name and click the **Done** button.

![Create app](./assets/module1/img9.jpg)

12. In the application designer, click the **Site Map** to edit it.

![Site map](./assets/module1/img10.jpg)

13. In the site map editor, select the **New Subarea** to get its current properties.

![Site map](./assets/module1/img11.jpg)

14. Select the Entity dropdown and select your custom table to add to the sitemap.

![Custom table](./assets/module1/img12.jpg)

15. Click **Save**, then **Publish**, then **Save and Close** to go back to the application designer.

![Save](./assets/module1/img13.jpg)

16. Click **Save** then Validate to validate the application.

![Validate](./assets/module1/img14.jpg)

17. You should see one warning. View the warning, then click **Publish** then **Play**.

![Warning](./assets/module1/img15.jpg)

18. This will take you to the application so that you can see how it looks. You can use the application and close the tab when you are satisfied.

![App](./assets/module1/img16.jpg)
