---
page_type: sample
languages:
- csharp
products:
- office-teams
description: Expert Finder bot allows users to search for experts based on certain attributes
urlFragment: microsoft-teams-apps-expertfinder
---

# Expert Finder Bot App Template
| [Documentation](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki) | [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Deployment-Guide)| [Architecture](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Solution-Overview)
|--|--|--|

Expert Finder bot allows employees to search for other individuals in an organization based on their skills, interests and schools. In addition, it also provides users the ability to update their profile information and keep it up to date.

**Expert Finder bot**
 - **My Profile**: Using this command users will be able to view their Azure Active Directory profile information in a card. The card will have call to action buttons that will let them add or modify information in their Azure Active Directory profile or view details about other attributes.
 
![MyProfileCard](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Images/MyProfileCard.png)

![EditProfile](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Images/EditProfile.png)

 - **Search**: This command allows the users to search for experts within the organization whose attributes match with the search keyword. They will be able to select a max of 5 user profiles and view details pertaining to them.

![SearchCard](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Images/SearchFeature.png)

![SearchTaskModule](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Images/SearchTaskModule.PNG)
 
 **Expert Finder messaging extension**
Users can search for individuals within the organization whose attributes match the user search keyword using the messaging extension.

![MessagingExtension](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Images/MessagingExtension.PNG)

## Legal Notices
Please read the license terms applicable to this template [here](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/blob/master/LICENSE). In addition to these terms, you agree to the following:

 - You are responsible for complying with privacy and security regulations applicable to your app.
 
 - Use, collection, and handling of any personal data by your app is your responsibility alone.  Microsoft will not have any access to data collected through your app.  Microsoft will not responsible for any data related incidents or data subject requests.
 
 - If your app is developed to be sideloaded internally within your organization, you agree to comply with all internal privacy and security policies of your organization.
 
 - Use of this template does not guarantee acceptance of your app to Teams app store.  To make this app available in the Teams app store, you will have to comply with [submission process and validation](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish), and all associated requirements such as including your own privacy statement and terms of use for your app.
 
 - Any Microsoft trademarks and logos included in this repository are property of Microsoft and should not be reused, redistributed, modified, repurposed, or otherwise altered or used outside of this repository.

## Getting Started
Begin with the [Solution overview](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Solution-Overview) to read about what the app does and how it works.

When you're ready to try out Expert Finder bot, or to use it in your own organization, follow the steps in the [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/wiki/Deployment-Guide).

## Feedback
Thoughts? Questions? Ideas? Share them with us on [Teams UserVoice](https://microsoftteams.uservoice.com/forums/555103-public)!

Please report bugs and other code issues [here](https://github.com/OfficeDev/microsoft-teams-apps-expertfinder/issues/new).

## Contributing
This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
