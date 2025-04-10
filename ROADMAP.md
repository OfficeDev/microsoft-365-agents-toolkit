# Teams Toolkit Roadmap

Our roadmap and priorities are shaped by your feedback, issues, discussions, surveys, social media, and market opportunities. We want to make building apps for Teams fast, simple, and delightful. We're constantly learning, and we listen to your feedback and adapt our plans if needed. The items below may not include everything we have planned. If you have feedback or suggestions, create a new GitHub issue, or react with a 👍 on an existing issue.

Teams Toolkit continuously focuses on:

- **Microsoft 365 Platform:** We’re committed to providing Day 1 support for new Microsoft 365 capabilities and features.

- **Support for Familiar Language & Tools:** Teams Toolkit works with the languages, frameworks, tools, and services you know and love.

- **Get Started Fast:** Teams Toolkit is the simplest way to get started building for Microsoft 365 with delightful tooling, samples, and documentation.

- **Development Belocity:** best-in-class productivity to create, build, and deliver your apps.

- **Ship with Confidence:** Teams Toolkit is enterprise-ready to help developers and IT administrators deliver, manage, and monitor apps with confidence.

## Q1-Q2, 2024

- **Build intelligent apps for Microsoft 365 and Copilot platform**: AI exploded in 2023, promising a transformation for the way people communicate, collaborate and work. Teams platform also formed a strategy to create engaging app experiences that leverage AI to ease complex conversational app development.There are 2 major types of experience: You can bring AI to your Teams application (Chat Bot in Teams with Teams AI Library): Teams AI Library enables developers to add natural language processing to their conversational apps and respond to end-user with AI generated content. Or integrate with Copilot (Plugin for Microsoft 365 Copilot): Copilot plugin allows developers to extend Microsoft Copilot via message extensions to access real-time information. 

    - Copilot plugins (Microsoft version of OpenAI plugin) E2E DevX.
    - Custom Copilots (Intelligent chat bots that use Teams AI Library) with AI Agent and RAG scenarios.  

- **AI Enhanced Developer Experience**: 
    - Getting Started with Teams app development using GitHub Copilot Chat. 
    - Help early stage developers navigate through Microsoft 365 platform concepts, terminology and technical stacks. 
    - Facilitate developers move fast in the development journey. 

- **Improve Teams app Testability**: Teams App Test Tool (Test Tool) makes debugging bot-based apps effortless. You can chat with your bot and see its messages and Adaptive Cards as they appear in Microsoft Teams. Now we are expanding its capability to cover Message Extensions. 

- **Streamlined DevX with Revamped Documentation**:
    - Revamp Information Architecture: Restructure the Table of Contents (TOC) in the documentation to improve clarity and navigation. 
    - Incremental Documentation Improvements: Continue to enhance the documentation by completing Visual Studio (VS) documentation, adding a samples overview table, and adding advanced usage tips content to the documentation site. 

- **Move Applications to Production Environment**: Teams Toolkit has been optimizing the experiences for Getting Started scenarios for developers who are new to the platform. There’s another chunk of users who have already invested a lot / familiar with the platform while we haven’t provided a lot of value for them. This Epic intends to optimize the experience for developers when they have completed the development process and want to ship their applications. 
    - Enable shift-left functional app testing for Teams applications.
    - Deploy static contents to Azure Static Web applications. 
    - Improved CI/CD pipeline creation through pipeline templates and tailored guides. 
    - Containerize the Teams application deployment with ACS. 

## Q3-Q4, 2023

- **Extend Microsoft 365 Copilot**: Teams Toolkit for Visual Studio and Visual Studio Code will soon let users quickly build a plugin to extend Microsoft 365 Copilot. This [demonstration](https://www.youtube.com/watch?v=6ZNi1GDxvf0&ab_channel=Microsoft365Developer) shows how simple it will be to build a plugin to extend Microsoft 365 Copilot using the Teams Toolkit extension in Visual Studio Code. This is a preview of the experience you will have when we release our first public preview (coming soon).
- **Continuous improvements of templates and samples**: App templates and samples continue to be the key starting point for developers to learn how to build apps for Teams, Outlook and the Microsoft 365 app. We will continue to simplify the app templates and improve the usability of sample app gallery as we incrementally adding more samples to it.
- **Debug in Teams Desktop clients and emulator**: We are working on enabling debugging Teams apps in Teams Desktop clients and emulator. This will allow developers to iterate on their development faster and without worrying about Microsoft 365 account or sideloading permissions upfront.
- **App scenarios**: We are exploring empowering customer service, frontline workers and IT help desk scenarios in Teams with AI capabilities.
- **AI-assisted developer experience**: Developers care about productivity and we are exploring how to leverage AI to help developers to be more productive in Teams Toolkit.

## Q1-Q2, 2023

- **Flexible and composable automation**: We are designing a new way for developers to write individual, focused tasks that are tuned for automating setup and other repetitive actions during development of Teams apps. Those tasks are composed into groups – referred to as life cycles: Provision, Deploy, and Publish. All that orchestration is controlled by you, the developer, and expressed in a YAML file that lives in the root of your project.
- **SPFx Solutions across Microsoft 365**: We will include support for the latest SPFx version, which means you can start extending SPFx-based solutions to Outlook and the Microsoft 365 app. For getting started developers, we will also include a project template that supports debugging SPFx apps across Microsoft 365.
- **A refreshed getting started experience**: We are working on a new getting started experience that will help you find the right project template by using familiar terminology that matches the rest of the Teams Platform. We will also make it easier for you to customize the templates for your apps by simplifying the scaffolded code. More complex examples will be available as samples.
- **GitHub Codespaces**: we’re designing a few samples that can run directly in GitHub Codespaces. This containerized experience sets up Teams Toolkit for Visual Studio Code in a new browser instance and will be ready to run after you sign-in with your Microsoft 365 account. This will be a great way to get started with Teams Toolkit without installing anything locally.

## Q3-Q4, 2022

- **Teams Toolkit General Availability**: We’re planning to make Teams Toolkit for Visual Studio generally available later this year in Visual Studio 2022.
- **Free, secure, integrated network tunneling**: We know how challenging it can be to debug different capabilities and services with your Teams apps and that it’s not always possible to use existing options for network tunneling in your enterprise environments. To help with this, Teams Toolkit will use new tooling that provides a network tunnel that allows secure communication only from Microsoft Teams.
- **Mobile preview**: Teams Toolkit will continue to make running your apps simple including on mobile devices. We’re bringing support for launching to Teams and Outlook mobile first, and Office mobile will follow later this year.  
- **Office add-ins**: We’re collaborating with the Office add-ins team to make it simpler to create add-ins with a single set of tools using Teams Toolkit. This will bring the same quick getting started and simple run and debug experience we’ve made for Teams to Office add-ins.
- **Yeoman generator support**: One of our themes this year is to support familiar tooling, and to continue this theme we’ll make it simpler to work with the popular tooling that you’re using to create apps for M365. Based on your feedback, we’re starting with the Yeoman generators for Microsoft Teams Apps and the SharePoint Framework so that these projects can easily use Teams Toolkit to run and debug their apps.
- **Deployment**: We’ve heard your feedback on the need for more flexible provisioning and deployment configurations. We’ll make it simpler to integrate with your enterprises’ existing cloud deployment process.
- **App scenarios**: Teams Toolkit is a terrific way to get started with Teams app development and we want to continue this theme and help solve common business problems with scenario-focused solutions. We’re starting with a few templates that help jump start your bot apps. We’ll add support for additional scenarios based on your feedback.

Not seeing what’s most important to you and your business? File [a feature request](https://github.com/OfficeDev/TeamsFx/issues/new/choose) and let us know!
