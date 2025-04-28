# SharePoint Embedded Copilot Sample App

## Summary

This template showcases a sample SharePointEmbedded Copilot app. This template allows developers to explore the integration of [SharePoint Embedded Copilot](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/declarative-agent/spe-da) with a SPE app. The sample app itself has a React frontend with an Azure Functions backend and provides an end to end example of adding the Copilot chat controller to the frontend, how to customize zeroQueryPrompts and the meta prompt for Copilot.

## Prerequisites

- An app registration in AAD with the right SPE related scopes. [See learning series](https://learn.microsoft.com/en-us/training/modules/sharepoint-embedded-setup/) for guide to set this up.

- A registered container type in your dev tenant. [See learning series](https://learn.microsoft.com/en-us/training/modules/sharepoint-embedded-setup/) for guide to set this up.

- Ensure that the `DiscoverabilityDisabled` flagt is set to false for your container type. [Refer to docs](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/declarative-agent/spe-da-adv#discoverabilitydisabled) for guide to set this up.

- Ensure that Content-Security-Policies for embedded chat hosts has been set up correctly. [Refer to docs](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/declarative-agent/spe-da-adv#csp-policies) for guide to set this up.

## Quick Start

To run the SPE Copilot template in your local dev machine, you will need to complete the following steps:

- Navigate to the `./function-api` subdirectory and create a new `local.settings.json` file. Copy the contents of `local.settings.template.json` and replace the placeholder values with values from the prerequisites.

- Navigate to the `./react-client` subdirectory and create a new `.env` file. Copy the contents of `.template.env` and replace the placeholder values with values from the prerequisites.