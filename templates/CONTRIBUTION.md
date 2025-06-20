# Contributing Guide: Adding a New Template

This guide outlines the process of adding a new template to Microsoft 365 Agents Toolkit (ATK).
![image](https://github.com/user-attachments/assets/acf13492-151c-46a7-9f41-b4efcf1d587a)


## Table of Contents
- [Overview](#overview)
- [Project Structure](#project-structure)
- [Step-by-Step Guide](#step-by-step-guide)

## Overview

Microsoft 365 Agents Toolkit follows a predefined scaffolding process that consists of:
1. Question Model - Handles user interaction
2. Template Selection - Maps user choices to templates
3. Generator - Processes templates and generates files

Below is an example illustrating the mapping between the question model (app category question tree), template names, and generators. Generators are components responsible for preprocessing, template retrieval, and post-processing. Leaf nodes corresponding to specific templates are labeled with their respective template names.
![image](https://github.com/user-attachments/assets/68927e4e-8365-446b-b675-d48de107e3aa)

## Project Structure

A typical ATK project includes:

- `.vscode/`: Debugging profiles
- `appPackage/`: Manifest files defining the app
- `infra/`: Bicep infrastructure code for Azure resources
- `m365agents.yml` & `m365agents.local.yml`: Action definitions for ATK commands
- `env/`: Environment files for ATK provisioning and deployment
- Other files: User-generated source code

## Step-by-Step Guide

### 1. Add Project Template

Add your project template to the `templates/<language>/<template-name>` directory. Templates should:
- Be organized by programming language
- Use placeholder variables for runtime replacement
- Support immediate local debugging without extra configurations

### 2. Update Template Metadata

1. Add your template's metadata to files in `packages/fx-core/src/component/generator/templates/metadata/`:
   - For templates using the default generator: Add to the corresponding project type file (e.g., `bot.ts`)
   - For templates requiring custom generation: Add to `special.ts` or create a new file

2. Include template ID, name, programming language, and description

3. Connect the template to the question model via the `data` field in the selection options
Example of a question selection option item with a `data` field linking to the default-bot template:
![image](https://github.com/user-attachments/assets/ab57c63a-f250-43d0-9440-3cac6f01ae63)


### 3. Update Question Model

1. Modify the scaffold question flow based on your extension:
   - For VS Code extension/CLI: Edit `scaffoldQuestionForVSCode()` function in `packages/fx-core/src/question/scaffold/vsc/createRootNode.ts`
   - For Visual Studio extension: Edit `packages/fx-core/src/question/scaffold/vs/createRootNode.ts`

The root tree comprises three children:

1. **Project type sub-tree**: Defines the app type with subcategories like Declarative Agent, Custom Engine Agent, Bot, Tab, Message Extension, and Office Add-in. Traversing this sub-tree determines the template name. Partners can extend it with additional subtypes and templates.

2. **Programming language node**: Queries the template's programming language. Some templates support multiple languages (e.g., `basic-bot` offers TypeScript and JavaScript), while others may not specify a language, indicating they contain only essential configuration files for the tooling. Available options depend on the previously selected template.

3. **Folder path and app name sub-tree**: A group node containing questions about the target folder path and app name, triggered based on specific conditions.

### 4. Use a custom generator (optional)

As previously discussed, if your project requires runtime processing, you'll need to develop a custom generator. ATK provides a default generator base class, `DefaultTemplateGenerator`, with three APIs for customization.

Place your generator in a subdirectory of [packages/fx-core/src/component/generator/](https://github.com/OfficeDev/microsoft-365-agents-toolkit/blob/dev/packages/fx-core/src/component/generator/). For example, [packages/fx-core/src/component/generator/officeAddin/generator.ts](https://github.com/OfficeDev/microsoft-365-agents-toolkit/blob/dev/packages/fx-core/src/component/generator/officeAddin/generator.ts):

Partner can customize the generator by the following APIs:

- `activate()`: Determines whether the generator activates, typically based on supported template names.
- `getTemplateNames()`: Performs preprocessing and maps user inputs to template names, returning a replacement map for resolving placeholders in template files.
- `post()`: Executes post-processing steps.

After defining a custom generator, register it in the global generator list (`packages/fx-core/src/component/generator/generatorProvider.ts`):