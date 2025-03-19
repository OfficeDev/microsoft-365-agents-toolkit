# Run "VSC USE" v0.1.7 On Your Machine

## Environment Setup (Windows Only)

1. **Install Python**

    - Download from [Python official website](https://www.python.org/downloads/)
    - During installation, ensure "Add Python to PATH" is checked
    - Recommended: Use the latest version (**python 3.12** is full tested, other version not tested yet)

2. **Install Node.js and npm**

    - Download from [Node.js official website](https://nodejs.org/)
    - Verify installation by running:
    ```bash
    node -v
    npm -v
    ```

3. **Install Appium**

    - Install globally using npm:
    ```bash
    npm install -g appium
    appium driver install windows
    ```

    - Verify installation:
    ```bash
    appium -v
    ```

4. **Install WinAppDriver**
    - Download from [WinAppDriver releases](https://github.com/microsoft/WinAppDriver/releases)
    - [**Enable Developer Model**](https://learn.microsoft.com/en-us/windows/apps/get-started/enable-your-device-for-development) in Windows Setting
    - After installation, start WinAppDriver (**run as Administrator**):
    ```bash
    cd C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe
    ```
   - Default listening address: `http://127.0.0.1:4723`

5. **Install Visual Studio Code**
    - Download from [VS Code website](https://code.visualstudio.com/)
    - Install with default settings

## VSC Use Installation

1. **Set API Key**
    - "VSC Use" support three model: 'google/gemini-2.0-flash', 'google/gemini-2.0-flash-lite', 'azure/gpt-4o'.
        - **Gemini Flash 2.0** and **Gemini Flash 2.0** are highly recommend. To get Google API key by just simply click a button on [Google API Doc](https://ai.google.dev/gemini-api/docs/api-key). It is free to try.
        - If Gemini Flash 2.0 is enable, "VSC Use" will enable image processing capability.

    - Setup Model API Key accordingly in your windows environment, choose one you preferred:

    > ### For Google Gemini 2.0 Flash Lite
    > MODEL_NAME="google/gemini-2.0-flash-lite"
    >
    > GOOGLE_API_KEY="your_google_api_key_here"

    > ### For Google Gemini 2.0 Flash
    > MODEL_NAME="google/gemini-2.0-flash"
    >
    > GOOGLE_API_KEY="your_google_api_key_here"

    **NOTICE: if choose "azure/gpt-4o", there is rate limit rule for internal employee: "50k token per minites". Due to this limitation, "VSC Use" may reach 50k in one minites, so it will execute the task slowly.**

    > ### For Azure GPT-4O
    > MODEL_NAME="azure/gpt-4o"
    > 
    > AZURE_OPENAI_KEY="your_azure_openai_key_here"
    >
    > AZURE_OPENAI_ENDPOINT="your_azure_openai_endpoint_here"

    - Set Azure Vision Endpoint in your windows environment

    > AZURE_VISION_ENDPOINT="your_azure_vision_endpoint_here"
    >
    > AZURE_VISION_KEY="your_azure_vision_key_here"

2. Start the webdriver server:
   ```bash
   appium --use-drivers=windows
   ```

3. Download "VSC Use" from here, unzip the file to any directory your want, cd to the directory and run it.
    ```bash
    vscuse.exe -v
    ```

## Run your first test plan

> IMPORTANT: **Test Plan = Prompt**
> In "VSC Use": Test plan is describe in natural language. "VSC Use" as an agent, play like a UI tester to interact with computer by Mouse and Keyboard. So when writing test plan, image how a human-being interact with a computer.

- **run vscuse in one command:**
```bash
vscuse agent -p "open vs code and open command palette. then type teams: get start. press enter."
```
If you see the "Get Started Page" of Teams Toolkit. Congragulation!!!

- **use vscuse to run test plan file:**
    - create a test plan file under same fodler where "Vsc Use" located. Let's say we have a file named "test_copilot_chat"
    ```text
    open vs code.
    open github copilot chat panel in vs code.
    type "what model are you?" in copilot chat box.
    ```

    - once test plan file is ready. Run below command:
    ```bash
    vsc agent -f test_copilot_chat -s 20
    ```

- Now, it is your turn to write your own first test plan.

- Try "vscuse agent --help" to see more parameter.

![alt text](images/vscuse_help.png)
