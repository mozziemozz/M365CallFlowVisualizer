# M365CallFlowVisualizer

# Synopsis
Uses PowerShell to read configurations of Microsoft Teams Auto Attendants, Call Queues and User Settings and renders them visually into a mermaid-js flowchart.

# Changelog

The changelog can be found [here](Changelog.md).

# How to use it

## Prerequisites

Please not that this script has only been tested on Windows systems. If you don't specify an identity or a name, `Out-GridView` is used which, to my knowledge is limited to Windows. Furthermore, I'm not sure how file operation and file paths behave on other, Non-Windows operating systems.

### PowerShell Modules

The following PowerShell Modules are required:

|Module|Required Version|Last Tested Version|
|---|---|---|
|MicrosoftTeams|4.6.0|4.9.0|
|Microsoft.Graph.Users|n/a|1.9.6|
|Microsoft.Graph.Groups|n/a|1.9.6|

### Optional Requirements

If you want to make use of the PNG export feature, you also need the following components.

- Node.JS
- @mermaid-js/mermaid-cli npm package

Please see [here](#install-nodejs-and-mermaid-cli) how to install them.

### Install Modules

Run these two commands in an elevated PowerShell window to install the modules from the PSGallery.

```PowerShell
Install-Module Microsoft.Graph 
```

```PowerShell
Install-Module MicrosoftTeams
```

### Install Node.JS and mermaid-cli

Run this in PowerShell.

```PowerShell
winget install --id=OpenJS.NodeJS  -e
```

```PowerShell
npm install -g @mermaid-js/mermaid-cli
```

Verify that mermaid-cli is installed by running `mmdc --version` in PowerShell.

Please see [this](https://github.com/mermaid-js/mermaid-cli#install-locally) site for more information.

## Parameters

Please see the parameter description directly inline in the script.

## Examples

### Example 1

```PowerShell
.\M365CallFlowVisualizerV2.ps1
```

This will run the script without any parameters / default parameters. You will be presented a list with available auto attendants and call queues.

<div class="notecard note">
<h4>Note</h4>
<p>As this uses "Out-Gridview" it's only supported on Windows platforms.</p>
</div>

### Example 2

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -Identity "6fb84b40-f045-45e8-8c1a-8fc18188exxx"
```

This will run the script for the voice app (auto attendant or call queue, not resource account) with the unique identity of "6fb84b40-f045-45e8-8c1a-8fc18188exxx"

### Example 3

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -VoiceAppName "PS Test AA" -VoiceAppType "Auto Attendant"
```

This will run the script for the auto attendant called "PS Test AA".

### Example 4

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -VoiceAppName "PS Test CQ" -VoiceAppType "Call Queue"
```

This will run the script for the call queue called "PS Test CQ".

### Example 5

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -CustomFilePath "C:\Temp"
```

This will run the script and save the output file to "C:\Temp".

### Example 6

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -VoiceAppName "PS Test AA" -VoiceAppType "Auto Attendant" -DisplayNestedCallFlows $false
```

This will run the script without expanding and rendering call flows of auto attendants or call queues which are nested behind "PS Test AA". Only the names and types of these voice apps will be displayed.

### Example 7

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -VoiceAppName "PS Test AA" -VoiceAppType "Auto Attendant" -Theme dark
```

This will run the script and set the Mermaid theme inside Markdown to dark theme.

## Preview Mermaid Code

The script supports outputting Mermaid-JS code in either a Markdown file (.md) or a Mermaid file (.mmd).

To preview Markdown files containing mermaid sections I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid)

To preview Mermaid files I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=tomoyukim.vscode-mermaid-editor). You can also save an SVG or PNG image with this extension.

# Example Outputs

## Example 1

![](/Examples/png/Support_Number_AA_CallFlow.png)

## Example 2

![](/Examples/png/Main_Number_AA_CallFlow.png)

These are some examples generated from real configurations inside a Microsoft 365 Tenant. Everthing you see has been generated and rendered 100% automatically. With the new V2 Version it's theoretically possible to render nested voice apps and user calling settings indefinitely. Loops should also be reflected correctly, altough the diagram can look a little weird. The logic now detects if multiple voice apps forward to the same target and will render each voice app only one time.

## Tips & Links

You can also copy the contents of the output file and paste it [here](https://mermaid-js.github.io/mermaid-live-editor), if you want to manually edit the generated Mermaid flowchart. The live editor supports exporting flow charts as *.png or *.svg images.

You can find more information about Mermaid syntax [here](https://mermaid-js.github.io/mermaid/#/)

If you want to implement Mermaid Diagrams into your markdown based documentation site, I suggest to take a look at [DocFx](https://dotnet.github.io/docfx/).

# Known limitations
- The tool has only been tested on Windows systems. Some functionalty might not be available on other platforms.
- Forwarding Targets in a holiday list are not expanded.
- IVRs in holiday call handlings are not supported.

# Planned feature updates
- Reflect if voicemail transcription or suppress system greeting is on --> Suppress system message was implemented in V 2.4.2
- Display call queue and auto attendant language settings --> Call Queue language implemented in V 2.4.4
- Custom HEX color support for the mermaid diagram --> Implemented in V 2.3.0
- Migrate from MSOnline to Microsoft Graph PowerShell --> Implemented in V 2.4.0

These are planned changes. There is no ETA nor is it guaranteed that these features will ever be added.

# Legal
This script is provided free of charge. Please do not sell it in any form. Please include my name, Twitter and GitHub handle/links if you plan to post about this tool online or offline. Thank you.
