# M365CallFlowVisualizer

>[!Warning]
>Test

# Synopsis
Reads a config from Microsoft 365 Phone System and renders them visually into a mermaid-js flowchart.

# How to use it

## Prerequisites

I suggest using Visual Studio Code and the official PowerShell Extension. This script needs the "MSOnline" and "MicrosoftTeams" PowerShell modules. It has been tested with MicrosoftTeams PowerShell version 3.0.0 and 3.0.1-preview.

### Install Modules

Run these two commands in an elevated PowerShell window.

```PowerShell
Install-Module MSOnline
```

```PowerShell
Install-Module MicrosoftTeams
```
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
<p>As this uses `| Out-Gridview` it's only supported on Windows platforms.</p>
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
.\M365CallFlowVisualizerV2.ps1 -DocType Markdown -SetClipBoard $false
```

This will run the script, present a list of the available voice apps and save the call flow to a markdown (*.md) file without copying the markdown syntax to the clipboard.

### Example 6

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -SafeToFile $false
```

This will run the script without saving the call flow to a file. Make sure to not set SetClipBoard to $false as this would result in no output at all.

### Example 7

```PowerShell
.\M365CallFlowVisualizerV2.ps1 -CustomFilePath "C:\Temp"
```

This will run the script and save the output file to "C:\Temp".

## Preview Mermaid Code

The script supports outputting Mermaid-JS code in either a Markdown file (.md) or a Mermaid file (.mmd).

To preview Markdown files containing mermaid sections I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=tomoyukim.vscode-mermaid-editor)

To preview Mermaid files I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid). You can also export to SVG directly from this extension.

# Example Outputs

## Example 1: Auto attendant forwards calls to a Teams user

![](/Examples/PS_Test_AA_CallFlow_aa_fwd_user.svg)

## Example 2: Auto attendant checks for business hours and forwards calls to a call queue

![](/Examples/PS_Test_AA_CallFlow_only_BusinessHours.svg)

## Example 3: Auto attendant checks for holidays and business hours and forwards calls to a call queue

![](/Examples/PS_Test_AA_CallFlow_Holidays_and_BusinessHours.svg)

## Example 4: Top-level call queue without auto attendant

![](/Examples/CQ_Team_Green_No_AA.svg)

## Example 5: Auto attendant check for holidays and business hours, forwards to nested auto attendant after hours

![](/Examples/USA_Toll_Free_Test_Example.svg)

These are some examples based on real configurations inside a Microsoft 365 Tenant. With the new V2 Version it's theoretically possible to render nested voice apps indefinitely. Loops should also be reflected correctly, altough the diagram can look a little weird. The logic now detects if multiple voice apps forward to the same target and will render each voice app only one time.

## Tips & Links

You can also copy the contents of the output file and paste it [here](https://mermaid-js.github.io/mermaid-live-editor), if you want to manually edit the generated Mermaid flowchart. The live editor supports exporting flow charts as *.png or *.svg images.

You can find more information about Mermaid syntax [here](https://mermaid-js.github.io/mermaid/#/)

If you want to implement Mermaid Diagrams into your markdown based documentation site, I suggest to take a look at [DocFx](https://dotnet.github.io/docfx/).

# Known limitations
- No support for IVRs yet
- The tool has only been tested on Windows systems!

# Planned feature updates
- Reflect if voicemail transcription or suppress system greeting is on
- Display call queue and auto attendant language settings
- Custom HEX color support for the mermaid diagram
- Migrate from MSOnline to Microsoft Graph PowerShell

These are planned changes. There is no ETA nor is it guaranteed that these features will ever be added.

# Legal
This script is provided free of charge. Please do not sell it in any form. Please include my name, Twitter and GitHub handle/links if you plan to post about this tool online or offline. Thank you.
