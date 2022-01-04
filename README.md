# M365CallFlowVisualizer

>[!Warning]
>Test

# Synopsis
Reads a config from Microsoft 365 Phone System and renders them visually into a mermaid-js flowchart.

# Examples

## Example 1: Auto attendant forwards calls to a Teams user

![](/Examples/PS_Test_AA_CallFlow_aa_fwd_user.svg)

## Example 2: Auto attendant checks for business hours and forwards calls to a call queue

![](/Examples/PS_Test_AA_CallFlow_only_BusinessHours.svg)

## Example 3: Auto attendant checks for holidays and business hours and forwards calls to a call queue

![](/Examples/PS_Test_AA_CallFlow_Holidays_and_BusinessHours.svg)

## Example 4: Top-level call queue without auto attendant

![](/Examples/CQ_Team_Green_No_AA.svg)

These are just examples which were dynamically rendered based on my Microsoft 365 Phone System configuration. If an auto attendant does not have business hours or holidays, the flowchart will be much smaller.

# How to use it

## Prerequisites

I suggest using Visual Studio Code and the official PowerShell Extension. This script needs the "MSOnline" and "MicrosoftTeams" PowerShell modules. It has been tested with MicrosoftTeams PowerShell version 2.3.1.

The script supports outputting Mermaid-JS code in either a Markdown file (.md) or a Mermaid file (.mmd).

To preview Markdown files containing mermaid sections I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=tomoyukim.vscode-mermaid-editor)

To preview Mermaid files I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid). You can also export to SVG directly from this extension.

### Before you run the script

Make sure that you are connected to MsolService and MicrosoftTeams by running the following commands:

```PowerShell
Connect-MsolService
```

```PowerShell
Connect-MicrosoftTeams
```

## Tips & Links

You can also copy the contents of the output file and paste it [here](https://mermaid-js.github.io/mermaid-live-editor), if you want to edit the generated Mermaid flowchart. You can also export to SVG or PNG directly from the live editor.

You can find more information about Mermaid syntax [here](https://mermaid-js.github.io/mermaid/#/)

If you want to implement Mermaid Diagrams into your markdown based documentation site, I suggest to take a look at [DocFx](https://dotnet.github.io/docfx/).

# Known limitations
- No support for IVRs yet
- No support for multiple resource accounts/numbers mapped to one auto attendant or call queue yet
- No support for cascaded call queues or auto attendants yet. If your call flow redirects to another queue at some point, it will only show the type and name of the voice app but will not visualize the settings of the target app.

# Planned feature updates
- Reflect if voicemail transcription or suppress system greeting is on
- Display call queue and auto attendant language settings
- Custom HEX color support for the mermaid diagram

These are planned changes. There is no ETA nor is it guaranteed that these features will ever be added.

# Legal
This script is provided free of charge. Please do not sell it in any form.
