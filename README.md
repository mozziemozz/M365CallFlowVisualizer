# M365CallFlowVisualizer

# Synopsis
Reads a config from Microsoft 365 Phone System and renders them visually into a mermaid-js flowchart.

# Examples

## No Holidays and no Business Hours

[![](https://mermaid.ink/img/eyJjb2RlIjoiZmxvd2NoYXJ0IFRCXG5zdGFydCgoSW5jb21pbmcgQ2FsbCBhdCA8YnI-ICs0MTQ0eHh4eHh4eCkpIC0tPiBlbGVtZW50QUEoW0F1dG8gQXR0ZW5kYW50IDxicj4gUFMgVGVzdCBBQV0pIC0tPiBkZWZhdWx0Q2FsbEZsb3dHcmVldGluZz5HcmVldGluZyA8YnI-IE5vbmVdIC0tPiBkZWZhdWx0Q2FsbEZsb3coVHJhbnNmZXJDYWxsVG9UYXJnZXQpIC0tPiBkZWZhdWx0Q2FsbEZsb3dBY3Rpb24oVXNlciA8YnI-IE1pa2UgV2FnbmVyKSIsIm1lcm1haWQiOnsidGhlbWUiOiJkYXJrIn0sInVwZGF0ZUVkaXRvciI6ZmFsc2UsImF1dG9TeW5jIjp0cnVlLCJ1cGRhdGVEaWFncmFtIjpmYWxzZX0)](https://mermaid-js.github.io/mermaid-live-editor/edit/#eyJjb2RlIjoiZmxvd2NoYXJ0IFRCXG5zdGFydCgoSW5jb21pbmcgQ2FsbCBhdCA8YnI-ICs0MTQ0eHh4eHh4eCkpIC0tPiBlbGVtZW50QUEoW0F1dG8gQXR0ZW5kYW50IDxicj4gUFMgVGVzdCBBQV0pIC0tPiBkZWZhdWx0Q2FsbEZsb3dHcmVldGluZz5HcmVldGluZyA8YnI-IE5vbmVdIC0tPiBkZWZhdWx0Q2FsbEZsb3coVHJhbnNmZXJDYWxsVG9UYXJnZXQpIC0tPiBkZWZhdWx0Q2FsbEZsb3dBY3Rpb24oVXNlciA8YnI-IE1pa2UgV2FnbmVyKSIsIm1lcm1haWQiOiJ7XG4gIFwidGhlbWVcIjogXCJkYXJrXCJcbn0iLCJ1cGRhdGVFZGl0b3IiOmZhbHNlLCJhdXRvU3luYyI6dHJ1ZSwidXBkYXRlRGlhZ3JhbSI6ZmFsc2V9)

# How to use it

## Prerequisites

I suggest using Visual Studio Code and the official PowerShell Extension. This script needs the "MSOnline" and "MicrosoftTeams" PowerShell modules. It has been tested with MicrosoftTeams PowerShell version 2.3.1.

The script supports outputting Mermaid-JS code in either a Markdown file (.md) or a Mermaid file (.mmd).

To preview Markdown files containing mermaid sections I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=tomoyukim.vscode-mermaid-editor)

To preview Mermaid files I suggest the following [VS Code Extension](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid). You can also export to SVG directly from this extension.

## Tips & Links

You can also copy the contents of the output file and paste it [here](https://mermaid-js.github.io/mermaid-live-editor), if you want to edit the generated Mermaid flowchart. You can also export to SVG or PNG directly from the live editor.

You can find more information about Mermaid syntax [here](https://mermaid-js.github.io/mermaid/#/)

If you want to implement Mermaid Diagrams into your markdown based documentation site, I suggest to take a look at [DocFx](https://dotnet.github.io/docfx/)

# Known limitations
- No support for IVRs yet
- No support for multiple resource accounts/numbers mapped to one auto attendant or call queue yet
- No support for cascaded call queues or auto attendants yet. If your call flow redirects to another queue at some point, it will only show the type and name of the voice app but will not visualize the settings of the target app.

# Planned feature updates
- Reflect if voicemail transcription or suppress system greeting is on
- Display call queue and auto attendant language settings
- Custom HEX color support for the mermaid diagram

These are planned changes. There is no ETA nor is it guaranteed that these features will ever be added.