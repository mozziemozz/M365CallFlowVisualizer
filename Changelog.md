# Changelog

|Date|Component|Version|Change|
|---|---|---|---|
|20.10.2021|M365CallFlowVisualizerV2.ps1||Creation|
|21.10.2021|M365CallFlowVisualizerV2.ps1||Add comments and streamline code, add longer arrow links for default call flow desicion node|
|21.10.2021|M365CallFlowVisualizerV2.ps1||Add support for top level call queues (besides auto attendants)|
|21.10.2021|M365CallFlowVisualizerV2.ps1||Move call queue specific operations into a function|
|24.10.2021|M365CallFlowVisualizerV2.ps1||Fixed a bug where Disconnect Call was not reflected in mermaid correctly when CQ timeout action was disconnect call|
|30.10.2021|M365CallFlowVisualizerV2.ps1|2.0.0| most of the script logic was moved into functions. Added parameters for specifig resource account (specified by phone number), added support for nested queues, added support to display only 1 queue if timeout and overflow go to the same queue.|
|01.11.2021|M365CallFlowVisualizerV2.ps1||Add support to display call queues for an after hours call flow of an auto attendant|
|01.11.2021|M365CallFlowVisualizerV2.ps1||Fix issue where additional entry point numbers were not shown on after hours call flow call queues|
|02.11.2021|M365CallFlowVisualizerV2.ps1||Add support for nested Auto Attendants|
|03.01.2022|M365CallFlowVisualizerV2.ps1|2.1.0|V2.1 more or less a complete rewrite of the script logic to make it really dynamic and support indefinite chaning/nesting of voice apps|
|        |M365CallFlowVisualizerV2.ps1|| Add support to disable rendering of nested voice apps|
|        |M365CallFlowVisualizerV2.ps1|| Add support for voice app name and type parameters|
|        |M365CallFlowVisualizerV2.ps1|| Fixed a bug where some phone numbers which contained extensions including a ";" were not rendered in mermaid. (replace ";" with ",")|
|        |M365CallFlowVisualizerV2.ps1||Fixed a bug where nested voice apps of an auto attendant were rendered even though business hours were set to default.|
|        |M365CallFlowVisualizerV2.ps1||Added support for custom file paths, option to disable saving the file|
|04.01.2022|M365CallFlowVisualizerV2.ps1|||Prettify format of business hours (remove seconds from string)|
|05.01.2022|M365CallFlowVisualizerV2.ps1|||Add H1 Title to Markdown document, add support for mermaid themes default, dark, neutral and forest, change default DocType to Markdown|
|05.01.2022|M365CallFlowVisualizerV2.ps1||Add new parameters and support for displaying call queue agents opt in status and phone number|
|05.01.2022|M365CallFlowVisualizerV2.ps1||Fix clipboard content when markdown is selected, add support to display phone numbers assigned to voice apps in grid view selection|
|05.01.2022|M365CallFlowVisualizerV2.ps1||Change Markdown title from H1 to H2. Fix bug in phone number listing on voice app selection|
|07.01.2022|M365CallFlowVisualizerV2.ps1||Add support for IVRs in an auto attendants default call flow. After Hours call flows and forward to announcements are not supported yet.|
|07.01.2022|M365CallFlowVisualizerV2.ps1||Merge changes from default call flow to after hours call flow (IVR support), some optimizations regarding after hours call flow id (more robust way)|
|07.01.2022|M365CallFlowVisualizerV2.ps1||Add support for announcements and operators in IVRs|
|08.01.2022|M365CallFlowVisualizerV2.ps1||Start implementing custom HEX colors for nodes, borders, links and fonts|
|09.01.2022|M365CallFlowVisualizerV2.ps1||Add support for custom hex colors|
|10.01.2022|M365CallFlowVisualizerV2.ps1||fix bug where custom hex colors were not applied if auto attendant doesn't have business hours|
|10.01.2022|M365CallFlowVisualizerV2.ps1||sometimes Teams PS fails to read leading + from OnlineApplicationInstance, added code to add + if not present|
|12.01.2022|M365CallFlowVisualizerV2.ps1||Migrate from MSOnline Cmdlets to Microsoft.Graph. Add support to export call flows to *.htm files for easier access and sharing.|
|12.01.2022|M365CallFlowVisualizerV2.ps1||Optimize some error messages / warnings.|
|13.01.2022|M365CallFlowVisualizerV2.ps1||Add support to display if MS System Message is being played back in holidays, CQ time out, overflow or AA default or after horus call flow|
|13.01.2022|M365CallFlowVisualizerV2.ps1||fixed a bug where operator AAs or CQs were added to nested voice apps even if not configured in call flows|
|13.01.2022|M365CallFlowVisualizerV2.ps1||Add numbers on links to reflect agent list order in call queues with serial routing method.|
|14.01.2022|M365CallFlowVisualizerV2.ps1||Optimize presentation of call queue agents (list vertically) because queues with many agents made the diagram too wide|
|14.01.2022|M365CallFlowVisualizerV2.ps1||Agent list type now lists group(s) name(s) or team name and channel name, prettify name of AA Holiday subgraphs|
|20.01.2022|M365CallFlowVisualizerV2.ps1||Change default value of ExportHtml to true|
|20.01.2022|M365CallFlowVisualizerV2.ps1||Add support to include TTS greeting texts and audio file names used in auto attendant calls flows, IVR announcements, call queues and music on hold|
|21.02.2022|M365CallFlowVisualizerV2.ps1||Add support to export audio files from auto attendants and call queues and link them in html output (on node click)|
|24.01.2022|M365CallFlowVisualizerV2.ps1||Add support to export TTS greeting values as txt files and link them on nodes|
|25.01.2022|M365CallFlowVisualizerV2.ps1||Fixed a bug where users or external PSTN numbers were added to nested voice apps, if configured as operator which caused the script to stop|
|26.01.2022|M365CallFlowVisualizerV2.ps1||Change SetClipboard default value to false, add parameter to start browser/open tab with exported html|
|26.01.2022|M365CallFlowVisualizerV2.ps1||Fixed a bug where it was not possible to run the script for a voice app which doesn't have a number|
|02.02.2022|M365CallFlowVisualizerV2.ps1|2.5.0| Add support to display if auto attendant is voice response enabled and show voice responses on IVR options, add support for custom hex color in subgraphs, optimize call queue structure, don't draw call queue greetings if none are set|
|03.02.2022|M365CallFlowVisualizerV2.ps1|2.5.1| Don't draw greeting nodes if no greeting is configured in auto attendant default or after hours call flows|
|03.02.2022|M365CallFlowVisualizerV2.ps1|2.5.2| Microsoft has changed how time ranges in schedules are displayed which caused the script to always show business hours desicion nodes, even when none were set. this has been addressed with a fix in this version.|
|03.02.2022|M365CallFlowVisualizerV2.ps1|2.5.3| Holiday greeting nodes are now also only drawn if a greeting is configured|
|03.02.2022|M365CallFlowVisualizerV2.ps1|2.5.4| Optimize login function to make sure that the tenants for Teams and Graph are always the same.|
|04.02.2022|M365CallFlowVisualizerV2.ps1|2.5.5| Fix bug with html export and mermaid theme, add theme support for mermaid export|
|09.02.2022|M365CallFlowVisualizerV2.ps1|2.5.6| Fix bug in Connect-CFV where the Teams and Graph TenantId check was not always working.|
|05.03.2022|M365CallFlowVisualizerV2.ps1|2.5.7| Add Leading + Agents phone numbers|
|14.03.2022|M365CallFlowVisualizerV2.ps1|2.5.8| Fix Connect-M365CFV function (Sometimes the check if Teams and Graph tenant are the same failed when there was a cached graph session)|
|15.03.2022|M365CallFlowVisualizerV2.ps1|2.5.9| Improve order of node shapes for call queue timeout and overflow to voicemail, don't show CQ greeting if overflow threshold is set to 0|
|19.03.2022|M365CallFlowVisualizerV2.ps1|2.6.0| Fix bug / optimzie error handling for finding after hours schedule (now looking for type instead of call flow name containing "after")|
|20.03.2022|M365CallFlowVisualizerV2.ps1|2.6.0b| Apply fix from 2.6.0 also to business hours|
|21.03.2022|M365CallFlowVisualizerV2.ps1|2.6.1| Fix detection of no business hours, don't draw call distribution for CQs which have overflow threshold 0 anymore|
|07.04.2022|M365CallFlowVisualizerV2.ps1|2.6.2| Optimize Connect-M365CFV login checks|
|08.04.2022|M365CallFlowVisualizerV2.ps1|2.6.3| Fix breaking changes from MicrosoftTeams PowerShell 4.1.0. This version is now required. Move Connect-M365CFV out of Script into seperate file. Fix display of CQ Agents without phone number. Move Changelog out of script into repository. Fix output errors when exporting audio files or TTS greetings but none were present in the voice apps. Fix display of CQ Agents without phone number. Move Changelog out of script into repository. Fix output errors when exporting audio files or TTS greetings but none were present in the voice apps. |
|08.04.2022|M365CallFlowVisualizerV2.ps1|2.6.4|Remove '(' and ')' from audio file names because this caused a syntax error in mermaid. |
|12.04.2022|M365CallFlowVisualizerV2.ps1|2.6.5|Sometimes CQ properties are returned in .Value and sometimes not. This version implements changes to handle these kind of differences. Optimize diagram when CQ Overflow threshold is 0. |
|16.04.2022|Get-TeamsUserCallFlow.ps1|1.0.0|Finalize first version of function for standalone use. Create example script to run the function for each enabled user of a tenant.|
|17.04.2022|Get-TeamsUserCallFlow.ps1|1.0.1|Create ouptut directory if it doesn't exist.|
|17.04.2022|M365CallFlowVisualizerV2.ps1|2.6.6|Create ouptut directory if it doesn't exist. Set default value of CustomFilePath to .\Output|
|18.04.2022|M365CallFlowVisualizerV2.ps1|2.6.6| add FixDisplayName function to fix DisplayNames with (), it affects the mermaid render|
