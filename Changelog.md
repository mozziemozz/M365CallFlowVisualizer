# Changelog

|Date|Component|Version|Change|
|---|---|---|---|
|20.10.2021|M365CallFlowVisualizerV2.ps1.ps1||Creation|
|21.10.2021|M365CallFlowVisualizerV2.ps1.ps1||Add comments and streamline code, add longer arrow links for default call flow desicion node|
|21.10.2021|M365CallFlowVisualizerV2.ps1.ps1||Add support for top level call queues (besides auto attendants)|
|21.10.2021|M365CallFlowVisualizerV2.ps1.ps1||Move call queue specific operations into a function|
|24.10.2021|M365CallFlowVisualizerV2.ps1.ps1||Fixed a bug where Disconnect Call was not reflected in mermaid correctly when CQ timeout action was disconnect call|
|30.10.2021|M365CallFlowVisualizerV2.ps1.ps1|2.0.0| most of the script logic was moved into functions. Added parameters for specifig resource account (specified by phone number), added support for nested queues, added support to display only 1 queue if timeout and overflow go to the same queue.|
|01.11.2021|M365CallFlowVisualizerV2.ps1.ps1||Add support to display call queues for an after hours call flow of an auto attendant|
|01.11.2021|M365CallFlowVisualizerV2.ps1.ps1||Fix issue where additional entry point numbers were not shown on after hours call flow call queues|
|02.11.2021|M365CallFlowVisualizerV2.ps1.ps1||Add support for nested Auto Attendants|
|03.01.2022|M365CallFlowVisualizerV2.ps1.ps1|2.1.0|V2.1 more or less a complete rewrite of the script logic to make it really dynamic and support indefinite chaning/nesting of voice apps|
|        |M365CallFlowVisualizerV2.ps1.ps1|| Add support to disable rendering of nested voice apps|
|        |M365CallFlowVisualizerV2.ps1.ps1|| Add support for voice app name and type parameters|
|        |M365CallFlowVisualizerV2.ps1.ps1|| Fixed a bug where some phone numbers which contained extensions including a ";" were not rendered in mermaid. (replace ";" with ",")|
|        |M365CallFlowVisualizerV2.ps1.ps1||Fixed a bug where nested voice apps of an auto attendant were rendered even though business hours were set to default.|
|        |M365CallFlowVisualizerV2.ps1.ps1||Added support for custom file paths, option to disable saving the file|
|04.01.2022|M365CallFlowVisualizerV2.ps1.ps1|||Prettify format of business hours (remove seconds from string)|
|05.01.2022|M365CallFlowVisualizerV2.ps1.ps1|||Add H1 Title to Markdown document, add support for mermaid themes default, dark, neutral and forest, change default DocType to Markdown|
|05.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add new parameters and support for displaying call queue agents opt in status and phone number|
|05.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Fix clipboard content when markdown is selected, add support to display phone numbers assigned to voice apps in grid view selection|
|05.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Change Markdown title from H1 to H2. Fix bug in phone number listing on voice app selection|
|07.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add support for IVRs in an auto attendants default call flow. After Hours call flows and forward to announcements are not supported yet.|
|07.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Merge changes from default call flow to after hours call flow (IVR support), some optimizations regarding after hours call flow id (more robust way)|
|07.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add support for announcements and operators in IVRs|
|08.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Start implementing custom HEX colors for nodes, borders, links and fonts|
|09.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add support for custom hex colors|
|10.01.2022|M365CallFlowVisualizerV2.ps1.ps1||fix bug where custom hex colors were not applied if auto attendant doesn't have business hours|
|10.01.2022|M365CallFlowVisualizerV2.ps1.ps1||sometimes Teams PS fails to read leading + from OnlineApplicationInstance, added code to add + if not present|
|12.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Migrate from MSOnline Cmdlets to Microsoft.Graph. Add support to export call flows to *.htm files for easier access and sharing.|
|12.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Optimize some error messages / warnings.|
|13.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add support to display if MS System Message is being played back in holidays, CQ time out, overflow or AA default or after horus call flow|
|13.01.2022|M365CallFlowVisualizerV2.ps1.ps1||fixed a bug where operator AAs or CQs were added to nested voice apps even if not configured in call flows|
|13.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add numbers on links to reflect agent list order in call queues with serial routing method.|
|14.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Optimize presentation of call queue agents (list vertically) because queues with many agents made the diagram too wide|
|14.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Agent list type now lists group(s) name(s) or team name and channel name, prettify name of AA Holiday subgraphs|
|20.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Change default value of ExportHtml to true|
|20.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add support to include TTS greeting texts and audio file names used in auto attendant calls flows, IVR announcements, call queues and music on hold|
|21.02.2022|M365CallFlowVisualizerV2.ps1.ps1||Add support to export audio files from auto attendants and call queues and link them in html output (on node click)|
|24.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Add support to export TTS greeting values as txt files and link them on nodes|
|25.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Fixed a bug where users or external PSTN numbers were added to nested voice apps, if configured as operator which caused the script to stop|
|26.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Change SetClipboard default value to false, add parameter to start browser/open tab with exported html|
|26.01.2022|M365CallFlowVisualizerV2.ps1.ps1||Fixed a bug where it was not possible to run the script for a voice app which doesn't have a number|
|02.02.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.0| Add support to display if auto attendant is voice response enabled and show voice responses on IVR options, add support for custom hex color in subgraphs, optimize call queue structure, don't draw call queue greetings if none are set|
|03.02.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.1| Don't draw greeting nodes if no greeting is configured in auto attendant default or after hours call flows|
|03.02.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.2| Microsoft has changed how time ranges in schedules are displayed which caused the script to always show business hours desicion nodes, even when none were set. this has been addressed with a fix in this version.|
|03.02.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.3| Holiday greeting nodes are now also only drawn if a greeting is configured|
|03.02.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.4| Optimize login function to make sure that the tenants for Teams and Graph are always the same.|
|04.02.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.5| Fix bug with html export and mermaid theme, add theme support for mermaid export|
|09.02.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.6| Fix bug in Connect-CFV where the Teams and Graph TenantId check was not always working.|
|05.03.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.7| Add Leading + Agents phone numbers|
|14.03.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.8| Fix Connect-M365CFV function (Sometimes the check if Teams and Graph tenant are the same failed when there was a cached graph session)|
|15.03.2022|M365CallFlowVisualizerV2.ps1.ps1|2.5.9| Improve order of node shapes for call queue timeout and overflow to voicemail, don't show CQ greeting if overflow threshold is set to 0|
|19.03.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.0| Fix bug / optimzie error handling for finding after hours schedule (now looking for type instead of call flow name containing "after")|
|20.03.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.0b| Apply fix from 2.6.0 also to business hours|
|21.03.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.1| Fix detection of no business hours, don't draw call distribution for CQs which have overflow threshold 0 anymore|
|07.04.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.2| Optimize Connect-M365CFV login checks|
|08.04.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.3| Fix breaking changes from MicrosoftTeams PowerShell 4.1.0. This version is now required. Move Connect-M365CFV out of Script into seperate file. Fix display of CQ Agents without phone number. Move Changelog out of script into repository. Fix output errors when exporting audio files or TTS greetings but none were present in the voice apps. Fix display of CQ Agents without phone number. Move Changelog out of script into repository. Fix output errors when exporting audio files or TTS greetings but none were present in the voice apps. |
|08.04.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.4|Remove '(' and ')' from audio file names because this caused a syntax error in mermaid. |
|12.04.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.5|Sometimes CQ properties are returned in .Value and sometimes not. This version implements changes to handle these kind of differences. Optimize diagram when CQ Overflow threshold is 0. |
|16.04.2022|Get-TeamsUserCallFlow.ps1|1.0.0|Finalize first version of function for standalone use. Create example script to run the function for each enabled user of a tenant.|
|17.04.2022|Get-TeamsUserCallFlow.ps1|1.0.1|Create ouptut directory if it doesn't exist.|
|17.04.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.6|Create ouptut directory if it doesn't exist. Set default value of CustomFilePath to .\Output|
|27.04.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.7|Microsoft changed some stuff again how values are returned. Many attributes were behind a ".Value" which has been removed by MS. Also Business Hours don't include "DisplayHours" anymore, a Function was added to properly read Business Hours|
|27.04.2022|Read-BusinessHours|1.0.0|Function created to read business hours properly|
|18.04.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.6|add FixDisplayName function to fix DisplayNames with (), it affects the mermaid render|
|17.05.2022|M365CallFlowVisualizerV2.ps1.ps1|2.6.8|Fix audio file or TTS greeting paths in HTML when using a custom file path|
|03.06.2022|Get-TeamsUserCallFlow.ps1|1.0.2|Add Mermaid nodes and subgraphs to variables of main script for theme support|
|03.06.2022|M365CallFlowVisualizerV2.ps1|2.6.9|Integrate Get-TeamsUserCallFlow.ps1 into main Script. This means that User Calling Settings can now be drawn as well.|
|11.06.2022|Get-MsSystemMessage.ps1|1.0.0|Add switch statement to determine what the MS System Greeting says in which language.|
|03.06.2022|M365CallFlowVisualizerV2.ps1|2.7.0|Implement FixDisplayName for Holiday Call Handling Names as well. Add support to display and export the MS System Greeting in German and English.|
|15.06.2022|M365CallFlowVisualizerV2.ps1|2.7.1|Fixed multiple links if a call flow forwarded to the same user on different actions. Local MS Graph Cache is now cleared if there are issues signing in.|
|11.06.2022|Get-MsSystemMessage.ps1|1.0.1|Added all supported languages to the switch statement as preparation for the translation by the community.|
|21.06.2022|M365CallFlowVisualizerV2.ps1|2.7.2|The tool is now also displaying user calling settings of an Operator (if set to a user)|
|17.08.2022|M365CallFlowVisualizerV2.ps1|2.7.3|Correctly displaying Skip voicemail system message on call queues with TTS and audio file greetings. Add support for redirect to a user's personal voicemail on CQ overflow and timeout.|
|21.08.2022|M365CallFlowVisualizerV2.ps1|2.7.4|Fix an error where nested voice apps were rendered when CQ overflow was 0. Added Support for 0 Timeouts on CQs (Call Distribution is not rendered anymore).|
|21.08.2022|Get-TeamsUserCallFlow.ps1|2.7.4|Add support for unlimited user and voice app nesting in user calling settings.|
|21.08.2022|Get-MsSystemMessage.ps1|1.0.2|Added safe character greetings for languages with special characters for in diagram nodes.|
|31.08.2022|M365CallFlowVisualizerV2.ps1|2.7.5|Improve HTML Output compatibility.|
|31.08.2022|M365CallFlowVisualizerV2.ps1|2.7.6|Rename FixDisplayName to Optimize-Displayname. Remove more special characters from display names. Improve formatting of holiday lists. Add Parameter for US and EU date formats.|
|02.09.2022|M365CallFlowVisualizerV2.ps1|2.7.7|Add support for obfuscating phone numbers. Suppress Conf Mode warnings on CQs. Correctly reflect order of greeting, system greeting and transfer for shared voicemail on AAs and CQs.|
|02.09.2022|Get-TeamsUserCallFlow.ps1|1.0.3|Add support for obfuscating phone numbers. Unify names for external transfers (change from External PSTN to Number).|
|04.09.2022|M365CallFlowVisualizerV2.ps1|2.7.8|Add support to write user information of users which are linked to CQs or AAs into an external variable.|
|04.09.2022|New-VoiceAppUserLinkProperties.ps1|1.0.0|Initial commit.|
|04.09.2022|Find-CallQueueAndAutoAttendantUserLinks.ps1|1.0.0|Initial commit.|
|04.09.2022|Find-CallQueueAndAutoAttendantUserLinks.ps1|1.0.1|Bug fixes.|
|04.09.2022|M365CallFlowVisualizerV2.ps1|2.7.9|Bug fixes (User Calling Settings were drawn on CQs when a user was configured as Overflow despite timeout being 0)|
|04.09.2022|M365CallFlowVisualizerV2.ps1|2.8.0|Add support to expand members of M365 Groups which are configured for Shared Voicemail on AAs and CQs|
|04.09.2022|Get-SharedVoicemailGroupMembers.ps1|1.0.0|Initial commit.|
|06.09.2022|M365CallFlowVisualizerV2.ps1|2.8.1|Fix typo (turncate --> truncate)|
|07.10.2022|M365CallFlowVisualizerV2.ps1|2.8.2|Add support for Tenants with up to 1000 AAs, CQs and Resource Accounts, general performance improvements.|
|07.10.2022|Get-TeamsUserCallFlow.ps1|1.0.4|Add support for Tenants with up to 1000 AAs, CQs and Resource Accounts, general performance improvements.|
|07.10.2022|M365CallFlowVisualizerV2.ps1|2.8.2b|Rename Call Connected Node to Agent Answered.|
|01.11.2022|M365CallFlowVisualizerV2.ps1|2.8.3|Improve robustness of audio file names, add support for PNG export through mermaid-cli, add support for force listen option on auto attendants, fix bug in displaying voice responses|
|04.11.2022|M365CallFlowVisualizerV2.ps1|2.8.4|Handle exception when download uri for MoH/Welcome Music were not available on call queue object|
|04.11.2022|M365CallFlowVisualizerV2.ps1|2.8.5|Add support to also obfuscate phone numbers of CQ agents. Fix filename related bugs for PNG export. Updated Readme and examples. Rename Agent Answered back to Call Connected (same as user calling settings. Fix bug where shared voicemail members were not expanded on AA holidays is skip vm system greeting was enabled.|
|04.11.2022|Get-SharedVoicemailGroupMembers.ps1|1.0.1|When ObfuscatePhoneNumbers is true, email addresses of shared mailbox members will also be anonymized if ShowSharedVoicemailGroupMembers is true.|
|06.11.2022|Get-IvrTransferMessage.ps1|1.0.0|Initial creation. This function will provide the text which is synthesized by an auto attendant when transferring from an IVR to operator, user or external PSTN.|
|06.11.2022|M365CallFlowVisualizerV2.ps1|2.8.6|Add support for accurate display of IVR transfer messages. Optimize Greeting/Transfer order of display for IVRs.|
|28.11.2022|M365CallFlowVisualizerV2.ps1|2.8.7|Fix bug / add Optimize-Displayname for auto attendant default and after hours call flow nodes.|
|28.11.2022|M365CallFlowVisualizerV2.ps1|2.8.7b|Add Optimize-Displayname to more elements to prevent mermaid syntax errors if special characters are used in TAC config elements.|
|03.01.2023|M365CallFlowVisualizerV2.ps1|2.8.8|Add Optimize-Displayname to all TTS greetings.|
|03.01.2023|Optimize-DisplayName.ps1|1.0.2|Replace "call" with "Call" in mermaid node text because this breaks mermaid.|
|07.01.2023|M365CallFlowVisualizerV2.ps1|2.8.9|Add parameter to also include outbound calling Ids of call queues to the diagram.|
|19.01.2023|M365CallFlowVisualizerV2.ps1|2.9.0|Merge Pull request: use global vars to fasten up parent runner scripts. Thanks to MicheleBomello :) |
|28.01.2023|M365CallFlowVisualizerV2.ps1|2.9.1|Add parameter to enable or disable the global variables. Add maxTextSize to mermaid init to support larger files. Add output for shared voicemail GroupId|
|30.01.2023|M365CallFlowVisualizerV2.ps1|2.9.2|Add parameter to expand/show nested AAs and CQs of holiday call handlings, fix greeting which was shown on AA call flows when none was configured with DisconnectCall action. Show voice stlye (Male/Female) on voice response enabled AA.|
|01.02.2023|M365CallFlowVisualizerV2.ps1|2.9.3|Add parameter to display IVRs of Holiday Call Handlings. Fix links when multiple holidays pointed to a Voice App.|
|01.02.2023|M365CallFlowVisualizerV2.ps1|2.9.4|Rename -NoCache to -CacheResults as it's easier to not to think about a double negative value.|
|03.02.2023|M365CallFlowVisualizerV2.ps1|2.9.5|Add support for Call Queue Welcome Text-To-Speech Greetings. Requires MicrosoftTeams PowerShell 4.9.3.|
|12.02.2023|M365CallFlowVisualizerV2.ps1|2.9.6|Performance improvements (reduce number of Get-Cs*), Add parameter to show number type for CQ Agents, outbound calling Ids, Loop until a top level voice app is selected.|
|13.02.2023|M365CallFlowVisualizerV2.ps1|2.9.7|Fix bug where transfer message was displayed in AA default/after hours call flow without IVR. Add user calling settings to nested holiday call flows. Extend -MaxResults from 1000 to 9999.|
|13.02.2023|Get-AutoAttendantDirectorySearchConfig.ps1|1.0.0|Initial development to read search scope configurations from AAs.|
|13.02.2023|M365CallFlowVisualizerV2.ps1|2.9.8|Include directory search scope for AAs with voice menus.|
|14.02.2023|M365CallFlowVisualizerV2.ps1|2.9.9|Add params to combine all "Call Connected" and "DisconnectCall" nodes.|
|14.02.2023|Get-TeamsUserCallFlow.ps1|1.0.5|Add param to combine all "Call Connected" and "DisconnectCall" nodes.|
|15.02.2023|M365CallFlowVisualizerV2.ps1|2.9.9b|Add holiday name on link text for holiday IVRs and nested call flows.|
|24.02.2023|HtmlTemplate.html|1.0.1|Add support for Mermaid Version 10.0.0.|
|28.02.2023|Optimize-DisplayName.ps1|1.0.3|Replace ’ with ' and re-save file as UTF-8 with BOM.|
|17.03.2023|M365CallFlowVisualizerV2.ps1|3.0.0|Make retrieving all AAs and CQs more robust. Minor changes to outputs/inputs.|
|04.04.2023|M365CallFlowVisualizerV2.ps1|3.0.1|Move retrieving all AAs, CQs and RAs into separate function. Change Markdown title from H1 to H2.|
|04.04.2023|Get-AllVoiceAppsAndResourceAccounts.ps1|1.0.0|Move retrieving all AAs, CQs and RAs into separate function.|
|04.04.2023|AllTopLevelVoiceAppsToMarkdown.ps1|1.0.3|Move retrieving all AAs, CQs and RAs into separate function. Set Output to `.\Output\AllTopLevelVoiceApps.`|
|05.04.2023|M365CallFlowVisualizerV2.ps1|3.0.2|Add support to display outbound caller Ids of individual CQ agents. Add "HardcoreMode" to easily enable all parameters which show additional information on the diagram.|
|21.04.2023|M365CallFlowVisualizerV2.ps1|3.0.3|Add support to use in combination with DocFx. More info will follow shortly.|
|21.04.2023|AllTopLevelVoiceAppsToMarkdownDocFx.ps1|1.0.4|Add support to use in combination with DocFx. More info will follow shortly.|
|07.05.2023|M365CallFlowVisualizerV2.ps1|3.0.4|Add 2 new parameters to expand and include user call groups and delegates.|
|07.05.2023|Get-TeamsUserCallFlow.ps1|1.0.6|Add support to also expand and include user call groups and delegates.|
|07.05.2023|Get-TeamsUserCallFlow.ps1|1.0.7|Bug fixes.|
|14.05.2023|Get-TeamsUserCallFlow.ps1|1.0.8|Add support for serial call group user nesting. Change ExportSvg and PreviewSVG default values to $false since they're currently broken.|
|02.06.2023|M365CallFlowVisualizerV2.ps1|3.0.5|Remove "Call Flow" prefix from H2 title in HTML output.|
|02.06.2023|AllTopLevelVoiceAppsToMarkdownDocFx.ps1|1.0.5|Add param for relative path.|
|02.06.2023|AllTopLevelVoiceAppsToMarkdownDocFx.ps1|1.0.6|Rename "call_flows.md" to "call-flows.md" to follow the awesome-docfx-template repo.|
|02.06.2023|Find-CallQueueAndAutoAttendantUserLinks.ps1|1.0.2|Migrate to Get-AllVoiceAppsAndResourceAccounts function and set -CacheResults to True for performance improvements.|
|04.07.2023|M365CallFlowVisualizerV2.ps1|3.0.6|Add new parameter `-ShowCqAuthorizedUsers` to display authorized users of call queues.|
|18.08.2023|M365CallFlowVisualizerV2.ps1|3.0.7|Add support for CQ no agents opted/logged.|
|18.08.2023|M365CallFlowVisualizerV2.ps1|3.0.8|Add new parameter `-ShowAaAuthorizedUsers` to display authorized users of auto attendants.|
|26.10.2023|M365CallFlowVisualizerV2.ps1|3.0.9|Fix Team names with special characters (Add Optimize-DisplayName function).|
|27.10.2023|M365CallFlowVisualizerV2.ps1|3.0.10|Add support for Teams and Graph sign in via Entra ID App Registration / Service Principal.|
|27.10.2023|Connect-M365CFV.ps1|1.1.0|Add support for Teams and Graph sign in via Entra ID App Registration / Service Principal.|
|27.10.2023|Get-M365CFVTeamsAdminToken.ps1|1.0.0|Add support for Teams and Graph sign in via Entra ID App Registration / Service Principal.|
|27.10.2023|SecureCredsMgmt.ps1|1.0.0|Add support for Teams and Graph sign in via Entra ID App Registration / Service Principal.|
|27.10.2023|Get-AllVoiceAppsAndResourceAccounts.ps1|1.0.1|Fix result caching by making all relevant variables global.|
|27.10.2023|M365CallFlowVisualizerV2.ps1|3.0.11|Add support for PDF Export. Thanks @MicheleBomello for this PR!|
|06.11.2023|M365CallFlowVisualizerV2.ps1|3.1.0|Add support to check if an Auto Attendant is currently in business- or after hours schedule or in holiday schedule by using -CheckCallFlowRouting.|
|07.11.2023|M365CallFlowVisualizerV2.ps1|3.1.1|Bug fixes, improvements in console and diagram output (support for ComplementEnabled: False schedules)|
|07.11.2023|M365CallFlowVisualizerV2.ps1|3.1.2|Remove (broken) support for ServicePrincipal auth. Add warning message when -ConnectWithServicePrincipal is used.|
|07.11.2023|Get-M365CFVTeamsAdminToken.ps1|1.0.0|File removed.|
|07.11.2023|Connect-M365CFV.ps1|1.1.1|Remove support for Teams and Graph sign in via Entra ID App Registration / Service Principal. This has been moved into its own function.|
|07.11.2023|Connect-MsTeamsServicePrincipal.ps1|1.0.0|Prepare for support for Teams and Graph sign in via Entra ID App Registration / Service Principal once it supports Get-CsOnlineApplicationInstance.|
|22.11.2023|M365CallFlowVisualizerV2.ps1|3.1.3|Fix bug that showed longest idle CQ as presence based routing false when in fact it's true.|
|07.02.2024|M365CallFlowVisualizerV2.ps1|3.1.4|Add parameter `-ShowSharedVoicemailGroupSubscribers` to display if people are following the group in their inbox.|
|07.02.2024|Get-SharedVoicemailGroupMembers.ps1|1.0.2|Add parameter `-ShowSharedVoicemailGroupSubscribers` to display if people are following the group in their inbox.|
|06.04.2024|Connect-M365CFV.ps1|1.1.2|Add `Connect-ExchangeOnline`.|
|15.04.2024|M365CallFlowVisualizerV2.ps1|3.1.5|Add support for app only auth for some features Use `-ConnectWithServicePrincipal` (Exchange Online/ `-ShowSharedVoicemailGroupSubscribers` are not supported yet).|
|15.04.2024|Get-AllVoiceAppsAndResourceAccountsAppAuth.ps1|1.0.1|Add support for app only auth for some features Use `-ConnectWithServicePrincipal` (Exchange Online/ `-ShowSharedVoicemailGroupSubscribers` are not supported yet).|
|15.04.2024|Connect-M365CFV.ps1|1.1.3|Add checks to also prompt for Exchange credentials when using Hardcore mode.|
|15.04.2024|Get-TeamsUserCallFlow.ps1|1.0.9|Optimize performance by checking objects in memory instead of using `Get-CsOnlineApplicationInstance`.|
|29.04.2024|M365CallFlowVisualizerV2.ps1|3.1.6|Add support for greetings on disconnect and redirect to phone number and voice app for overflow, timeout and no agent exceptions.|
|19.12.2024|M365CallFlowVisualizerV2.ps1|3.1.7|Fix outbound calling line ids of CQs not shown when CQ is configured for immediate overflow (threshold 0).|
|20.12.2024|M365CallFlowVisualizerV2.ps1|3.1.8|Add support for nested AAs/CQs without resource accounts (ConfigurationEndpoint instead of ApplicationEndpoint).|
|13.01.2025|M365CallFlowVisualizerV2.ps1|3.1.9|Fix incoming phone numbers shown multiple times in loopback to auto attendants. Fix CQs/AAs configured as ConfigurationEndpoints.|
|22.01.2025|M365CallFlowVisualizerV2.ps1|3.2.0|Add Voice App or Resource Account information to TransferCallToTarget and TransferCallToOperator actions.|
