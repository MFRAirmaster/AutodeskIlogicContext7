' Title: Custom Ribbon Tabs
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-tabs/td-p/13831088
' Category: advanced
' Scraped: 2025-10-07T13:23:10.519492

<?xml version="1.0" encoding="utf-8"?>
<Addin Type="Standard">
	<!-- The Assembly name is the name of your output DLL. -->
	<Assembly>NewTab.dll</Assembly>
	<!-- The ClientId must match the GUID in your StandardAddInServer class. -->
	<ClientId>{33B4722C-392E-4E34-9EA3-2E73DD0F1C13}</ClientId>
	<!-- The ClassId must also match the GUID in your StandardAddInServer class. -->
	<ClassId>{33B4722C-392E-4E34-9EA3-2E73DD0F1C13}</ClassId>
	<DisplayName>New Tab Add-in</DisplayName>
	<Description>A simple add-in that creates a new ribbon tab.</Description>
	<LoadOnStartup>1</LoadOnStartup>
	<UserUnloadable>1</UserUnloadable>
	<Hidden>0</Hidden>
	<SupportedSoftwareVersionGreaterThan>28</SupportedSoftwareVersionGreaterThan>
</Addin>