' Title: Custom Ribbon Tabs
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/custom-ribbon-tabs/td-p/13831088
' Category: advanced
' Scraped: 2025-10-09T09:02:57.759518

<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <EnableComHosting>true</EnableComHosting>
    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>
  </PropertyGroup>

  <ItemGroup>
    <COMReference Include="Inventor">
      <WrapperTool>tlbimp</WrapperTool>
      <VersionMinor>0</VersionMinor>
      <VersionMajor>1</VersionMajor>
      <Guid>d98a091d-3a0f-4c3e-b36e-61f62068d488</Guid>
      <Lcid>0</Lcid>
      <Isolated>false</Isolated>
      <EmbedInteropTypes>False</EmbedInteropTypes>
    </COMReference>
  </ItemGroup>

  <ItemGroup>
    <Reference Include="AdWindows">
      <HintPath>C:\Program Files\Autodesk\Inventor 2026\Bin\AdWindows.dll</HintPath>
      <Private>False</Private>
    </Reference>
  </ItemGroup>

  <ItemGroup>
    <None Update="NewTab.addin">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
  </ItemGroup>

  <Target Name="PostBuild" AfterTargets="PostBuildEvent">
    <Exec Command="xcopy /y /d &quot;$(TargetDir)*.*&quot; &quot;%APPDATA%\Autodesk\Inventor 2026\Addins\&quot;" />
  </Target>

</Project>