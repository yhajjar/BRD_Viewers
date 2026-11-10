#!/usr/bin/env python3
"""Generate an XML Office add-in manifest using the externally accessible HTTPS URL."""
import argparse
from pathlib import Path
from xml.sax.saxutils import escape
from urllib.parse import urlparse

p=argparse.ArgumentParser()
p.add_argument('--url',required=True,help='Public HTTPS origin, e.g. https://brd.example.org')
p.add_argument('--out',default='deploy/manifest.xml')
a=p.parse_args()
origin=a.url.rstrip('/')
u=urlparse(origin)
if u.scheme!='https' or not u.hostname or u.path or u.query or u.fragment:
    p.error('--url must be an HTTPS origin with no path or query')
h=escape(origin, {'"':'&quot;'})
xml=f'''<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="TaskPaneApp">
  <Id>6c797a12-7948-4424-b3b0-a909a609e238</Id>
  <Version>1.0.0.0</Version><ProviderName>Interactive BRD Viewer</ProviderName><DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Interactive BRD Viewer"/><Description DefaultValue="View versioned interactive BRD diagrams inside Microsoft Word."/>
  <IconUrl DefaultValue="{h}/assets/icon32.png"/><HighResolutionIconUrl DefaultValue="{h}/assets/icon80.png"/>
  <SupportUrl DefaultValue="{h}/"/>
  <AppDomains><AppDomain>{h}</AppDomain></AppDomains>
  <Hosts><Host Name="Document"/></Hosts>
  <DefaultSettings><SourceLocation DefaultValue="{h}/index.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts><Host xsi:type="Document"><DesktopFormFactor>
      <GetStarted><Title resid="welcomeTitle"/><Description resid="welcomeDescription"/><LearnMoreUrl resid="viewerUrl"/></GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface"><OfficeTab id="TabHome"><Group id="BRD.Group">
        <Label resid="groupLabel"/><Icon><bt:Image size="16" resid="icon16"/><bt:Image size="32" resid="icon32"/><bt:Image size="80" resid="icon80"/></Icon>
        <Control xsi:type="Button" id="BRD.Open"><Label resid="openLabel"/>
          <Supertip><Title resid="openLabel"/><Description resid="openDescription"/></Supertip>
          <Icon><bt:Image size="16" resid="icon16"/><bt:Image size="32" resid="icon32"/><bt:Image size="80" resid="icon80"/></Icon>
          <Action xsi:type="ShowTaskpane"><TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId><SourceLocation resid="viewerUrl"/></Action>
        </Control>
      </Group></OfficeTab></ExtensionPoint>
    </DesktopFormFactor></Host></Hosts>
    <Resources>
      <bt:Images><bt:Image id="icon16" DefaultValue="{h}/assets/icon16.png"/><bt:Image id="icon32" DefaultValue="{h}/assets/icon32.png"/><bt:Image id="icon80" DefaultValue="{h}/assets/icon80.png"/></bt:Images>
      <bt:Urls><bt:Url id="viewerUrl" DefaultValue="{h}/index.html"/></bt:Urls>
      <bt:ShortStrings><bt:String id="groupLabel" DefaultValue="Interactive BRD"/><bt:String id="openLabel" DefaultValue="Open BRD Viewer"/><bt:String id="welcomeTitle" DefaultValue="Interactive BRD Viewer"/></bt:ShortStrings>
      <bt:LongStrings><bt:String id="openDescription" DefaultValue="Open interactive sequence diagrams for this document."/><bt:String id="welcomeDescription" DefaultValue="Use the Home ribbon to open the interactive BRD task pane."/></bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'''
out=Path(a.out);out.parent.mkdir(parents=True,exist_ok=True);out.write_text(xml,encoding='utf8')
print(f'Wrote {out} for {origin}')
