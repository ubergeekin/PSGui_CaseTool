<#
###################################################################################################

.Synopsis: CaseTool - Case folder and tool management dashboard

.Name: CaseTool

.DESCRIPTION: GUI tool for SR case folder creation and creating common subfolders. 
              This tool creates a parent foder and subfolders commonly used by SST analysts
              and opens commonly used tools

.CREATED BY: Paul Marquardt

.Date of Creation: 7/2/2018

.Version: 1.0 - Initial form
.Version: 1.1 - Updated form and tools
.version: 1.2 - Removed checkboxes
.Change Log:
 
 07/03/2018: 
 Removed checkboxes to open tools
 Added shaded rectangles
 Added Percolate
 Added FLckr
 Added author info
 .version: 1.1.1 - Updated form to include instructions for use
 
 07/05/2018:
 Added instructions
 Added Quality
 Added Quality including tag component lookup
 Added Delta
 Added OraKB
 Added TurboTech
 Added FileExchanger
 Added SSFD
 
 07/18/2018:
 Updated icon to png:  Author credit: Ions made by https://www.flaticon.com/authors/freepik title="Screwdriver and wrench"
 Add tool tips to form controls: https://www.petri.com/add-popup-tips-powershell-winforms-script
 
 08/14/2018
 Updated DriFT link to version 1.20
 
 08/31/2018
 Updated DriFT link to version 1.22
 
 9/18/2018
 Updated Tesseract location
 Added MSDT, Diags, PAL Reports, SOS Reports check boxes
 Updated DriFT link to programmaticaly pull the latest version
 Added Delta button
 Added CaseTool button
 Added OCP Docs button
 Added HW on-call button
 Added SW on-call buttion
 
 9/19/2018
 Added SATC button
 
 12/6/2018
 Added Get-Filename function
 Add icon to temp directory

 12/7/2018
 Added IE open functionality for Delta link

 12/11/2018 v1.3
 Removed local icon since compiltion changes the path to the development workstation path
 Removed tooltips since they are not a click function
 Removed Get-Filename function
 
 02/14/2019 v1.4
 Changed DriFT button code to grab RunDrift.exe

 02/18/2019 v1.5
 Added OEM button to test for MS non-OEM
 
 02/21/2019
 Change DriFT button code to grab the Drift.zip file from new location
 Updated the SW On-Call page to the on-call calendar

 02/26/2019
 Updated OEM button function to overcome powershell update causing $null value (Kudos to Tommy Paulk for his assistance!)

 03/22/2019
 Fixed bug with MSDT folder name
 
 04/18/2019
 Added Docs listing for documentation links
 Changed check boxes and tools to be alphabetical
 Removed FileExchanger
 Removed old Delta button function, only opens with IE now
 
 4/19/2019
 Removed text from browse text box
 
 
###################################################################################################
#>

# Variable cleanup
Remove-Variable * -ErrorAction SilentlyContinue

# Assign URL variables and tool variables
$IconURI = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/Shared%20Documents/CaseTool/wrench.jpg"
$DriFTURI = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/DRiFT/SitePages/Community%20Home.aspx"
$DriFTURL="https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/DRiFT/DRiFT%20Docs/Drift.zip"
$TesseractURL = "https://internal.software/tesseract/releases/6.0.0/tesseract.7z"
$FLCkrURL = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Software-Documents/FLCkr_v1.0.exe.remove"
$PercURL = "http://percolate.internal.software/releases/latest/percolate.zip"
$QiURL = "http://quality.dell.com/services"
$QiTagURL = "http://quality.dell.com/search/?tag="
$SPMDURL = "https://spmd.dell.com/SPMD/Search/default"
$DeltaURL = "https://isp.us.dell.com/callcenter_enu/start.swe?"
$OraKBURL = "https://kb.dell.com/infocenter/index?page=home"
$TTURL = "https://solutions.one.dell.com/sites/CSO/Programs/TT/turbo/default.htm"
$FEURL = "http://fileexchangerinside.dell.com/tech/DefaultTechView.aspx"
$SSFDURL = "\\AUSVNX01MP02.amer.dell.com\Tumbleweedpl_32483ssfdp2MP\ssfdp2_prod"
$MSDTURL = "https://home.diagnostics.support.microsoft.com/SelfHelp/?wa=wsignin1.0"
$SDDCURL = "https://github.com/PowerShell/$module/archive/master.zip"
$ErrURL = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/Shared%20Documents/CaseTool/Err.zip"
$MoxieURL = "https://channelslb.us.dell.com/netagent/mainlogin.aspx/"
$OCPURL = "https://kb.dell.com/infocenter/index?page=content&id=HOW11455"
$HWoncallURL = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Pages/Server/Welcome-Server.aspx"
$SWoncallURL = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Lists/Software-OnCall/calendar.aspx"
$CaseToolURL = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/_layouts/15/start.aspx#/Shared Documents/Forms/AllItems.aspx?RootFolder=%2fsites%2fNAEnterprise%2fSST%2fCommunities%2fShared Documents%2fCaseTool&FolderCTID=0x012000DD0478FFBE6FB047B62B0391B5FCF907"
$SATCURL = "https://satc.dell.com"
$CaseTool_Help = "https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/Shared%20Documents/CaseTool/Docs/CaseTool%20Documentation.html"

# Set 7-zip alias
Set-Alias 7z "$env:ProgramFiles\7-zip\7z.exe"

# Assign icon path to temp directory
$Icon = "$env:TEMP\CaseTool\wrench.jpg"

# Copy icon to temp directory
#	if (Test-Path "$env:TEMP\CaseTool") {Remove-Item "$env:TEMP\CaseTool" -force -Recurse}
#    new-item -itemtype Directory -force -path "$env:TEMP" -name "CaseTool"
#    Invoke-WebRequest -Uri $IconURI -UseDefaultCredentials -OutFile "$env:TEMP\CaseTool\wrench.jpg"

# Read XAML form
# NOTE: When adding new buttons or XAML form items, remove: x:Name="CaseTool" from <Windows and TextChanged from SrPath,  Click="OpenDriFT_Click"  Click="Browse_Click_1" from buttons
$inputXML = @"
<Window x:Class="CaseTool.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:CaseTool"
        mc:Ignorable="d"
        Title="CaseTool 1.7 Folder/Tool Mgt" Height="765.788" Width="281.417" Icon="https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/Shared%20Documents/CaseTool/wrench.jpg">
    <Grid Margin="0,1,0.2,1.4" RenderTransformOrigin="0.505,0.55">

        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="396" Margin="111,160,0,0" Stroke="Black" VerticalAlignment="Top" Width="151"/>

        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="117" Margin="111,40,0,0" Stroke="Black" VerticalAlignment="Top" Width="151"/>

        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="263" Margin="9,40,0,0" Stroke="Black" VerticalAlignment="Top" Width="101"/>
        <TextBox x:Name="SrPath" HorizontalAlignment="Left" Height="23" Margin="9,11,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="167" FontFamily="Segoe UI Light"/>
        <Button x:Name="Make" Content="Create Folders" HorizontalAlignment="Left" Margin="18,248,0,0" VerticalAlignment="Top" Width="81" Height="23"/>
        <CheckBox x:Name="ChkTSR" Content="TSR" HorizontalAlignment="Left" Margin="15,230,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkEvt" Content="Event Logs" HorizontalAlignment="Left" Margin="14,130,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkDrift" Content="DriFT" HorizontalAlignment="Left" Margin="14,115,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkClus" Content="Cluster" HorizontalAlignment="Left" Margin="14,72,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkPerf" Content="PerfLogs" HorizontalAlignment="Left" Margin="14,202,0,0" VerticalAlignment="Top"/>
        <Button x:Name="Browse" Content="Browse" HorizontalAlignment="Left" Margin="181,11,0,0" VerticalAlignment="Top" Width="81" Height="23" AutomationProperties.HelpText="Click to browse or create a new folder"/>
        <CheckBox x:Name="ChkNIC" Content="Network" HorizontalAlignment="Left" Margin="14,173,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkDsk" Content="Disk" HorizontalAlignment="Left" Margin="14,101,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkDmp" Content="Checkdump" HorizontalAlignment="Left" Margin="14,58,0,0" VerticalAlignment="Top"/>
        <Button x:Name="Open" Content="Open Folder" HorizontalAlignment="Left" Margin="18,276,0,0" VerticalAlignment="Top" Width="81" Height="23" RenderTransformOrigin="0.23,1.347"/>
        <Image x:Name="Wrench" HorizontalAlignment="Left" Height="58" VerticalAlignment="Top" Width="60" RenderTransformOrigin="0.479,0.49" Source="https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/Shared%20Documents/CaseTool/wrench.jpg" Margin="199,607,0,0"/>
        <Label Content="By Paul Marquardt" HorizontalAlignment="Left" VerticalAlignment="Top" Opacity="0.5" FontSize="6" Margin="199,665,0,0"/>
        <Button x:Name="OpenQiTag" Content="Tag Components" HorizontalAlignment="Left" Margin="116,97,0,0" VerticalAlignment="Top" Width="70" Height="23" FontSize="8"/>
        <TextBox x:Name="SvcTagBox" HorizontalAlignment="Left" Height="23" Margin="160,69,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="85" FontStyle="Italic" Opacity="0.5" FontWeight="Bold"/>
        <Button x:Name="OpenDelta" Content="Delta" HorizontalAlignment="Left" Margin="187,187,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenOKB" Content="OraKB" HorizontalAlignment="Left" Margin="188,309,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenTT" Content="TurboTech" HorizontalAlignment="Left" Margin="187,370,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenSSFD" Content="SSFD" HorizontalAlignment="Left" Margin="188,339,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenMSDT" Content="MSDT" HorizontalAlignment="Left" Margin="188,280,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Label x:Name="Links" Content="Links" HorizontalAlignment="Left" Margin="197,158,0,0" VerticalAlignment="Top" Width="39" Height="25" Opacity="0.35" FontWeight="Bold"/>
        <Label x:Name="ToolsLabel_Copy1" Content="Service Tag Lookup" HorizontalAlignment="Left" Margin="125,44,0,0" VerticalAlignment="Top" Width="140" Height="25" Opacity="0.35" FontWeight="Bold"/>
        <CheckBox x:Name="ChkCBS" Content="CBS Logs" HorizontalAlignment="Left" Margin="14,44,0,0" VerticalAlignment="Top"/>
        <Button x:Name="OpenEntitle" Content="Entitlements" HorizontalAlignment="Left" Margin="188,97,0,0" VerticalAlignment="Top" Width="70" Height="23" FontSize="8"/>
        <Label x:Name="SvcTagLabel" Content="SvcTag" HorizontalAlignment="Left" Margin="124,69,0,0" VerticalAlignment="Top" FontSize="8"/>
        <Button x:Name="OpenSDDC" Content="Get-SDDC" HorizontalAlignment="Left" Margin="187,432,0,0" VerticalAlignment="Top" Width="58" Height="23" FontSize="9"/>
        <Button x:Name="OpenTess" Content="Tesseract" HorizontalAlignment="Left" Margin="121,461,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenDriFT" Content="DRiFT" HorizontalAlignment="Left" Margin="121,218,0,0" VerticalAlignment="Top" Width="58" Height="23" />
        <Label x:Name="ToolsLabel" Content="Tools" HorizontalAlignment="Left" Margin="130,158,0,0" VerticalAlignment="Top" Width="39" Height="25" Opacity="0.35" FontWeight="Bold"/>
        <Button x:Name="OpenFLCkr" Content="FLCkr" HorizontalAlignment="Left" Margin="121,280,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenPerc" Content="PERColate" HorizontalAlignment="Left" Margin="121,339,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenQi" Content="Quality" HorizontalAlignment="Left" Margin="121,369,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenSPMD" Content="SPMD" HorizontalAlignment="Left" Margin="121,430,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <CheckBox x:Name="ChkMSinfo" Content="MSinfo32" HorizontalAlignment="Left" Margin="14,158,0,0" VerticalAlignment="Top"/>
        <Button x:Name="OpenErr" Content="Err" HorizontalAlignment="Left" Margin="121,249,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Button x:Name="OpenOCP" Content="OCP Docs" HorizontalAlignment="Left" Margin="187,462,0,0" VerticalAlignment="Top" Width="58" Height="23" FontSize="9"/>
        <Button x:Name="OpenMoxie" Content="Moxie" HorizontalAlignment="Left" Margin="187,249,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <CheckBox x:Name="ChkMSDT" Content="MSDT" HorizontalAlignment="Left" Margin="14,144,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkDiags" Content="Diagnotics" HorizontalAlignment="Left" Margin="14,86,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="ChkPAL" Content="PAL Reports" HorizontalAlignment="Left" Margin="14,187,0,0" VerticalAlignment="Top"/>
        <Button x:Name="OpenHWoncall" Content="HW On-Call" HorizontalAlignment="Left" Margin="187,493,0,0" VerticalAlignment="Top" Width="58" Height="23" FontSize="9"/>
        <Button x:Name="OpenSWoncall" Content="SW On-Call" HorizontalAlignment="Left" Margin="187,523,0,0" VerticalAlignment="Top" Width="58" Height="23" FontSize="9"/>
        <Button x:Name="OpenCaseTool" Content="CaseTool" HorizontalAlignment="Left" Margin="121,187,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <CheckBox x:Name="ChkSOS" Content="SOS Reports" HorizontalAlignment="Left" Margin="14,216,0,0" VerticalAlignment="Top"/>
        <Button x:Name="OpenSATC" Content="SATC" HorizontalAlignment="Left" Margin="121,400,0,0" VerticalAlignment="Top" Width="58" Height="23" RenderTransformOrigin="0.506,-0.649"/>
        <Button x:Name="CaseTool_Help" Content="Help" HorizontalAlignment="Left" Margin="26,623,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <TextBlock x:Name="HelpTxt" HorizontalAlignment="Left" Margin="18,433,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="122"/>
        <Label Content="Click Help for instructions&#xD;&#xA;on how to use CaseTool&#xD;&#xA;" HorizontalAlignment="Left" Margin="18,652,0,0" VerticalAlignment="Top" Height="51" Width="161"/>
        <Button x:Name="OEM" Content="OEM" HorizontalAlignment="Left" Margin="152,123,0,0" VerticalAlignment="Top" Width="70" Height="23" RenderTransformOrigin="0.506,-0.649"/>
        <Button x:Name="OpenLighning" Content="Lightning" HorizontalAlignment="Left" Margin="187,218,0,0" VerticalAlignment="Top" Width="58" Height="23"/>
        <Label x:Name="Docs" Content="Docs" HorizontalAlignment="Left" Margin="197,399,0,0" VerticalAlignment="Top" Width="39" Height="25" Opacity="0.35" FontWeight="Bold"/>
        <Button x:Name="Open1Note" Content="OneNote" HorizontalAlignment="Left" Margin="121,309,0,0" VerticalAlignment="Top" Width="58" Height="23"/>

    </Grid>
</Window>




"@       
 
# Cleanup XAML form metadata for PowerShell
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
# try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}

# Comment: Check for "changed" properties in form which PowerShell cannot process
# catch{Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged properties (PowerShell cannot process them)"
#    throw}
 

# Load XAML Objects In PowerShell
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Form=[Windows.Markup.XamlReader]::Load( $reader )
$xaml.SelectNodes("//*[@Name]") | %{try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
     catch{throw}
    }

Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
# write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
# Make form objects work
add-type -AssemblyName System.Windows.Forms

# Assign the tooltips variable
#$toolstips1 = New-Object System.Windows.Forms.ToolTip

# Browse button function
$WPFBrowse.add_Click({
    $sr = New-Object System.Windows.Forms.FolderBrowserDialog
    if($sr.ShowDialog() -eq 'OK'){
    $WPFSrPath.Text = $sr.SelectedPath
    }
})

# Create form contol functions
$WPFMake.Add_Click({
        if ($WPFChkClus.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "Cluster"}
		if ($WPFChkDrift.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "Drift"}
        if ($WPFChkDsk.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "Disk"}
        if ($WPFChkEvt.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "Evt_Logs"}
        if ($WPFChkNIC.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "NIC"}
        if ($WPFChkPerf.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "Performance"}
        if ($WPFChkTSR.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "TSR"}
        if ($WPFChkDmp.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "Dumps"}
        if ($WPFChkCBS.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "CBS_Logs"}
        if ($WPFChkMSinfo.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "MSinfo32"}
		if ($WPFChkMSDT.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "MSDT"}
		if ($WPFChkDiags.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "Diagnostics"}
		if ($WPFChkPAL.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "PAL_Reports"}
		if ($WPFChkSOS.IsChecked -eq $true){new-item -itemtype Directory -force -path $WPFSrPath.Text -name "SOS_Reports"}
})

# Case Folder Open button function
$WPFOpen.Add_Click({
     ii $WPFSrPath.Text
})

# DriFT button function
$WPFOpenDriFT.Add_Click({
    if (Test-Path "$env:TEMP\DriFT") {Remove-Item "$env:TEMP\DriFT" -force -Recurse}
    new-item -itemtype Directory -force -path "$env:TEMP" -name "DriFT"
	Invoke-Webrequest -uri $DriFTURL -UseDefaultCredentials -Outfile "$env:TEMP\DriFT\Drift.zip"
	7z x -o"$env:TEMP\DriFT" "$env:TEMP\DriFT\Drift.zip" -r;
	ii  "$env:TEMP\DriFT\Drift.exe"

})

# Tesseract button function
$WPFOpenTess.Add_Click({
    if (Test-Path "$env:TEMP\Tesseract") {Remove-Item "$env:TEMP\Tesseract" -force -Recurse}
    new-item -itemtype Directory -force -path "$env:TEMP" -name "Tesseract"
    Invoke-WebRequest -Uri $TesseractURL -UseDefaultCredentials -OutFile "$env:TEMP\Tesseract\Tesseract.7z"
    7z x -o"$env:TEMP\Tesseract" "$env:TEMP\Tesseract\Tesseract.7z" -r;
    ii "$env:TEMP\Tesseract\Tesseract.exe"
})

# FLCkr button function
$WPFOpenFLCkr.Add_Click({
    if (Test-Path "$env:TEMP\FLCkr") {Remove-Item "$env:TEMP\FLCkr" -Force -Recurse}
    new-item -itemtype Directory -force -path "$env:TEMP" -name "FLCkr"
    Invoke-WebRequest -Uri $FLCkrURL -UseDefaultCredentials -OutFile "$env:TEMP\FLCkr\FLCkr_v1.0.exe"
    ii "$env:TEMP\FLCkr"
})

# Percolate button function
$WPFOpenPerc.Add_Click({
    if (Test-Path "$env:TEMP\Percolate") {Remove-Item "$env:TEMP\Percolate" -Force -Recurse}
    new-item -itemtype Directory -force -path "$env:TEMP" -name "Percolate"
    Invoke-WebRequest -Uri $PercURL -UseDefaultCredentials -OutFile "$env:TEMP\Percolate\Percolate.zip"
    7z x -o"$env:TEMP\Percolate" "$env:TEMP\Percolate\Percolate.zip" -r;
    ii "$env:TEMP\Percolate\Percolate.exe"
})

# Quality button function OpenQi
$WPFOpenQi.Add_Click({
  # Invoke-WebRequest -Uri $QiURL -UseDefaultCredentials
    Start-Process -FilePath $QiURL
})

# QiTagged button fuction
$WPFOpenQiTag.Add_Click({
    $SvcTag = $WPFSvcTagBox.Text
    Start-Process -FilePath $QiTagURL$SvcTag"#Components"
})

# OpenEntitle button function
$WPFOpenEntitle.Add_Click({
    $SvcTag = $WPFSvcTagBox.Text
    Start-Process -FilePath $QiTagURL$SvcTag"#AERO"
})

# SPMD button fucntion
$WPFOpenSPMD.Add_Click({
    Start-Process -FilePath $SPMDURL
})

# OKB button function
$WPFOpenOKB.Add_Click({
    Start-Process -FilePath $OraKBURL
})

# Open TurboTech button function
$WPFOpenTT.Add_Click({
    Start-Process -FilePath $TTURL
})

# Open SSFD button function
$WPFOpenSSFD.Add_Click({
    Start-Process -FilePath $SSFDURL
})

# Open OCP docs button function
$WPFOpenOCP.Add_Click({
    Start-Process -FilePath $OCPURL
})

# Open Moxie button function
$WPFOpenMoxie.Add_Click({
    Start-Process -FilePath $MoxieURL
})

# Open Delta button function
$WPFOpenDelta.Add_Click({
    $ie = New-Object -com internetexplorer.application
    $ie.navigate2("$DeltaURL")
    $ie.visible=$true
})

# Open CaseTool button function
$WPFOpenCaseTool.Add_Click({
    Start-Process -FilePath $CaseToolURL
})

# Open HW on-call button function
$WPFOpenHWoncall.Add_Click({
    Start-Process -FilePath $HWoncallURL
})

# Open OCP docs button function
$WPFOpenSWoncall.Add_Click({
    Start-Process -FilePath $SWoncallURL
})

# Open SATC button function
$WPFOpenSATC.Add_Click({
    $SvcTag = $WPFSvcTagBox.Text
    Start-Process -FilePath $SATCURL"/#/"$SvcTag"/devices/"$SvcTag #Opens SATC with SVCTag from SvcTag lookup
})

# Open SDDC button function
$WPFOpenSDDC.Add_Click({
    if (Test-Path "$env:TEMP\Get-SDDC") {Remove-Item "$env:TEMP\Get-SDDC" -Force -Recurse}
    new-item -itemtype Directory -force -path "$env:TEMP" -name "Get-SDDC"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $module = 'PrivateCloud.DiagnosticInfo'
    Invoke-WebRequest -Uri https://github.com/PowerShell/$module/archive/master.zip -OutFile $env:TEMP\Get-SDDC\Get-SDDC.zip
    ii "$env:TEMP\Get-SDDC"
})

# Open MSDT button function
$WPFOpenMSDT.Add_Click({
    Start-Process -FilePath $MSDTURL
})

# Open Help docs button function
$WPFCaseTool_Help.Add_Click({
    Start-Process -FilePath $CaseTool_Help
})

# OEM button function
$WPFOEM.Add_Click({
    $SvcTag = $WPFSvcTagBox.Text
    $OpenQiTag = "http://quality.dell.com/search/?tag=$SvcTag#Components"
	$qitag = wget -Uri $OpenQiTag -UseDefaultCredentials
	$noos = $qitag.RAwContent | select-string -pattern "KD483" -quiet
    if ($noos -eq $True){
    Start-Process -FilePath "https://quality.dell.com/quicksearch?q=KD483"
    }
   else {
    Start-Process -FilePath $OpenQiTag
}   
})

<# $qitag = wget -Uri $QiTagUL$SvcTag"#Components" -UseDefaultCredentials
    $qitag
    $noos = $qitag.RAwContent | select-string -pattern "KD483" -quiet
    if ($noos -eq $true){
    Start-Process -FilePath "https://quality.dell.com/quicksearch?q=KD483"
    else
    Start-Process -FilePath $QiTagURL$SvcTag"#Components"
}
#>

# Err button function
$WPFOpenErr.Add_Click({
<#    if (Test-Path "$env:SystemRoot\System32\Err.exe"){
       #Start PowerShell -noexit -Command "&""$env:SystemRoot\System32\Err.exe" -ArgumentList $WPFErrBox.Text
		Start-Process -FilePath "$env:SystemRoot\System32\Err.exe" -ArgumentList $WPFErrBox.Text 
        $input = Read-Host -Prompt "Are you finished? (Y/N)"
        if($input -eq "Y"){
        exit
        }
})
    else{
#>
        if (Test-Path "$env:TEMP\Err") {Remove-Item "$env:TEMP\Err" -Force -Recurse}
        new-item -itemtype Directory -force -path "$env:TEMP" -name "Err"    
        Invoke-WebRequest -Uri $ErrURL -UseDefaultCredentials -OutFile $env:TEMP\Err\Err.zip
        7z x -o"$env:TEMP" "$env:TEMP\Err\Err.zip" -r;
        ii "$env:TEMP\Err"
        # Copy-Item -path "$env:TEMP\Err\Err.exe" -Destination "$env:SystemRoot\System32" -Force
        # ii "$env:SystemRoot\System32\Err.exe"$WPFErrBox.Text
        # Invoke-Expression -command "Err "$WPFErrBox.Text -NoExit
#    }
})


#===========================================================================
# Show the form
#===========================================================================
# write-host "To show the form, run the following" -ForegroundColor Cyan
$Form.ShowDialog() | out-null