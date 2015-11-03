[CmdletBinding()]
param (
    [Parameter(Mandatory=0,Position=0)][string]$SharePath='',
    [Parameter(Mandatory=0,Position=1)][string]$IncludeSettings='',
    [Parameter(Mandatory=0,Position=2)][string]$IncludeData='',
    [Parameter(Mandatory=0,Position=3)][string]$ExcludeSettings='',
    [Parameter(Mandatory=0,Position=4)][string]$ExcludeData='',
    [Parameter(Mandatory=0,Position=5)][string]$LogPath='',
    [Parameter(Mandatory=0)][switch]$Force,
    [Parameter(Mandatory=0)][switch]$PassThru
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

#region ## Logging ##
if(!$LogPath){
    if(Test-Path $env:USERPROFILE -PathType Container){
        $LogFile = Get-ChildItem $PSCommandPath | ForEach-Object{Write-Output $(($env:USERPROFILE)+'\'+($_.BaseName)+"_$env:USERNAME"+'.log')}
    } else {throw "Could not find a valid log path."}
} else {
    if(Test-Path $LogPath -PathType Container){
        $LogFile = Get-ChildItem $PSCommandPath | ForEach-Object{Write-Output $("$LogPath"+'\'+($_.BaseName)+"_$env:USERNAME"+'.log')}
    } else {throw "Log path not found."}
}
if(!(Test-Path $LogFile -PathType Leaf)){(Get-Date).ToString() + ' Log file created.' > $LogFile}
#endregion Logging

#region ## Constants ##
$CMPATH = 'CM Profiles'
$CMPROFILEPATH = "$SharePath\$CMPATH\$env:USERNAME"
$CMREGPATH = "$CMPROFILEPATH"
$CMREGFILE = 'ntuser.reg'
$CMDATAPATH = "$CMPROFILEPATH"

$AMPATH = 'AM Profiles'
$AMPROFILEPATH = "$SharePath\$AMPATH\$env:USERNAME"
$AMREG = "$env:USERNAME"+'_settings'
$AMREGZIP = "$AMREG"+'.zip'
$AMREGPATH = "$AMPROFILEPATH\$AMREGZIP"
$AMREGFILE = '*.reg'
$AMDATA = "$env:USERNAME"+'_data'
$AMDATAZIP = "$AMDATA"+'.zip'
$AMDATAPATH = "$AMPROFILEPATH\$AMDATAZIP"
$TMPSUFFIX = (Get-Date).Subtract((Get-Date -Date '1/1/2012')).Ticks
$AMLOCALTMP = "$env:USERPROFILE\"+'AMProfileTmp'+"$TMPSUFFIX"
$AMLOCALREGZIP = "$AMLOCALTMP\$AMREGZIP"
$AMLOCALREGPATH = "$AMLOCALTMP\$AMREG"
$AMLOCALDATAZIP = "$AMLOCALTMP\$AMDATAZIP"
$AMLOCALDATAPATH = "$AMLOCALTMP\$AMDATA"

$REGHEADER = 'Windows Registry Editor Version 5.00'
$REGTMPKEYPREFIX = 'HKEY_LOCAL_MACHINE\UserReg'
$REGKEYPREFIX = 'HKEY_CURRENT_USER'
$REGTMPFILE = 'ntuser_'+"$env:USERNAME"+'.reg'
$REGTMPPATH = "$env:USERPROFILE\$REGTMPFILE"
#endregion Constants

#region ## Functions ##
function Append-Log([string]$msg){
    $msg = (Get-Date).ToString() + " $msg"
    $msg >> $LogFile
    Write-Verbose $msg
    Write-Debug $msg
}
function Die([string]$msg){Append-Log $msg; throw $msg}
function CleanUp-Session(){
    if(Test-Path $AMLOCALTMP -PathType Container){Remove-Item $AMLOCALTMP -Recurse -Force}
}
function Unzip-File(){
    param([string]$Path,[string]$Destination)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Path,$Destination)
}
function Validate-Params(){
    $Params = '' | Select-Object SharePath,IncludeSettings,IncludeData,ExcludeSettings,ExcludeData,LogPath
    $Params.SharePath = $SharePath
    $Params.IncludeSettings = $IncludeSettings
    $Params.IncludeData = $IncludeData
    $Params.ExcludeSettings = $ExcludeSettings
    $Params.ExcludeData = $ExcludeData
    $Params.LogPath = $LogPath
    Append-Log "Validating params: $Params"

    if(!$SharePath -or !(Test-Path $SharePath -PathType Container)){Die 'Aborting. Remote user profile path not found.'}
    $SharePath = $SharePath.TrimEnd('\')
    if(!(Test-Path $SharePath\$AMPATH -PathType Container) -or !(Test-Path $SharePath\$CMPATH -PathType Container)){Die 'Aborting. Invalid profile path. Check constants.'}
    if($IncludeSettings -and !(Test-Path $IncludeSettings -PathType Leaf)){Die 'Aborting. Settings include file not found.'}
    if($IncludeData -and !(Test-Path $IncludeData -PathType Leaf)){Die 'Aborting. Data include file not found.'}
    if($ExcludeSettings -and !(Test-Path $ExcludeSettings -PathType Leaf)){Die 'Aborting. Settings exclude file not found.'}
    if($ExcludeData -and !(Test-Path $ExcludeData -PathType Leaf)){Die 'Aborting. Data exclude file not found.'}
}
function Get-UserProfile(){
    Append-Log 'Searching for user profile.'
    $UP = '' | Select-Object Type,SettingsPath,DataPath,Reg,Data
    if(Test-Path "$CMPROFILEPATH" -PathType Container){
        # Citrix UPM
        Append-Log 'Found Citrix UPM profile.'
        if(!(Test-Path "$CMREGPATH\$CMREGFILE" -PathType Leaf)){Die 'Aborting. Registry file not found.'}
        $UP.Type = 'CM'
        $UP.SettingsPath = Get-ChildItem "$CMREGPATH\$CMREGFILE"
        $UP.DataPath = "$CMDATAPATH"
    } else {
        if(Test-Path "$AMPROFILEPATH" -PathType Container){
            # Autometrix
            Append-Log 'Found Autometrix profile.'
            if(!(Test-Path "$AMREGPATH" -PathType Leaf)){Die 'Aborting. Settings file not found.'}
            if(!(Test-Path "$AMDATAPATH" -PathType Leaf)){Die 'Aborting. Data file not found.'}

            Append-Log 'Copying/extracting Autometrix zip files to local machine.'
            try{# to copy Autometrix profile locally
                New-Item "$AMLOCALTMP" -ItemType directory -Force | Out-Null
                Copy-Item -Path "$AMREGPATH" -Destination "$AMLOCALTMP"
                Copy-Item -Path "$AMDATAPATH" -Destination "$AMLOCALTMP"
                Unzip-File -Path "$AMLOCALREGZIP" -Destination "$AMLOCALREGPATH"
                Unzip-File -Path "$AMLOCALDATAZIP" -Destination "$AMLOCALDATAPATH"
            }catch{
                Append-Log 'Error copying/extracting Autometrix zip files to local machine.'
                Append-Log $_.Exception.ItemName
                Append-Log $_.Exception.Message
                CleanUp-Session
                Die 'Aborting.'
            }
            $UP.Type = 'AM'
            if(Test-Path $AMLOCALREGPATH -PathType Container){
                $UP.SettingsPath = Get-ChildItem "$AMLOCALREGPATH" -Include "$AMREGFILE" -Recurse
            } else {
                $UP.SettingsPath = ''
                Die 'Aborting. Registry files not found.'
            }
            $UP.DataPath = "$AMLOCALDATAPATH"
        } else {
            # FS-Logix
            Append-Log 'Profile not found.'
            Append-Log 'User has already been migrated.'
            $UP.Type = 'FS'
        }
    }
    $UP
}
function Get-Settings(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    try{
        Append-Log 'Loading registry settings.'
        if($UP.SettingsPath){
            $UP.Reg = $UP.SettingsPath | Get-Content
            if($UP.Reg){
                $UP.Reg = $UP.Reg.replace("$REGTMPKEYPREFIX","$REGKEYPREFIX")
            } else {
                $UP.Reg = "$REGHEADER"
                Append-Log 'Warning: Registry files contain no settings.'
            }
        } else {
            $UP.Reg = "$REGHEADER"
            Append-Log 'Warning: No registry files found in user profile.'
        }
        Append-Log 'Registry settings loaded.'
        $UP
    }catch{
        Append-Log 'Error processing registry settings.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
function Get-Data(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    Append-Log 'Locating user data.'
    try{
        $UP.Data = Get-ChildItem "$($UP.DataPath)" -Recurse
        Append-Log 'Data location loaded.'
        $UP
    }catch{
        Append-Log 'Error locating data.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
function Include-Settings(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    try{
        if($IncludeSettings){
            Append-Log 'Processing settings include file.'
            $include = Get-Content $IncludeSettings
            if($include){
                $Script:delete = $false
                $UP.Reg = $UP.Reg | %{
                    foreach ($line in $include){
                        if($line -eq ''){continue}
                        if($_.StartsWith("[$REGKEYPREFIX")){$Script:delete = $true}
                        if($_.StartsWith("[$REGKEYPREFIX$line") -and $_ -ne "[$REGKEYPREFIX]"){$Script:delete = $false; break}
                    }
                    if(!$Script:delete){$_}
                }
            Append-Log 'Settings include file applied.'
            } else {
                Append-Log 'Warning: Settings include file empty. No registry settings will be copied.'
                $UP.Reg = $REGHEADER
            }
        } else {
            Append-Log 'Warning: Settings include file not found. No registry settings will be copied.'
            $UP.Reg = $REGHEADER
        }
        $UP
    }catch{
        Append-Log 'Error processing settings include file.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
function Include-Data(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    try{
        if($IncludeData){
            Append-Log 'Processing data include file.'
            $include = Get-Content $IncludeData
            $include = $include | %{Get-ChildItem "$($UP.DataPath)\$($_.Trim('\'))" -Recurse -Force} | Get-ChildItem -Recurse -Force
            $UP.Data = $UP.Data | ? name -In $include.Name
            Append-Log 'Data include file applied.'
        } else {
            Append-Log 'Warning: Data include file not found. No user data will be copied.'
            $UP.Data = ''
        }
        $UP
    }catch{
        Append-Log 'Error processing data include file.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
function Exclude-Settings(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    try{
        if($ExcludeSettings -and $UP.Reg.Count -gt 1){
            Append-Log 'Processing settings exclude file.'
            $exclude = Get-Content $ExcludeSettings
            if($exclude){
                $Script:delete = $false
                $UP.Reg = $UP.Reg | %{
                    foreach ($line in $exclude){
                        if($line -eq ''){continue}
                        if($_.StartsWith("[$REGKEYPREFIX")){$Script:delete = $false}
                        if($_.StartsWith("[$REGKEYPREFIX$line")){$Script:delete = $true; break}
                    }
                    if(!$Script:delete){$_}
                }
            }
            Append-Log 'Settings exclude file applied.'
        } else {
            if($UP.Reg.Count -gt 1){Append-Log 'Warning: Settings exclude file not found.'}
        }
        $UP
    }catch{
        Append-Log 'Error processing settings include file.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
function Exclude-Data(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    try{
        if($ExcludeData -and $IncludeData){
            Append-Log 'Processing data exclude file.'
            $exclude = Get-Content $ExcludeData
            $exclude = $exclude | %{Get-ChildItem "$($UP.DataPath)\$($_.Trim('\'))" -Recurse -Force} | Get-ChildItem -Recurse -Force
            $UP.Data = $UP.Data | ? name -NotIn $exclude.Name
            Append-Log 'Data exclude file applied.'
        } else {
            if($IncludeData){Append-Log 'Warning: Data exclude file not found.'}
        }
        $UP
    }catch{
        Append-Log 'Error processing data exclude file.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
function Set-Settings(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    try{
        if($UP.Reg.Count -gt 1){
            Append-Log 'Importing registry settings.'
            Set-Content -Value $UP.Reg -Path $REGTMPPATH -Encoding Unicode
            reg.exe import $REGTMPPATH
            Append-Log 'User registry settings imported.'
        } else {
            Append-Log 'Warning: No registry settings found. Check include/exclude files.'
            Set-Content -Value $REGHEADER -Path $REGTMPPATH -Encoding Unicode
        }
        $UP
    }catch{
        Append-Log 'Error importing registry settings.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
function Set-Data(){
    param([Parameter(ValueFromPipeline=$true)][psobject]$UP)
    try{
        if($UP.Data){
            Append-Log 'Copying user data to local profile.'
            $UP.Data | %{
                $Destination = $env:USERPROFILE + $_.FullName.Substring($UP.DataPath.length)
                if($_.psiscontainer -and !(Test-Path $Destination -PathType Container)){
                    New-Item -Path $Destination -ItemType 'Container'
                } else {
                    Copy-Item $_.fullname -Destination $Destination -Force
                }
            }
            Append-Log 'User profile data copied.'
        } else {
            Append-Log 'Warning: No user data found. Check include/exclude files.'
        }
        $UP
    }catch{
        Append-Log 'Error copying files to local profile.'
        Append-Log $_.Exception.ItemName
        Append-Log $_.Exception.Message
        CleanUp-Session
        Die 'Aborting.'
    }
}
#endregion Functions

#region ## Main ##
Append-Log "Begin operation. User: $env:USERNAME on Computer: $env:COMPUTERNAME"
Validate-Params

if(!(Test-Path $REGTMPPATH -PathType Leaf) -or $Force){
    $UserProfile = Get-UserProfile
    if($UserProfile.Type -ne 'FS'){
        $UserProfile = $UserProfile | Get-Data | Include-Data | Exclude-Data | Set-Data
        $UserProfile = $UserProfile | Get-Settings | Include-Settings | Exclude-Settings | Set-Settings
        if($PassThru){$UserProfile}
    }
} else {
    Append-Log 'User has already been migrated.'
}
#endregion Main

#region ## Cleanup ##
if($UserProfile.Type -eq 'AM'){CleanUp-Session}
Append-Log 'Operation completed successfully.'
Append-Log 'Exiting.'
#endregion Cleanup