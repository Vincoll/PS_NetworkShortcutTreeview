
<#
    .SYNOPSIS   
        Based on Windows Explorer Network Location;
        Create Network shortcuts TreeView from .csv       
              
    .NOTES
        Only User browseable Paths will be create/display as Shortcut
        You can change the .ico

    .AUTHOR
        Vincent Collard (@Vincoll)
     
    .EXAMPLE 
        This script is design for being Exec at User Logon (GPO Rule)


    .csv Structure:
    RootFolder,Name,Path

        Exemple:
        RootFolder,Name,Path
        LOC1,Folder A,\\172.17.6.10\FOLDER_A
        LOC1,Folder B,\\SRV-STOCK\FOLDER_B
        LOC2,Folder A,\\172.17.90.6\FOLDER_A
        LOC2,Folder B,\\SRV-STOCK\FOLDER_B

    Explorer Structure:
    RootFolder\Name

        Exemple:
        .LOC1
        |       Folder A (Shortcut)=> \\172.17.6.10\FOLDER_A
        |       Folder B (Shortcut)=> \\SRV-STOCK1\FOLDER_B

        .LOC2
        |       Folder A (Shortcut)=> \\172.17.90.6\FOLDER_A
        |       Folder B (Shortcut)=> \\SRV-STOCK2\FOLDER_B

#>

#########################################
#GLOBAL VAR
#########################################

# Get the basepath for Network Shortcut
$shellApplication = New-Object -ComObject Shell.Application
$script:NetHoodPath = $shellApplication.Namespace(0x13).Self.Path #C:\Users\%username%\AppData\Roaming\Microsoft\Windows\Network Shortcuts 
$script:PSExecPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition # For PS <=2
$script:PathIco = "$PSExecPath\X.ico"

#Import csv
$Script:IMPORT_Loc = Import-Csv $PSExecPath\NtwShortcut.csv 


#region Functions

Function Add-NetworkShortcut
{
    param(
        [string]$RootFolder, #"LOC1"
        [string]$ExplorerName, #"Folder B"
        [string]$targetPath    #"\\SRV-STOCK1\FOLDER_B"
         )    
    
         $NiceExplorerName = Format-ShortcutTxt $ExplorerName # Format Explorer Name

    if (!(Test-Path "$NetHoodPath\$RootFolder"))
        {
            # Create the folderin ~\AppData\Roaming\Microsoft\Windows\Network Shortcuts\LOC1
            $CreateRootFolder = New-Item -Name $RootFolder -Path "$NetHoodPath" -type directory -ErrorAction SilentlyContinue
            # Create the ini file for ico
            Set-FolderIcon -Icon $PathIco -Path "$NetHoodPath\$RootFolder"
        }

            # Create the folder $ExplorerName in ~\AppData\Roaming\Microsoft\Windows\Network Shortcuts\LOC1\
            $CreateLink = New-Item -Name $NiceExplorerName -Path "$NetHoodPath\$RootFolder" -type directory   

                 # Create the ini file
            $desktopIniContent = '
                                        [.ShellClassInfo]
                                        CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}
                                        Flags=2
                                        ConfirmFileOp=1
'

                # Create the shortcut file
                $shortcut = (New-Object -ComObject WScript.Shell).Createshortcut("$NetHoodPath\$RootFolder\$NiceExplorerName\target.lnk")
                $shortcut.TargetPath = $targetPath
                $shortcut.IconLocation = "%SystemRoot%\system32\SHELL32.DLL, 275"
                $shortcut.Description = $targetPath
                $shortcut.WorkingDirectory = $targetPath
                $shortcut.Save()
        
                $desktopIniContent | Out-File -FilePath "$NetHoodPath\$RootFolder\$NiceExplorerName\Desktop.ini"

                # Set attributes on the files & folders
                Set-ItemProperty "$nethoodPath\$RootFolder\$NiceExplorerName\Desktop.ini" -Name Attributes -Value ([IO.FileAttributes]::System -bxor [IO.FileAttributes]::Hidden)
                Set-ItemProperty "$nethoodPath\$RootFolder\$NiceExplorerName" -Name Attributes -Value ([IO.FileAttributes]::ReadOnly)

    }# End Func

Function Remove-NetworkShortcut
{

    #Only Select Unique occurence in RootFolder
    $UniqueFolderDir = $IMPORT_Loc | Select-Object -ExpandProperty 'Rootfolder' | Sort-Object | Get-Unique #PS >3 # $IMPORT_Loc.RootFolder | Sort-Object | Get-Unique

    $LocationsIT = Get-ChildItem $NetHoodPath | Where-Object {$UniqueFolderDir -match $_.Name}
    $LocCount = $LocationsIT.count
    #Remove "LOC1-*" "LOC2-*"
    ForEach ($loc in $LocationsIT)
        {Remove-Item -Path $loc.FullName -Recurse -Force}
        
}

Function Format-ShortcutTxt
{
 param([string]$ShortcutName)

    $SplitName = $ShortcutName.Split('-')# Folder A-SRV-STOCK1
    $ShortcutName = $SplitName[0]# Folder A

    $TextInfo = (Get-Culture).TextInfo
        #Try{# $ShortShortcutName = $ShortShortcutName.Replace('_',' ')} #   Catch{}     
    $NiceString = $TextInfo.ToTitleCase($ShortcutName.ToLower()); #Folder A
       
    Return $NiceString 
}

function Test_Shortcut
{
    param([string]$Pth)

    Try
        {$ACL = Get-Acl $Pth -ErrorAction Stop}
    Catch
        {#If ACL attributs can not be read = No Right > No Shortcut
         Return $False}
    
    Return $True
}
#endregion Functions


#region Icon

function Set-FolderIcon 
{ #https://gallery.technet.microsoft.com/scriptcenter/Set-FolderIcon-0bd56629
    [CmdletBinding()] 
    param 
    (     
        [Parameter(Mandatory=$True, 
        Position=0)] 
        [string[]]$Icon, 
        [Parameter(Mandatory=$True, 
        Position=1)] 
        [string]$Path, 
        [Parameter(Mandatory=$False)] 
        [switch] 
        $Recurse     
    ) 
    BEGIN 
    { 
        $originallocale = $PWD 
        #Creating content of the DESKTOP.INI file. 
        $ini = '[.ShellClassInfo] 
                IconFile=folder.ico 
                IconIndex=0 
                ConfirmFileOp=0' 
        Set-Location $Path 
        Set-Location ..     
        Get-ChildItem | Where-Object {$_.FullName -eq "$Path"} | ForEach {$_.Attributes = 'Directory, System'} 
    }     
    PROCESS 
    { 
        $ini | Out-File $Path\DESKTOP.INI 
        If ($Recurse -eq $True) 
        { 
            Copy-Item -Path $Icon -Destination $Path\FOLDER.ICO     
            $recursepath = Get-ChildItem $Path -r | Where-Object {$_.Attributes -match "Directory"} 
            ForEach ($folder in $recursepath) 
            { 
                Set-FolderIcon -Icon $Icon -Path $folder.FullName 
            } 
         
        } 
        else 
        { 
            Copy-Item -Path $Icon -Destination $Path\FOLDER.ICO 
        }     
    }     
    END 
    { 
        $inifile = Get-Item $Path\DESKTOP.INI 
        $inifile.Attributes = 'Hidden,System' 
        $icofile = Get-Item $Path\FOLDER.ICO 
        $icofile.Attributes = 'Hidden,System' 
        Set-Location $originallocale         
    } 
} 

function Remove-SetIcon 
{ 
    [CmdletBinding()] 
    param 
    (     
        [Parameter(Mandatory=$True, 
        Position=0)] 
        [string]$Path 
    ) 
    BEGIN 
    { 
        $originallocale = $PWD 
        $iconfiles = Get-ChildItem $Path -Recurse -Force | Where-Object {$_.Name -like "FOLDER.ICO"} 
        $iconfiles = $iconfiles.FullName 
        $inifiles = Get-ChildItem $Path -Recurse -Force | where-Object {$_.Name -like "DESKTOP.INI"} 
        $inifiles = $inifiles.FullName 
    } 
    PROCESS 
    { 
        Remove-Item $iconfiles -Force 
        Remove-Item $inifiles -Force 
        Set-Location $Path 
        Set-Location .. 
        Get-ChildItem | Where-Object {$_.FullName -eq "$Path"} | ForEach {$_.Attributes = 'Directory'}     
    } 
    END 
    { 
        Set-Location $originallocale 
    } 
}

#endregion


Function Main {

#Delete Windows Network Shortcut based on .csv for a clean Re/Start
Remove-NetworkShortcut

$Locations_Count =$IMPORT_Loc.count ; $i=1
Write-Host 'Importing Network Shortcut.' `n -ForegroundColor Green

Foreach($shortcut in $IMPORT_Loc)#Progression
    {
        $Progress = "[$i/$Locations_Count]"
        if( Test_Shortcut $shortcut.Path) #Check if Access Rights; Only create shortcut if browseable
        {
            Write-Host '' $Progress `t $shortcut.RootFolder $shortcut.Name 
            Add-NetworkShortcut $shortcut.RootFolder $shortcut.Name $shortcut.Path # True : Create folder
        }
        else
        {
            Write-Host '' $Progress `t $shortcut.RootFolder $shortcut.Name -ForegroundColor DarkGray # False : Do not create
        }
        $i++
    }

Write-Host ''; Write-Host 'Import done.' -ForegroundColor Cyan
Start-Sleep -Seconds 3

}

#START SCRIPT
Main