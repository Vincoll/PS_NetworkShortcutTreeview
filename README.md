# PS_NetworkShortcutTreeview
Based on Windows Explorer Network Location; Create Network shortcuts TreeView from .csv 

###Purpose:

The goal of this script is to quikly create a clean Shortcut Treeview for the Users in Windows File Explorer.
![Explorer View](https://gallery.technet.microsoft.com/scriptcenter/site/view/file/142227/1/ExplorerLeftSide.gif)

###How To Use :

Modify the CSV file. Add everything you want.
You can change the icon.
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


###Execution:

Set_NtwShortcut.ps1 will load ./NtwShortcut.csv
If the user have no access on one of the Ntw Path: No shortcut will be create for this Ntw Path
       
If execution through a GPO rule:
Silent Execution

If execution is manualy forced
A PS windows will show up showing status.

![Explorer View](https://gallery.technet.microsoft.com/scriptcenter/site/view/file/142228/1/PS_Display.gif)
