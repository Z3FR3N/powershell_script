#############
# FUNCTIONS #
#############

function Separator {
    <#
    .SYNOPSIS
    A separator
    .DESCRIPTION
    A simple separator which confine statement of the script
    #>
    Write-Host "`n%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%`n "
}

Function Get-Folder{
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # InitialDirectory help description
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Cartella iniziale",
            Position = 0)]
        [String]$SelectedPath,

        # Description help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Messaggop del titolo")]
        [String]$Description="Seleziona una cartella",

        # ShowNewFolderButton help description
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Show New Folder Button when used")]
        [Switch]$ShowNewFolderButton
    )

    # Load Assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Open Class
    $FolderBrowser= New-Object System.Windows.Forms.FolderBrowserDialog

   # Define Title
    $FolderBrowser.Description = $Description

    # Define Initial Directory
    if (-not [String]::IsNullOrWhiteSpace($SelectedPath))
    {
        $FolderBrowser.SelectedPath=$SelectedPath
    }

    $FiledialogResult = $folderBrowser.ShowDialog()

    if($FiledialogResult -eq 'OK')
    {
        $Folder += $FolderBrowser.SelectedPath
    } 
    
    elseif($FiledialogResult -eq 'Cancel')
    {
        exit
    }

    return $Folder
}

function Get-FileName
{
<#
.SYNOPSIS
   Show an Open File Dialog and return the file selected by the user

.DESCRIPTION
   Show an Open File Dialog and return the file selected by the user

.PARAMETER WindowTitle
   Message Box title
   Mandatory - [String]

.PARAMETER InitialDirectory
   Initial Directory for browsing
   Mandatory - [string]

.PARAMETER Filter
   Filter to apply
   Optional - [string]

.PARAMETER AllowMultiSelect
   Allow multi file selection
   Optional - switch

 .EXAMPLE
   Get-FileName
    cmdlet Get-FileName at position 1 of the command pipeline
    Provide values for the following parameters:
    WindowTitle: My Dialog Box
    InitialDirectory: c:\temp
    C:\Temp\42258.txt

    No passthru paramater then function requires the mandatory parameters (WindowsTitle and InitialDirectory)

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp
   C:\Temp\41553.txt

   Choose only one file. All files extensions are allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect
   C:\Temp\8544.txt
   C:\Temp\42258.txt

   Choose multiple files. All files are allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "text file (*.txt) | *.txt"
   C:\Temp\AES_PASSWORD_FILE.txt

   Choose multiple files but only one specific extension (here : .txt) is allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "Text files (*.txt)|*.txt| csv files (*.csv)|*.csv | log files (*.log) | *.log"
   C:\Temp\logrobo.log
   C:\Temp\mylogfile.log

   Choose multiple file with the same extension

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "selected extensions (*.txt, *.log) | *.txt;*.log"
   C:\Temp\IPAddresses.txt
   C:\Temp\log.log

   Choose multiple file with different extensions
   Nota :It's important to have no white space in the extension name if you want to show them

.EXAMPLE
 Get-Help Get-FileName -Full

.INPUTS
   System.String
   System.Management.Automation.SwitchParameter

.OUTPUTS
   System.String

.NOTESs
  Version         : 1.0
  Author          : O. FERRIERE
  Creation Date   : 11/09/2019
  Purpose/Change  : Initial development

  Based on different pages :
   mainly based on https://blog.danskingdom.com/powershell-multi-line-input-box-dialog-open-file-dialog-folder-browser-dialog-input-box-and-message-box/
   https://code.adonline.id.au/folder-file-browser-dialogues-powershell/
   https://thomasrayner.ca/open-file-dialog-box-in-powershell/
#>
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # WindowsTitle help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Message Box Title",
            Position = 0)]
        [String]$WindowTitle,

        # InitialDirectory help description
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Initial Directory for browsing",
            Position = 1)]
        [String]$InitialDirectory,

        # Filter help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Filter to apply",
            Position = 2)]
        [String]$Filter = "All files (*.*)|*.*",

        # AllowMultiSelect help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Allow multi files selection",
            Position = 3)]
        [Switch]$AllowMultiSelect
    )

    # Load Assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Open Class
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog

    # Define Title
    $OpenFileDialog.Title = $WindowTitle

    # Define Initial Directory
    if (-Not [String]::IsNullOrWhiteSpace($InitialDirectory))
    {
        $OpenFileDialog.InitialDirectory = $InitialDirectory
    }

    # Define Filter
    $OpenFileDialog.Filter = $Filter

    # Check If Multi-select if used
    if ($AllowMultiSelect)
    {
        $OpenFileDialog.MultiSelect = $true
    }
    $OpenFileDialog.ShowHelp = $true    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.
    $OpenFileDialog.ShowDialog() | Out-Null
    if ($AllowMultiSelect)
    {
        return $OpenFileDialog.Filenames
    }
    else
    {
        return $OpenFileDialog.Filename
    }
}

function ListingFiles {
    <#
    .SYNOPSIS
    Function which generate a csv with all file complete names.
    .DESCRIPTION
    Asks for a folder from the user and read the content of the folder, printing it in a csv in the same folder.
    .OUTPUTS
    A csv file containing all file's names.
    .NOTES
    Feature request: adding a revert function to recover the previous state
    Better error handling and messagging
    #>
    Separator
    Read-Host "Questa sezione genera un elenco in .csv dei file della cartella in cui si trova.`nPremere invio per continuare" | Out-Null
    Separator

    Write-Host "Selezionare una cartella."
    

    $current_folder = Get-Folder $PSScriptRoot
    
    $keep_alive = $true
    
    # input validation
    
    do {
        Separator
        $ans = Read-Host "La cartella selezionata e': $current_folder`nProcedere?`n`n 1 >> SI`n 2 >> CAMBIO CARTELLA`n 3 >> USCITA`n 4 >> MENU' PRINCIPALE`n`nDigitare il numero e premere invio"

        switch ($ans) {
            1 { $keep_alive = $false; break }
            2 { $current_folder = Get-Folder($PSScriptRoot); break }
            3 { $keep_alive = $false
                return 3; break }
            4 { $keep_alive = $false
                return 4; break }
            Default {Write-Host "Inserire un valore corretto"}
        }
    } while ($keep_alive)
    
    $files = Get-ChildItem -Path $current_folder -File -Name
    $names = New-Object System.Collections.ArrayList

    foreach ($file in $files) {
        # Using custom PSObject due to mishandling of hashtables by Export-CSV (PS version 5.1)
        $hashTable = New-Object -TypeName PSObject -Property @{
            NOME = $file
        }
        $names.Add($hashTable) | Out-Null
    }

    $names | Export-Csv -Path $current_folder\'nomi files.csv' -NoTypeInformation -Delimiter ';' -Encoding UTF8

    Start-Sleep -Seconds 1
    
    Separator
    Write-Host 'Fatto.'

    return 0
}

function ChangingFileNames {
    <#
    .SYNOPSIS
    Main function which display the content of the file
    .DESCRIPTION
    Asks for a folder from the user and read the content of the folder, printing it in a csv in the same folder
    .NOTES
    Feature request: adding a revert function to recover the previous state
    #>
    
    $folder = Get-Folder $PSScriptRoot
    $csv_path = Get-FileName "Seleziona un file csv" $PSScriptRoot "File (*.csv)|*.csv" $false # it works
    Separator

    Write-Host "Ecco la prima riga, controlla il separatore:`n"
    Get-Content -Path $csv_path -TotalCount 1
    Separator

    $sep = Read-Host "Inserisci il separatore corretto"
    $csv = Import-Csv -Path $csv_path -Delimiter $sep
    Separator

    # Per chiarezza, si stampano le colonne con un contatore affianco
    $count = 1
    Write-Host ("Ecco le colonne:")
    foreach($col in $csv[0].psobject.properties.name){
        Write-Host ("`t`t" + ($count++) +")" + " " + "`"" + $col +"`"") # array che ritorna le colonne
    }

    # Il numero selezionato corrisponde ad (indice_array + 1)
    # Per brevità, si usa un semplice contatore sottraendo 1
    $count = Read-Host "Inserisci il numero corrispondete alla colonna dei nomi da verificare`n>>"
    $oldname_col = $csv[0].psobject.properties.name[$count -1] # prelevo il valore dall'array
    Write-Host ("Colonna selezionata >> " + "`"" + $oldname_col + "`"")
    Separator

    $count = Read-Host "Inserisci il numero corrispondete alla colonna dei nomi da sostituire`n>>"
    $new_name_col = $csv[0].psobject.properties.name[$count -1]
    Write-Host (("Colonna selezionata "), ("`"" + $new_name_col + "`"")) -Separator " >> "
    Separator

    # Si scansiona la cartella alla ricerca dei file da rinominare
    $format = Read-Host "Inserisci il formato dei file da rinominare (senza il punto)"
    $n_file = 0
    Get-ChildItem -Path $folder -File -Filter ("*." + $format) | ForEach-Object {
        foreach($row in $csv){
            # String cleaning per un confronto accurato
            $file_name = $_.Name.ToString().ToLower()
            $old_name = $row.$oldname_col.ToString().Trim().ToLower()

            if($old_name -eq $file_name){
                $new_name = $row.$new_name_col.Trim() # trimmed version to avoid extra spaces
                $complete_name = "{0}{1}" -f $new_name, $_.Extension
                $n_file++
                Rename-Item -Path ($folder + "\" + $_.Name) -NewName $complete_name
            }
        }
    }

    Write-Host "Rinominati " $n_file " file."
    return 0
}

########
# MAIN #
########

<#
    .SYNOPSIS
    Script to automate 2 main function: generation of a csv file with the content of a folder and renaming all the files of a specific folder given a csv with names
    .DESCRIPTION
    The script automate some actions to speed up file handling and management
    .OUTPUTS
    A csv file containing all file's names.
    .NOTES
    Feature request: adding a revert function to recover the previous state
    Better error handling and messagging
#>

[console]::WindowHeight=30
[console]::WindowWidth=85
Write-Host 
"       +++++++++++++++++++++++++++++++++++++=:...:=++++++++++++++++++++++++
        +++++++++++++++++++++++++++++++++++..       ..=+++++++++++++++++++++
        +++++++++++++++++++++++++++++++++-. .:=+++=.. .-++++++++++++++++++++
        ++++++++++++++++++++++++++++++++-. .=+++++++=. .-+++++++++++++++++++
        +++++++++++=...   ...=+++++.:+++:..-+++::++++-  :+++++++++++++++++++
        ++++++++++.  ..:-:.. ..=++.  .=++++++:. .:+++:  :+++++++++++++++++++
        +++++++++. .:+++++++:  .+++-  .+++++-. .=+++:. .=+++++++++++++++++++
        ++++++++-  .+++++++++:  :+++:  .++++. .-+++:. .=++++++++++++++++++++
        ++++++++:  :+++:  :+++++++++:  :++++. .-++++:=+++. .-+++++++++++++++
        ++++++++=. .=++=. .-+++++++:  .+++++=.  -+++++++:. .++++++++++++++++
        +++++++++-. .:++=.  .:===:   .=++++++-.  .:==-:.  .+++:.    .:=+++++
        ++++++++++=..+++++:..     ..-++++++++++:.      ..-++++=......  .=+++
        ++++++++++++++++++++++---++++++++++++++++++=-=+++++++++++++++:. .=++
        ++++++++++=-::-=++++++++++++++++=::.:-=+++++++++++++-:.::=++++:  .++
        +++++++++:.      .-++++++++++-.      .:++++++++++:.      .-+++=  .++
        ++++++++++=++++:.  .++++++++.  .-++++=+++++++++=. ..-++++=++++:  :++
        +++++++++++++++++:  :++++++:  :++++++++++++++++.  -+++++++++=.. .=++
        +++++:.      .++++.  ++++++.  =+++++++++++++++=  .++++-.       :++++
        +++-.  ..-:..++++=. .++++++.  -+++++++++++++++=. .++++-.    .-++++++
        ++:  .=+++++++++=.  -++++++:  .=+++++++++++++++:  .=++++++++++++++++
        ++.  =++++:::::.  .-++++++++-. ...:::.++++++++++:. ..:::.:++++++++++
        ++   ++++.      .:++++++++++++-.      .+++++++++++:.      :+++++++++
        ++.  :++++++==++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ++=.  .-++++=+++++++-:...:-++++++++++++++:....:-+++++++=++++++++++++
        ++++-.      .-+++=..        .=++++++++=.        ..=+++:..-++++++++++
        +++++++-:.::-+++:. .-++++=-.  :++++++:  .-+++++:. .:++=.  :+++++++++
        +++++++++++++++-  .+++++++++. .-++++-. .++++++++=. .=+++. .=++++++++
        +++++++++++++++:..-++=..++++-  .++++.  -+++:::=++:..:+++:  -++++++++
        ++++++++++++++++++++:  .:+++:  :++++:  :+++:  .++++++++=. .=++++++++
        +++++++++++++++++++-  .++++:  .++++++.  :+++. ..=+++++:. .:+++++++++
        +++++++++++++++++++. .=+++:...++++++++...-+++:.   ...   .=++++++++++
        +++++++++++++++++++. .-++++-=+++.  -++++=++++++=:.....:+++++++++++++
        +++++++++++++++++++=.  -+++++++.  .+++++++++++++++++++++++++++++++++
        ++++++++++++++++++++=.   .--:.  ..++++++++++++++++++++++++++++++++++
        ++++++++++++++++++++++-.........=+++++++++++++++++++++++++++++++++++
        
"

do{
    Separator
    $ans = Read-Host "Selezionare l'applicazione da eseguire, e' possibile uscire in qualsiasi momento`npremendo la combinazione di tasti Ctrl + C. Selezionare una funzione:`n`n 1 >> Generazione del file di nomi`n 2 >> Rinominare in massa i file`n 3 >> Uscita`n`nDigitare un numero e premere invio"

    if ($ans -eq 1) {
     $ans = ListingFiles
    }
   
    elseif ($ans -eq 2) {
     $ans = ChangingFileNames
    }
   
    Separator
} while ($ans -ne 3)

Separator
Write-Host "...Uscita..."
Separator
Start-Sleep 2