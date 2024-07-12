#############
# FUNCTIONS #
#############

function SetConsoleSize {
    <#
    .Synopsis
    Set the size of the current console window

    .Description
    Set-ConsoleSize sets or resets the size of the current console window. By default, it
    sets the window to a height of 40 lines, with a 2000 line buffer, and sets the 
    the width and width buffer to 120 characters. 

    .Example
    Restore the console window to 40h x 120w:
    Set-ConsoleSize

    .Example
    Change the current console to a height = 30 lines and a width = 180 chars:
    Set-ConsoleSize -Height 30 -Width 180

    .Parameter Height
    The number of lines to which to set the current console. Default = 40 lines. 

    .Parameter Width
    The number of characters to which to set the current console. Default = 120 chars.

    .Inputs
    [int]
    [int]

    .Notes
        Author: ss64.com/ps/syntax-consolesize.html
     Last edit: 2019-08-29
    #>
    [CmdletBinding()]
    Param(
         [Parameter(Mandatory=$False,Position=0)]
         [int]$Height = 40,
         [Parameter(Mandatory=$False,Position=1)]
         [int]$Width = 120
         )
    $console = $host.ui.rawui
    $ConBuffer  = $console.BufferSize
    $ConSize = $console.WindowSize

    $currWidth = $ConSize.Width
    $currHeight = $ConSize.Height

    # if height is too large, set to max allowed size
    if ($Height -gt $host.UI.RawUI.MaxPhysicalWindowSize.Height) {
        $Height = $host.UI.RawUI.MaxPhysicalWindowSize.Height
    }

    # if width is too large, set to max allowed size
    if ($Width -gt $host.UI.RawUI.MaxPhysicalWindowSize.Width) {
        $Width = $host.UI.RawUI.MaxPhysicalWindowSize.Width
    }

    # If the Buffer is wider than the new console setting, first reduce the width
    If ($ConBuffer.Width -gt $Width ) {
       $currWidth = $Width
    }
    # If the Buffer is higher than the new console setting, first reduce the height
    If ($ConBuffer.Height -gt $Height ) {
        $currHeight = $Height
    }
    # initial resizing if needed
    $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($currWidth,$currHeight)

    # Set the Buffer
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($Width,2000)

    # Now set the WindowSize
    $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($Width,$Height)

    # Display the new sizes (Optional/for debugging)
    # "Height: " + $host.ui.rawui.WindowSize.Height
    # "Width:  " + $host.ui.rawui.WindowSize.width
}

function Separator {
    <#
    .SYNOPSIS
    A separator
    .DESCRIPTION
    A simple separator which confine statement of the script
    #>
    Write-Host "`n%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%`n "
}

Function Get-Folder {
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

    if ($FiledialogResult -eq 'Cancel') {
        Separator
        Write-Host "NESSUNA CARTELLA SELEZIONATA, RITORNO AL MENU' PRINCIPALE"
        Start-Sleep 1
        Main 0
    }

    return $Folder
}

function Get-FileName {
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

    .NOTES
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
    $FiledialogResult = $OpenFileDialog.ShowDialog()

    if ($FiledialogResult -eq 'OK') {
        if ($AllowMultiSelect)
        {
            return $OpenFileDialog.Filenames
        }
        else
        {
            return $OpenFileDialog.Filename
        }
    } elseif ($FiledialogResult -eq 'Cancel') {
        Separator
        Write-Host "NESSUNA FILE SELEZIONATO, RITORNO AL MENU' PRINCIPALE"
        Start-Sleep 1
        Main 0
    }
}

function Save-File(){

    [CmdletBinding()]
    [OutputType([string])]
    Param ([Parameter(
        Mandatory = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Inital directory",
        Position = 3)]
    [string]$initialDirectory)
    
    Add-Type -AssemblyName System.Windows.Forms

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.DefaultExt = '.csv'
    $OpenFileDialog.filter = "File csv (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null

    if (-not [string]::IsNullOrWhiteSpace($OpenFileDialog.Filename)) {
        return $OpenFileDialog.filename
    } else {
        Separator
        Write-Host "SELEZIONARE UN FILE."
        Start-Sleep 0.5
        return 0
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
    [CmdletBinding()]
    [OutputType([string])]
    param()

    Separator
    Read-Host "Questa sezione genera un elenco dei file di una cartella, salvando i risultati in un file csv per la consultazione.`n`nPREMERE INVIO PER CONTINUARE" | Out-Null
    Separator

    Write-Host "SELEZIONARE LA CARTELLA CONTENENTE I FILES DA INDIVIDUARE"

    $current_folder = Get-Folder $PSScriptRoot
    
    $keep_alive = $true

    $selected_folder = $current_folder.Split('\')[-1]
     
    do {
        Separator
        $ans = Read-Host "Cartella selezionata: `"$selected_folder`"`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio cartella`n`t`t3 >> Menu' principale`n`t`t4 >> Uscita`n`nDIGITARE UN NUMERO E PREMERE INVIO"

        switch ($ans) {
            1 { $keep_alive = $false; break }
            2 { $current_folder = Get-Folder $PSScriptRoot 
                $selected_folder = $current_folder.Split('\')[-1] ; break }
            3 { return 0 }
            4 { return 4 }
            Default {Write-Host "INSERIRE UN VALORE CORRETTO"}
        }
    } while ($keep_alive)
    
    Separator
    $format =  Read-Host "Inserire il formato dei files da individuare.`nPremere invio per individuare TUTTI i file della cartella"

    if (($format.Length -ge 3) -and (-not $format.Contains("."))) {
        $format = "." + $format
    } else {
        Separator
        Write-Host "NESSUN FORMATO RICONOSCIUTO, VERRRANNO INDIVIDUATI TUTTI I FILES."
        $format = ".*"
    }
    
    #rimuovere il punto dai file
    $files = Get-ChildItem -Path $current_folder -File -Name -Filter ("*" + $format.ToLower())
    $names = New-Object System.Collections.ArrayList
    if ($files.Count -eq 0) {
        Separator
        Write-Host "NESSUN FILE TROVATO, RITORNO AL MENU' PRINCIPALE"
        Start-Sleep 1
        Main 0
    } else {
        Separator
        Write-Host "TROVATI" $files.Count "FILES"
        Start-Sleep 1
    }

    foreach ($file in $files) {
        # Using custom PSObject due to mishandling of hashtables by Export-CSV (PS version 5.1)
        $hashTable = New-Object -TypeName PSObject -Property @{ NOME = $file }
        $names.Add($hashTable) | Out-Null
    }

    Separator
    Write-Host "FILE CSV GENERATO, INDICARE UNA POSIZIONE DOVE SALVARE"
    Start-Sleep 1
    $save_file = Save-File $current_folder

    if (-not ($save_file -eq 0)) {
        $names | Export-Csv -Path $save_file -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Separator
        Write-Host "FILE CSV SALVATO, RITORNO AL MENU' PRINCIPALE"
        Start-Sleep 1
        return 0
    } else {
        Separator
        Write-Host "RITORNO AL MENU' PRINCIPALE"
        Start-Sleep 1
        return 0
    }
}

function  ChangingFileNames {
    <#
    .SYNOPSIS
    Main function which display the content of the file

    .DESCRIPTION
    Asks for a folder from the user and read the content of the folder, printing it in a csv in the same folder

    .NOTES
    Feature request: adding a revert function to recover the previous state
    #>

    Separator
    Read-Host "Questa sezione permette di rinominare in maniera massiva i file `ndi una cartella partendo da un file csv dove, in due colonne`ndistinte, vengono accoppiati i nomi attualmente assegnati`ned i nomi che verranno rimpiazzati.`n`nE' FONDAMENTALE che i rispettivi nomi contengano anche l'ESTENSIONE`n(.pdf,.csv,.png, etc...).`n`nES: nome documento.pdf    ->    nuovo nome documento.pdf`n`n`PREMERE INVIO PER CONTINUARE" | Out-Null
    Separator

    $keep_alive = $true
    Write-Host "SELEZIONARE il FILE CSV CONTENENTE I NOMI DA RINOMINARE"
    $csv_path = Get-FileName "Seleziona un file csv" $PSScriptRoot "File (*.csv)|*.csv" $false
    $selected_csv = $csv_path.Split('\')[-1]
    Separator
    Write-Host "SELEZIONARE LA CARTELLA CONTENENTE I FILES"
    $folder = Get-Folder $PSScriptRoot 
    $selected_folder = $folder.Split('\')[-1]

    do {
        Separator
        $ans = Read-Host "Cartella selezionata: `"$selected_folder`"`n`nFile csv: `"$selected_csv`"`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio cartella`n`t`t3 >> Cambio File`n`t`t4 >> Menu' principale`n`t`t5 >> Uscita`n`nDIGITARE UN NUMERO E PREMERE INVIO"

        switch ($ans) {
            1 { $keep_alive = $false; break }
            2 { $folder = Get-Folder $PSScriptRoot
                $selected_folder = $folder.Split('\')[-1]; break }
            3 { [string]$csv_path = Get-FileName "Seleziona un file csv" $PSScriptRoot "File (*.csv)|*.csv" $false
                $selected_csv = $csv_path.Split('\')[-1]; break }
            4 { return 0 }
            5 { return 4 }
            Default {
                        Separator
                        Write-Host "INSERIRE UN VALORE CORRETTO"}
        }
        
    } while ($keep_alive)

    if (-not [string]::IsNullOrWhiteSpace($csv_path)) {
        $keep_alive = $true
        $delimiter = ','
        do {
            Separator
            $csv = Import-Csv -Path $csv_path -Delimiter $delimiter
            $count = 1
            Write-Host ("Ecco le colonne individuate:`n")
            foreach($col in $csv[0].psobject.properties.name){
                Write-Host ("`t`t" + ($count++) +")" + " " + "`"" + $col +"`"")
            }
            
            Separator
            $ans = Read-Host "Sono corrette?`n`n`t`t1 >> Si`n`t`t2 >> No, cambio separatore`n"
            switch ($ans) {
                1 { $keep_alive = $false; break }
                2 { $delimiter = Read-Host "INSERIRE IL NUOVO SEPARATORE"; break}
                Default {   Separator
                            Write-Host "Inserire un valore corretto."
                            Start-Sleep 1}
            }

        } while($keep_alive)

        $keep_alive = $true
        do {
            # Il numero selezionato corrisponde ad (indice_array + 1)
            # Per brevità, si usa un semplice contatore sottraendo 1
            try {
                [int]$col_num = Read-Host "`nInserire il numero della colonna dei nomi DI ORIGINE"
                if ($col_num -ge 1 -and $col_num -le $count) {
                    $oldname_col = $csv[0].psobject.properties.name[$col_num -1] # prelevo il valore dall'array
                    Write-Host ("`n`t`tCOLONNA SELEZIONATA >> " + "`"" + $oldname_col + "`"")
                } else {
                    Separator
                    Write-Host "INSERIRE UN NUMERO CORRETTO"
                }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO"
            }

            try {
                [int]$col_num = Read-Host "`nInserire il numero della colonna dei nomi DA RINOMINARE"
                    if ($col_num -ge 1 -and $col_num -le $count) {
                        $new_name_col = $csv[0].psobject.properties.name[$col_num -1]
                        Write-Host (("`n`t`tCOLONNA SELEZIONATA"), ("`"" + $new_name_col + "`"")) -Separator " >> "
                    } else {
                        Separator
                        Write-Host "INSERIRE UN NUMERO CORRETTO"
                    }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO"
            }

            Separator
            $ans = Read-Host "Ricapitolando, i nomi nella colonna `"$oldname_col`" verranno rimpiazzati`ncon i nomi contenuti nella colonna `"$new_name_col`". Procedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio colonne`n`nDIGITARE UN NUMERO E PREMERE INVIO"
            
            switch ($ans) {
                1 { $keep_alive = $false; break }
                2 { Write-Host ("`nEcco le colonne individuate:`n")
                    $count = 1
                    foreach($col in $csv[0].psobject.properties.name){
                    Write-Host ("`t`t" + ($count++) +")" + " " + "`"" + $col +"`"")
                    }; break}
                Default {Write-Host "INSERIRE UN VALORE CORRETTO."}
            }
            
        }while($keep_alive)

        $n_file = 0
        Get-ChildItem -Path $folder -File | ForEach-Object {
            foreach($row in $csv){
                # String cleaning per un confronto accurato
                $file_name = $_.Name.ToString().ToLower()
                $old_name = $row.$oldname_col.ToString().Trim().ToLower()

                if($old_name -eq $file_name){
                    $new_name = $row.$new_name_col.Trim() # trimmed version to avoid extra spaces
                    if ($new_name.Contains(".")) { 
                        $new_name = $new_name.Split(".")[0]
                    }
                    $complete_name = "{0}{1}" -f $new_name, $_.Extension
                    $n_file++
                    Rename-Item -Path ($folder + "\" + $_.Name) -NewName $complete_name
                }

            }
        }

        Write-Host "RINOMINATI " $n_file " FILES"
        return 0
    } else {
        Separator
        Write-Host "FILE NON VALIDO"
        return 0
    }

}

function MoveFile {
    # Caricare un file CSV che contiene dei nomi di file da spostare ed indicare una cartella dove spostarli.
    Read-Host "Questo modulo permette di dividere dei file in cartelle.  `n`nATTENZIONE: le cartelle devono trovarsi nella stessa `"cartella madre`". `n`nAd esempio:`n`nnome_documento.pdf    ->    cartella madre/NOMECARTELLA`nnome_documento2.pdf    ->    cartella madre/ALTRACARTELLA" | Out-Null
    Separator

    $keep_alive = $true
    Write-Host "SELEZIONARE il FILE CSV CONTENENTE I NOMI DA RINOMINARE"
    $csv_path = Get-FileName "Seleziona un file csv" $PSScriptRoot "File (*.csv)|*.csv" $false
    $selected_csv = $csv_path.Split('\')[-1]
    Separator
    Write-Host "SELEZIONARE LA CARTELLA CONTENENTE I FILES"
    $folder = Get-Folder $PSScriptRoot 
    $selected_folder = $folder.Split('\')[-1]

    do {
        Separator
        $ans = Read-Host "Cartella selezionata: `"$selected_folder`"`n`nFile csv: `"$selected_csv`"`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio cartella`n`t`t3 >> Cambio File`n`t`t4 >> Menu' principale`n`t`t5 >> Uscita`n`nDIGITARE UN NUMERO E PREMERE INVIO"

        switch ($ans) {
            1 { $keep_alive = $false; break }
            2 { $folder = Get-Folder $PSScriptRoot
                $selected_folder = $folder.Split('\')[-1]; break }
            3 { [string]$csv_path = Get-FileName "Seleziona un file csv" $PSScriptRoot "File (*.csv)|*.csv" $false
                $selected_csv = $csv_path.Split('\')[-1]; break }
            4 { return 0 }
            5 { return 4 }
            Default {
                        Separator
                        Write-Host "INSERIRE UN VALORE CORRETTO"}
        }
        
    } while ($keep_alive)

    if (-not [string]::IsNullOrWhiteSpace($csv_path)) {
        $keep_alive = $true
        $delimiter = ','
        do {
            Separator
            $csv = Import-Csv -Path $csv_path -Delimiter $delimiter
            $count = 1
            Write-Host ("Ecco le colonne individuate:`n")
            foreach($col in $csv[0].psobject.properties.name){
                Write-Host ("`t`t" + ($count++) +")" + " " + "`"" + $col +"`"")
            }
            
            Separator
            $ans = Read-Host "Sono corrette?`n`n`t`t1 >> Si`n`t`t2 >> No, cambio separatore`n"
            switch ($ans) {
                1 { $keep_alive = $false; break }
                2 { $delimiter = Read-Host "INSERIRE IL NUOVO SEPARATORE"; break}
                Default {   Separator
                            Write-Host "Inserire un valore corretto."
                            Start-Sleep 1}
            }

        } while($keep_alive)
    
        $keep_alive = $true
        do {
            # Il numero selezionato corrisponde ad (indice_array + 1)
            # Per brevità, si usa un semplice contatore sottraendo 1
            try {
                [int]$col_num = Read-Host "`nInserire il numero della colonna dei nomi dei FILES DA SPOSTARE"
                if ($col_num -ge 1 -and $col_num -le $count) {
                    $name_col = $csv[0].psobject.properties.name[$col_num -1] # prelevo il valore dall'array
                    Write-Host ("`n`t`tCOLONNA SELEZIONATA >> " + "`"" + $name_col + "`"")
                } else {
                    Separator
                    Write-Host "INSERIRE UN NUMERO CORRETTO"
                }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO"
            }

            try {
                [int]$col_num = Read-Host "`nInserire il numero della colonna contenente il nome della CARTELLA"
                    if ($col_num -ge 1 -and $col_num -le $count) {
                        $folder_col = $csv[0].psobject.properties.name[$col_num -1]
                        Write-Host (("`n`t`tCOLONNA SELEZIONATA"), ("`"" + $folder_col + "`"")) -Separator " >> "
                    } else {
                        Separator
                        Write-Host "INSERIRE UN NUMERO CORRETTO"
                    }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO"
            }

            Separator
            $ans = Read-Host "Ricapitolando, i file nella colonna `"$name_col`" verranno spostati`nnelle cartelle contenute in `"$folder_col`". Se la cartella non esiste, verrà creata.`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio colonne`n`nDIGITARE UN NUMERO E PREMERE INVIO"
            
            switch ($ans) {
                1 { $keep_alive = $false; break }
                2 { Write-Host ("`nEcco le colonne individuate:`n")
                    $count = 1
                    foreach($col in $csv[0].psobject.properties.name){
                    Write-Host ("`t`t" + ($count++) +")" + " " + "`"" + $col +"`"")
                    }; break}
                Default {Write-Host "INSERIRE UN VALORE CORRETTO."}
            }
            
        }while($keep_alive)

        Get-ChildItem -Path $folder -File | ForEach-Object {
            foreach($row in $csv){
                # String cleaning per un confronto accurato
                $file_name = $_.Name
                $file_to_move = $row.$name_col
                $folder_name = $row.$folder_col

                if (-not (Test-Path -Path ($folder + "\" + $folder_name))) {
                    New-Item -Path $folder -Name $folder_name -ItemType "directory"
                }

                if($file_name -eq $file_to_move){
                    Move-Item $_.FullName $folder_name
                    $n_file++
                }
            }
        }

        Write-Host "SPOSTATI " $n_file " FILE."
    } else {
        Separator
        Write-Host "FILE NON VALIDO"
        return 0
    }
}

function License {

}
function Main {
    param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "ID entry",
            Position = 0)]
        [string]$ans
    )

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

    do{
        Clear-Host
        Separator
        $ans = Read-Host "Selezionare l'applicazione da eseguire, e' possibile uscire`nin qualsiasi momento premendo la combinazione di tasti Ctrl + C. `nSelezionare la funzione da avviare:`n`n`t`t1 >> Generazione del file di nomi`n`t`t2 >> Rinominare in massa i file`n`t`t3 >> Spostare file`n`t`t4 >> Licenza di utilizzo`n`t`t5 >> Uscita`n`nDigitare un numero e premere invio"
        
        switch ($ans) {
            1 { $ans = ListingFiles; break }
            2 { $ans = ChangingFileNames; break }
            3 { $ans = MoveFile; break }
            4 { $ans = License; break }
            5 {break}
            Default {Write-Host "`nINSERIRE UN VALORE CORRETTO"}
        }

        Start-Sleep 1
        Separator
    } while ($ans -ne 4)
    
    Separator
    Write-Host "...USCITA..."
    Separator
    Start-Sleep 1
}

# Window adjustment
$H = Get-Host
$Win = $H.UI.RawUI
$Win.WindowTitle = 'Automazione su file'
SetConsoleSize -Height 40 -Width 69

Clear-Host
Read-Host "
+++++++++++++++++++++++++++++++++++++=:...:=++++++++++++++++++++++++
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
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
++++++++++++++++++++ PREMERE INVIO PER INIZIARE ++++++++++++++++++++
" | Out-Null

Main