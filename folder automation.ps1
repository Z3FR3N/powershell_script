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

Function RemoveInvalidChars {
    param(
      [Parameter(Mandatory=$true,
        Position=0,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
      [String]$Name
    )
  
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
    return ($Name -replace $re)
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
    Read-Host "Questo modulo genera un file csv contenente un elenco dei file contenuti in una cartella, è possibile indicare l'estensione del file o catalogare tutti i file contenuti all'interno di essa.`n`nPREMERE INVIO PER CONTINUARE" | Out-Null
    Separator

    Write-Host "SELEZIONARE LA CARTELLA CONTENENTE I FILES DA INDIVIDUARE.`n"

    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    $current_folder = Get-Folder $DesktopPath
    Write-Host $current_folder

    if ([string]::IsNullOrWhiteSpace($current_folder)) {
        Write-Host "NESSUN FILE SELEZIONATO, RITORNO AL MENU' PRINCIPALE."
        return 0
    }
    
    $keep_alive = $true

    $selected_folder = $current_folder.Split('\')[-1]
     
    do {
        Separator
        $ans = Read-Host "Cartella selezionata: `"$selected_folder`"`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio cartella`n`t`t3 >> Menu' principale`n`t`t4 >> Uscita`n`nDIGITARE UN NUMERO E PREMERE INVIO"

        switch ($ans) {
            1 { $keep_alive = $false; break }
            2 { $current_folder = Get-Folder $PSScriptRoot
                if ([string]::IsNullOrWhiteSpace($current_folder)) {
                    Write-Host "NESSUN FILE SELEZIONATO, RITORNO AL MENU' PRINCIPALE."
                    return 0
                }
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
    Read-Host "Questo modulo permette di rinominare in maniera massiva i file di una cartella partendo da un file csv dove, in due colonne distinte, vengono accoppiati i nomi attualmente assegnati ed i nomi che verranno rimpiazzati.`n`nE' FONDAMENTALE che la colonna dei nomi di origine contenga anche l'ESTENSIONE (.pdf,.csv,.png, etc...).`n`nES: nome documento.pdf    ->    nuovo nome documento.pdf`n`nPREMERE INVIO PER CONTINUARE" | Out-Null
    Separator

    $keep_alive = $true
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    Write-Host "SELEZIONARE il FILE CSV CONTENENTE I NOMI DA RINOMINARE.`n"
    $csv_path = Get-FileName "Seleziona un file csv" $DesktopPath "File (*.csv)|*.csv" $false
    Write-Host $csv_path
    
    if ([string]::IsNullOrWhiteSpace($csv_path)) {
        Write-Host "NESSUN FILE SELEZIONATO, RITORNO AL MENU' PRINCIPALE."
        return 0
    }

    $selected_csv = $csv_path.Split('\')[-1]
    Separator
    Write-Host "SELEZIONARE LA CARTELLA CONTENENTE I FILES.`n"
    $folder = Get-Folder $DesktopPath
    Write-Host $folder

    if ([string]::IsNullOrWhiteSpace($folder)) {
        Write-Host "NESSUNA CARTELLA SELEZIONATA, RITORNO AL MENU' PRINCIPALE."
        return 0
    }

    $selected_folder = $folder.Split('\')[-1]

    do {
        Separator
        $ans = Read-Host "Cartella selezionata: `"$selected_folder`"`n`nFile csv: `"$selected_csv`"`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio cartella`n`t`t3 >> Cambio File`n`t`t4 >> Menu' principale`n`t`t5 >> Uscita`n`nDIGITARE UN NUMERO E PREMERE INVIO"

        switch ($ans) {
            1 { $keep_alive = $false; break }
            2 { $folder = Get-Folder $DesktopPath
                if ([string]::IsNullOrWhiteSpace($folder)) {

                    Write-Host "NESSUN FILE SELEZIONATO, RITORNO AL MENU' PRINCIPALE."
                    return 0
                }
                $selected_folder = $folder.Split('\')[-1]; break }
            3 { [string]$csv_path = Get-FileName "Seleziona un file csv" $DesktopPath "File (*.csv)|*.csv" $false
            if ([string]::IsNullOrWhiteSpace($csv_path)) {
                Write-Host "NESSUN FILE SELEZIONATO, RITORNO AL MENU' PRINCIPALE."
                return 0
                }
                $selected_csv = $csv_path.Split('\')[-1]; break }
            4 { return 0 }
            5 { return 4 }
            Default {
                        Separator
                        Write-Host "INSERIRE UN VALORE CORRETTO."}
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
                [int]$col_num = Read-Host "`nInserire il numero della colonna dei nomi DI ORIGINE e premere INVIO`n"
                if ($col_num -ge 1 -and $col_num -le $count) {
                    $oldname_col = $csv[0].psobject.properties.name[$col_num -1] # prelevo il valore dall'array
                    Write-Host ("`n`t`tCOLONNA SELEZIONATA >> " + "`"" + $oldname_col + "`"")
                } else {
                    Separator
                    Write-Host "INSERIRE UN NUMERO CORRETTO."
                }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO."
            }

            try {
                [int]$col_num = Read-Host "`nInserire il numero della colonna dei nomi DA RINOMINARE e premere INVIO`n"
                    if ($col_num -ge 1 -and $col_num -le $count) {
                        $new_name_col = $csv[0].psobject.properties.name[$col_num -1]
                        Write-Host (("`n`t`tCOLONNA SELEZIONATA"), ("`"" + $new_name_col + "`"")) -Separator " >> "
                    } else {
                        Separator
                        Write-Host "INSERIRE UN NUMERO CORRETTO."
                    }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO."
            }

            Separator
            $ans = Read-Host "Ricapitolando: i nomi nella colonna `"$oldname_col`" verranno rimpiazzati con i nomi contenuti nella colonna `"$new_name_col`"`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio colonne`n`nDIGITARE UN NUMERO E PREMERE INVIO"
            
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
            $file_name = $_.Name
            
            foreach($row in $csv){
                # String cleaning per un confronto accurato
                $old_name = $row.$oldname_col

                if($old_name -eq $file_name){
                    # Inseriamo il test per il file
                    if ([string]::IsNullOrWhiteSpace($old_name)) {
                        Write-Host `"$old_name`" "-> NOME DEL FILE NON VALIDO."
                        break
                    }

                    $new_name = $row.$new_name_col.Trim()

                    if (($new_name + $_.Extension) -eq $old_name) {
                        Write-Host `"$new_name`" "-> NOME GIA' ASSEGNATO."
                        break
                    }

                    if([string]::IsNullOrWhiteSpace($new_name)) {
                        Write-Host `"$new_name`" "-> NOME DA ASSEGNARE VUOTO."
                        continue
                    }
                    
                    if ($new_name.Contains(".")) { 
                        $new_name = $new_name.Split(".")[0]
                    }

                    $new_name = RemoveInvalidChars $new_name
                    
                    

                    $copies = 0 # counter for already exsting copies
                    while (Test-Path -Path ($folder + "\" + $new_name + $_.Extension)) {
                        $copies++
                        $new_name = $new_name + "(" + $copies.ToString() + ")"
                    }

                    $complete_name = "{0}{1}" -f $new_name, $_.Extension
                    
                    Rename-Item -Path ($folder + "\" + $_.Name) -NewName $complete_name
                    $n_file++
                    Write-Host $_.Name "->" $complete_name
                    break
                }
            }
        }

        Separator

        Read-Host "RINOMINATI" $n_file "FILES.`nPREMERE INVIO PER CONTINUARE"|Out-Null
        
        return 0
    } else {
        Separator
        Write-Host "FILE CSV NON VALIDO."
        return 0
    }

}

function MoveFile {
    # Caricare un file CSV che contiene dei nomi di file da spostare ed indicare una cartella dove spostarli.
    Separator
    Read-Host "Questo modulo permette di spostare dei file in cartelle a partire da un file csv dove devono essere indicate, in due colonne separate, rispettivamente il nome del file e la cartella di arrivo. Non è richiesto di indicare l'estensione del file da spostare, e' pero' necessario che i file da spostare si trovino tutti all'interno della stessa `"cartella madre`".`n`nAd esempio:`n`nnome_documento.pdf    ->    cartella madre/NOMECARTELLA`nnome_documento2.pdf    ->    cartella madre/ALTRACARTELLA`n`nIn caso NOMECARTELLA e ALTRACARTELLA non esistessero, verranno create."  | Out-Null
    Separator

    $keep_alive = $true
    $DesktopPath = [Environment]::GetFolderPath("Desktop")

    Write-Host "SELEZIONARE il FILE CSV CONTENENTE I NOMI DA RINOMINARE.`n"
    $csv_path = Get-FileName "Seleziona un file csv" $DesktopPath "File (*.csv)|*.csv" $false
    Write-Host $csv_path
    if ([string]::IsNullOrWhiteSpace($csv_path)) {
        Write-Host "NESSUN FILE SELEZIONATO, RITORNO AL MENU' PRINCIPALE."
        return 0
    }
    $selected_csv = $csv_path.Split('\')[-1]
    Separator
    Write-Host "SELEZIONARE LA CARTELLA CONTENENTE I FILES.`n"
    $folder = Get-Folder $DesktopPath
    Write-Host $folder
    if ([string]::IsNullOrWhiteSpace($folder)) {
        Write-Host "NESSUNA CARTELLA SELEZIONATA, RITORNO AL MENU' PRINCIPALE."
        return 0
    }
    $selected_folder = $folder.Split('\')[-1]

    do {
        Separator
        $ans = Read-Host "Cartella selezionata: `"$selected_folder`"`n`nFile csv: `"$selected_csv`"`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio cartella`n`t`t3 >> Cambio File`n`t`t4 >> Menu' principale`n`t`t5 >> Uscita`n`nDIGITARE UN NUMERO E PREMERE INVIO"

        switch ($ans) {
            1 { $keep_alive = $false; break }
            2 { $folder = Get-Folder $DesktopPath
                $selected_folder = $folder.Split('\')[-1]; break }
            3 { [string]$csv_path = Get-FileName "Seleziona un file csv" $DesktopPath "File (*.csv)|*.csv" $false
                $selected_csv = $csv_path.Split('\')[-1]; break }
            4 { return 0 }
            5 { return 4 }
            
            Default {
                        Separator
                        Write-Host "INSERIRE UN VALORE CORRETTO."}
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
            try {
                [int]$col_num = Read-Host "`nInserire il numero corrispondente alla colonna dei FILES DA SPOSTARE e premere invio`n"
                if ($col_num -ge 1 -and $col_num -le $count) {
                    $files_col = $csv[0].psobject.properties.name[$col_num -1] 
                    Write-Host ("`n`t`tCOLONNA SELEZIONATA >> " + "`"" + $files_col + "`"")
                } else {
                    Separator
                    Write-Host "INSERIRE UN NUMERO CORRETTO"
                }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO"
            }

            try {
                [int]$col_num = Read-Host "`nInserire il numero della colonna contenente il nome della CARTELLA e premere invio`n"
                    if ($col_num -ge 1 -and $col_num -le $count) {
                        $folder_col = $csv[0].psobject.properties.name[$col_num -1]
                        Write-Host (("`n`t`tCOLONNA SELEZIONATA"), ("`"" + $folder_col + "`"")) -Separator " >> "
                    } else {
                        Separator
                        Write-Host "INSERIRE UN NUMERO CORRETTO."
                    }
            } catch {
                Separator
                Write-Host "INSERIRE UN VALORE CORRETTO."
            }

            Separator
            $ans = Read-Host "Ricapitolando: i file nella colonna `"$files_col`" verranno spostati nelle cartelle contenute in `"$folder_col`". Se la cartella non esiste, verra' creata.`n`nProcedere?`n`n`t`t1 >> Si`n`t`t2 >> Cambio colonne`n`nDIGITARE UN NUMERO E PREMERE INVIO"
            
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
        Separator

        [System.Collections.ArrayList]$csv = $csv
        Get-ChildItem -Path $folder -File | ForEach-Object {            
            $file_name = $_.BaseName
            
            # lt = less than
            foreach ($row in $csv){
                $file_to_move = RemoveInvalidChars $row.$files_col

                if ($file_to_move.Contains(".")) {
                    $file_to_move = $file_to_move.Split(".")[0]
                }
                
                if($file_name -eq $file_to_move){
                    if ([string]::IsNullOrWhiteSpace($row.$folder_col)) {
                        Write-Host "NESSUNA CARTELLA INDICATA PER" $file_to_move
                        break
                    }
                    $folder_to_move = RemoveInvalidChars $row.$folder_col
                    
                    
                    if (-not (Test-Path -Path ($folder + "\" + $folder_to_move))) {
                        New-Item -Path $folder -Name $folder_to_move -ItemType "directory"
                    }
                    
                    $fileExists = Test-Path -Path ($folder + "\" + $folder_to_move + "\" + $_.Name)
                    $newName = $_.Name
                    $count = 0

                    while ($fileExists) {
                        $count++
                        $newName = $_.BaseName + "(" + $count + ")" + $_.Extension
                        $fileExists = Test-Path -Path ($folder + "\" + $folder_to_move + "\" + $newName)
                    }
                    
                    Move-Item -Path $_.FullName -Destination ($folder + "\" + $folder_to_move + "\" + $newName)
                    Write-Host $newName "->" $folder"\"$folder_to_move
                    $n_file++
                    $csv.Remove($row)
                    break
                }
            }
        }

        Separator
        Read-Host "SPOSTATI" $n_file "FILE.`nPREMERE INVIO PER CONTINUARE" | Out-Null
    } else {
        Separator
        Write-Host "FILE NON VALIDO."
        return 0
    }
}

function License {
    Read-Host "`nGNU GENERAL PUBLIC LICENSE
Version 3, 29 June 2007
    
Copyright (c) 2007 Free Software Foundation, Inc. <https://fsf.org/>
    
This program is used by the Marche's chamber of commerce as a wrapper to some powershell automations.
Copyright (C) 2024 - Martinangeli Luca

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.
If not, see <https://www.gnu.org/licenses/>.

Contact: luke.d3v@gmail.com`n`n`PREMERE INVIO PER CONTINUARE" | Out-Null

    return 0
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
        $ans = Read-Host "Questo programma e' diviso in `"moduli`" che permetto di eseguire in maniera massiva delle azioni, e' possibile uscire in qualsiasi momento premendo la combinazione di tasti Ctrl + C o selezionando l'opzione `"Uscita`" quando richiesto. `nSelezionare il modulo da avviare:`n`n`t`t1 >> Generazione del file di nomi`n`t`t2 >> Rinominare in massa i file`n`t`t3 >> Organizzare file in cartelle`n`t`t4 >> Licenza di utilizzo`n`t`t5 >> Informazioni aggiuntive`n`t`t6 >> Uscita`n`nDigitare un numero e premere invio"
        
        switch ($ans) {
            1 { $ans = ListingFiles; break }
            2 { $ans = ChangingFileNames; break }
            3 { $ans = MoveFile; break }
            4 { $ans = License; break }
            5 { $ans = Info; break }
            6 {break}
            Default {Write-Host "`nINSERIRE UN VALORE CORRETTO."}
        }

        Start-Sleep 1
        Separator
    } while ($ans -ne 6)
    
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

Write-Host "
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
"
Start-Sleep 1

Write-Host "Automazione file  Copyright (C) 2024  Martinangeli Luca
This program comes with ABSOLUTELY NO WARRANTY; for details type 4.
This is free software, and you are welcome to redistribute it
under certain conditions expressed in the license."

Start-Sleep 2.5
Main