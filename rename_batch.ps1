$folder = Read-Host "Inserisci indirizzo della cartella"
$csv_path = Read-Host "Inserisci l'indirizzo del file csv"
Write-Host("==========================================")

Write-Host "Ecco la prima riga, controlla il separatore:`n"
Get-Content -Path $csv_path -TotalCount 1
Write-Host("==========================================")

$sep = Read-Host "Inserisci il separatore corretto"
$csv = Import-Csv -Path $csv_path -Delimiter $sep
Write-Host("==========================================")

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
Write-Host("==========================================")

$count = Read-Host "Inserisci il numero corrispondete alla colonna dei nomi da sostituire`n>>"
$new_name_col = $csv[0].psobject.properties.name[$count -1]
Write-Host (("Colonna selezionata "), ("`"" + $new_name_col + "`"")) -Separator " >> "
Write-Host("==========================================")

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