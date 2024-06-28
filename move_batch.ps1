$search_folder = Read-Host "Inserisci l'indirizzo della cartella sorgente"
$destination_folder = Read-Host "Inserisci indirizzo della cartella di destinazione"
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
$count = Read-Host "Inserisci il numero corrispondete alla colonna dei file da spostare`n>>"
$file_col = $csv[0].psobject.properties.name[$count -1] # prelevo il valore dall'array
Write-Host ("Colonna selezionata >> " + "`"" + $file_col + "`"")
Write-Host("==========================================")

$ext = Read-Host "Inserisci il formato dei file (senza il punto)"

$n_file = 0
Get-ChildItem -Path $search_folder -File -Filter ("*." + $ext) | ForEach-Object {
    foreach($row in $csv){
        # String cleaning per un confronto accurato
        $file_name = $_.BaseName
        $file_to_move = $row.$file_col
        if ($file_to_move.Contains($ext)){
            $file_to_move = $row.$file_col.Split(".")[0]
        }

        if($file_name -eq $file_to_move){
            Move-Item $_.FullName $destination_folder
            $n_file++
        }
    }
}

Write-Host "Spostati " $n_file " file."
