$folder = Read-Host "Inserisci indirizzo della cartella"
$search_path = $folder + "\*"
$format = Read-Host "Inserisci il formato dei file"
$format = "*." + $format
$result = Get-ChildItem -Path $search_path -Name -Include $format

$result | Out-File ($folder + '\nomi_attuali.csv')