<#
    .SYNOPSIS
    Déplace les fichiers temporaires créés lors de l'exécution du script.

    .DESCRIPTION
    Cette fonction créé un dossier nommé '_old{Timestamp}' dans le répertoire WorkingDirectory.
    Les fichiers de configuration .inf, .req, .cer et .rsp sont déplacés vers ce dossier.
    C'est ensuite à l'utilisateur de décider s'il souhaite conserver ou supprimer ces fichiers.

    .PARAMETER Directory
    L'emplacement du réperoire de travail.

    .PARAMETER INFFilePath
    L'emplacement du fichier de configuration pour créer la demande de certificat.

    .PARAMETER CSRFilePath
    L'emplacement du fichier de requête du certificat.

    .PARAMETER CERFilePath
    L'emplacement du fichier .cer contenant le certificat.
#>
function Clear-WorkingDirectory {
    param(
        [String]
        $Directory,

        [String]
        $INFFilePath,

        [String]
        $CSRFilePath,

        [String]
        $CERFilePath
    )

    $Timestamp = Get-Date -Format "yyyymmddHHMMss"
    $OldDir = "$Directory\_old$Timestamp"    

    try {
        New-Item -ItemType "directory" -Path $OldDir -ErrorAction Stop > $null 
        Get-ChildItem -Path @($INFFilePath, $CSRFilePath, ($CERFilePath -replace ('\.cer$', '.rsp'))) | Move-Item -Destination $OldDir -ErrorAction Stop
        Write-Host "Déplacement des fichiers temporaires dans le dossier $OldDir`n"
    } catch {
        Write-Host -ForegroundColor Red "Erreur lors du déplacement des fichiers temporaires"
        Write-Host -ForegroundColor Red "$_`n"
    }
}