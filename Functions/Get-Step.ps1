<#
    .SYNOPSIS
    Détermine la prochaine étape à effectuer.

    .DESCRIPTION
    D'après les paramètre passés à la fonction, retourne le numéro correspondant.

    .PARAMETER Thumbprint
    L'empreinte du certificat.

    .PARAMETER CERFile
    Le contenu du fichier .cer (peut être vide).

    .PARAMETER RequestID
    L'identifiant de la demande de certificat.

    .PARAMETER CSRFile
    Le contenu du fichier .csr (peut être vide).

    .PARAMETER INFFile
    Le contenu du fichier .inf (peut être vide).
#>
function Get-Step {
    param (
        [String]
        $Thumbprint,

        [String]
        $CERFile,

        [Int]
        $RequestID,

        [String]
        $CSRFile,

        [String]
        $INFFile
    )

    if ($Thumbprint) {
        return 4
    } elseif ($CERFile) {
        return 3
    } elseif ($CSRFile -or $RequestID) {
        return 2
    } elseif ($INFFile) {
        return 1
    } else {
        return 0
    }
}