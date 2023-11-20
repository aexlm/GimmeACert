<#
    .SYNOPSIS
    Retourne le chemin d'un fichier.

    .DESCRIPTION
    Construit le chemin complet d'un fichier à partir du répertoire de travail et de son nom.
    Si le nom du fichier comporte des caractères '\', il est supposé qu'il s'agit déjà d'un chemin complet et il est renvoyé tel quel.

    .PARAMETER FileName
    Le nom du fichier.

    .PARAMETER WD
    Le chemin vers le répertoire de travail.
#>
function Get-FilePath {
    param(
        [String]
        $FileName,

        [String]
        $WD
    )

    if ($FileName -match '\\') {
        return $FileName        
    } else {
        return "$WD\$FileName"        
    }
}