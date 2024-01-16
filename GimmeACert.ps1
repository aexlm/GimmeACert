<#PSScriptInfo

    .VERSION 3.0.0

    .AUTHOR axel.mauchand@metsys.fr

    .COMPANYNAME METSYS
#>


<#
    .SYNOPSIS
    Automatise la demande de certificat et son export.

    .DESCRIPTION
    Ce script a pour bur d'automatiser les demandes de certificats pour les serveurs Web.
    Le processus en place se repose sur un service obsolète de Microsoft dont il faut se séparer.
    Ce script permettra également de spécifier des noms de sujets alternatifs sans avoir à activer une clé de registre dangereuse sur l'autorité de distribution des certificats.

    Fonctionnement du script :
    - Deux paramètres sont obligatoires : le nom du Sujet et le nom du Template souhaité
    - Les autres paramètres utilisent des valeurs par défaut
    - Il est possible de personnaliser le répertoire dans lequel seront placés les fichiers
    - Si les mêmes paramètres sont passés au lancement répété du script, le script reprendra à la dernière étape réussie

    Déroulement du script :
    - D'après les paramètres passés, vérification des fichier déjà créés afin de déterminer à quelle étape le script doit démarrer
    - En partant de rien, création du fichier de configuration contenant le nom du sujet et ses noms alternatifs ainsi que la taille de la clé
    - Utilisation de ce fichier pour créer le fichier de requête du certificat (CSR)
    - Soumission de la requête à l'autorité de certification
    - Si le template n'a pas d'approbation, le certificat est directement enregistré
    - Sinon, l'ID de la requête est affiché et enregistré tant que le certificat n'a pas été délivré
    - Lorsque la tentative de récupérer le certificat réussie, le certificat est enregistré
    - Le certificat est ensuite installé dans le magasin spécifié et est lié à la clé privée
    - Des fonctions d'export en PEM et PFX sont proposées
    - Il est ensuite possible de supprimer la clé privée, déplacer les fichiers temporaires (.cer, .inf, .req, .rsp) et supprimer le certificat du magasin

    .PARAMETER UseCSR
    Ce paramètre spécifie si un fichier CSR doit être passer en entrée.
    Son utilisation ouvre une boîte de dialogue permettant de sélectionner le CSR voulu.
    Certaines fonctionnalités sont automatiquement passées :
    - Installation du certificat dans le magasin utilisateur
    - Export de la clé privée / PFX / PEM
    Si le paramètre est spécifié, les informations concernant la requête seront également affichées.

    .PARAMETER UsePublicDeposit
    Ce paramètre permet de copier le certificat récupéré directement dans le répertoire partagé avec les utilisateurs.
    Le chemin vers le répertoire est défini avec le paramètre -PublicDeposit.

    .PARAMETER PublicDepositPath
    Spécifie le chemin d'accès au partage réseau contenant les certificats distribués aux utilisateurs.
    Sa valeur par défaut est <PathTo\PublicDeposit>.
    
    .PARAMETER ObjectName
    Ce paramètre défini le sujet pour lequel ce certificat sera généré.
    Si le paramètre n'est pas spécifié au moment de construire le fichier de configuration, il est demandé en input.

    .PARAMETER CertificateTemplate
    Ce paramètre défini quel template choisir pour le certificat demandé.
    Si le paramètre n'est pas spécifié au moment de soumettre la demande à l'autorité de certification, l'utilisateur doit choisir parmi les modèles disponibles.

    .PARAMETER San
    Ce paramètre permet d'indiquer les Subject Alternate Names (SAN) à indiquer pour le certificat.
    Son format est : dns=www.example.com&ipaddress=0.0.0.0
    Plus d'informations sur cette construction : https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/ff625722(v=ws.10)?redirectedfrom=MSDN#to-create-a-requestpolicyinf-file

    Si aucune valeur n'est spécifiée, une résolution DNS est faite pour essayer de construire le SAN.
    Si la tentative échoue, l'utilisateur est interrogé pour renseigner les SAN manuellement.

    .PARAMETER ExportableKey
    Indique si la clé privée du certificat peut être exportée.
    Par défaut, le paramètre prend la valeur `$True`, sauf si le paramètre -InstallMachine est spécifié - il faut alors indiquer ce paramètre au lancement du script.
    Si l'on veut forcer le paramètre à $False, il faut indiquer `-ExportableKey:$False`.    

    .PARAMETER KeyLength
    Spécifie la taille de la clé privée et de la clé publique pour le certificat demandé.
    La taille spécifiée doit être de 1024, 2048 ou 4096.
    Si la taille indiquée ne correspond à aucune des valeurs possibles, la variable prend la valeur par défaut.
    La valeur par défaut est 2048.
    Selon le template choisi, la taille indiquée ne doit pas être inférieure à la valeur requise par le modèle de certificat.

    .PARAMETER OrganizationalUnit
    Spécifie l'attribut OU dans le sujet du certificat.

    .PARAMETER Email
    Spécifie l'attibut Email dans le sujet du certificat.
    
    .PARAMETER Organisation
    Spécifie l'attribut Organisation dans le sujet du certificat.

    .PARAMETER Localisation
    Spécifie l'attribut Localisation dans le sujet du certificat.

    .PARAMETER Region
    Spécifie l'attribut Région dans le sujet du certificat.

    .PARAMETER Pays
    Spécifie l'attribut Pays dans le sujet du certificat.
    
    .PARAMETER Thumbprint
    Permet de spécifier directement l'empreinte du certificat à extraire.
    Il faut également spécifier le magasin dans lequel le certificat est installé.
    En utilisant ce paramètre, il faut d'abord s'assurer que la clé privée soit bien associée au certificat.
    Si ce n'est pas le cas, il faut utiliser certutil avec le paramètre -repairstore.

    .PARAMETER RequestID
    Spécifie l'identifiant du certificat à récupérer depuis l'autorité de certification.
    Si le certificat a été délivré, il est enregistré.
    Si le certificat est toujours en attente d'être délivré, l'utilisateur peut réessayer.
    Si la demande de certificat a été rejetée par un adminstrateur, l'utilisateur est prévenu et le programme se termine.
    Si l'identifiant ne correspond à aucun certificat, il s'agit d'une erreur et le programme se termine.

    .PARAMETER NoCertInstall
    Spécifie si le certificat doit être installé dans le magasin de l'utilisateur après avoir été récupéré.
    Si le paramètre UseCSR est utilisé, ce paramètre passe à True.

    .PARAMETER InstallMachine
    Si ce paramètre est spécifié, l'installation du certificat se fera dans le magasin Personnel de l'ordinateur.
    Par défaut, l'export du certificat aux formats PEM et PFX n'est pas demandé (sauf si spécifié lors de l'appel du script).
    Il en va de même pour les actions concernant la suppression de la clé privée et du certificat.
    Ce paramètre n'est appliqué que si le scipt est exécuté avec les droits d'administrateurs.

    .PARAMETER ExportPrivateKey
    Si ce paramètre est spécifié, l'utilisateur ne sera pas sollicité lors de l'export de la clé privée du certificat.

    .PARAMETER ExportPEM
    Si ce paramètre est spécifié, l'utilisateur ne sera pas sollicité et la chaîne de certificat sera directement exporté au format .PEM.

    .PARAMETER ExportPFX
    Si ce paramètre est spécifié, l'utilisateur ne sera pas sollicité et la chaîne de certificat sera directement exporté au format .PFX.

    .PARAMETER PfxPassword
    Spécifie le mot de passe pour le fichier .PFX.
    Attention, le type attendu est SecureString, il faut donc initialiser le mot de passe dans une variable de ce type.
    Si le paramètre est spécifié et qu'une chaîne de caractère simple est spécifiée, le programme ne pourra pas se lancer.
    Si aucun paramètre n'est spécifié, l'utiliateur aura l'option de donner un mot de passe au moment de l'export en PFX.

    .PARAMETER ClearWorkingDirectory
    Si ce paramètre est spécifié, l'utilisateur ne sera pas sollicité et les fichiers temporaires seront déplacés dans un nouveau dossier, en attente de suppression.
    Un dossier _old{Timestamp} est créé dans le répertoire spécifié par le paramètre -WorkingDirectory puis les fichiers y sont déplacés.
    Les fichiers à déplacer sont ceux comportant les extensions suivantes [.inf, .req, .cer, .rsp]

    .PARAMETER DeleteWorkingDirectory
    Si ce paramètre est spécifié, le répertoire de travail sera supprimé une fois l'exécution du script terminé.
    Ce paramètre est activé par défaut si les paramètres -UseCSR et -UsePublicDeposit sont utilisés.

    .PARAMETER DeleteCertFromStore
    Si ce paramètre est spécifié, l'utilisateur ne sera pas sollicité et le certificat installé dans le magasin sera automatiquement supprimé à la fin de l'exécution du programme.

    .PARAMETER DeletePrivateKey
    Si ce paramètre est spécifié, l'utilisateur ne sera pas sollicité et la commande à exécuter pour supprimer la clé privée associée au certificat installé sera affichée.
    Il y a de toute manière une demande de validation pour l'utilisateur avant que la commande ne soit effectuée.

    .PARAMETER ExportResults
    Si ce paramètre est spécifié, l'utilisateur ne sera pas sollicité et les informations concernant le certificat seront exportés dans un fichier CSV.
    L'emplacement du fichier CSV peut être défini par le paramètre -ResultsFilePath.
    Si aucun paramètre n'est fourni, un emplacement par défaut est suggéré et s'il ne convient pas, une boîte de dialogue s'ouvre afin de chosir l'emplacement du fichier.
    Une fois les résultats enregistrés, les informations sont affichés à l'écran.

    .PARAMETER ResultsFilePath
    Indique l'emplacement du fichier .csv contenant les informations sur le certificat demandé.
    Aucune valeur par défaut n'est attribuée au lancement du programme.

    .PARAMETER TargetCA
    Permet d'indiquer l'autorité de certification à interroger lors de la soumission de la requête ainsi qu'au moment de récupérer le certificat soumis.
    Syntaxe : <FQDN du serveur>\<Nom de l'autorité>
    Exemple : CA-Server.domain.com\Distribution-CA
    Si le paramètre n'est pas fourni lors de l'envoi de la demande ou la tentative de récupération du certificat, une action utilisateur est requise.
    Un menu appraît et affiche la liste des CA disponibles.
    Si une seule CA est disponible, le menu n'apparaît pas et le choix s'effectue automatiquement.

    .PARAMETER WorkingDirectory
    Indique le chemin par défaut dans lequel seront placés les fichiers créés lors de l'exécution de ce programme.
    Si le chemin indiqué n'existe pas, il sera créé si l'utilisateur possède les droits nécessaires.

    .PARAMETER PolicyFileName
    Indique le nom du fichiers à donner au fichier de configuration pour créer la requête.
    Si le chemin complet est spécifié, il faut que le repértoire soit créé et que l'utilisateur y ait les droits d'écritures.
    Autrement, le fichier est créé dans le répertoire spécifié dans le paramètre -WorkingDirectory.
    Par défaut, le nom donné à ce fichier est <ObjectName>Policy.inf

    .PARAMETER CSRFileName
    Indique le nom du fichiers à donner au fichier contenant la demande de certificat à envoyer à l'autorité.
    Si le chemin complet est spécifié, il faut que le repértoire soit créé et que l'utilisateur y ait les droits d'écritures.
    Autrement, le fichier est créé dans le répertoire spécifié dans le paramètre -WorkingDirectory.
    Par défaut, le nom donné à ce fichier est <ObjectName>Request.req

    .PARAMETER CERFileName
    Indique le nom du fichiers à donner au fichier contenant le certificat demandé.
    Si le chemin complet est spécifié, il faut que le repértoire soit créé et que l'utilisateur y ait les droits d'écritures.
    Autrement, le fichier est créé dans le répertoire spécifié dans le paramètre -WorkingDirectory.
    Par défaut, le nom donné à ce fichier est <ObjectName>Certificate.cer

    .PARAMETER PEMFileName
    Indique le nom du fichiers à donner à l'export du certificat au format .pem.
    Si le chemin complet est spécifié, il faut que le repértoire soit créé et que l'utilisateur y ait les droits d'écritures.
    Autrement, le fichier est créé dans le répertoire spécifié dans le paramètre -WorkingDirectory.
    Par défaut, le nom donné à ce fichier est <ObjectName>Certificate.pem

    .PARAMETER PFXFileName
    Indique le nom du fichiers à donner à l'export du certificat au format .pfx.
    Si le chemin complet est spécifié, il faut que le repértoire soit créé et que l'utilisateur y ait les droits d'écritures.
    Autrement, le fichier est créé dans le répertoire spécifié dans le paramètre -WorkingDirectory.
    Par défaut, le nom donné à ce fichier est <ObjectName>Certificate.pfx

    .PARAMETER CertStore
    Ce paramètre indique le magasin dans lequel installé le certificat demandé afin de le lier à sa clé privée.
    Par défaut, le certificat est installé dans le magasin personnel de l'utilisateur courrant.
    Ce chemin est indiqué ainsi : 'Cert:\CurrentUser\My'
#>

<# TODO : 
    - Régler le soucis de redemander la Target CA quand une demande est en attente
    - Ménage dans le WD : try catch si le fichier n'existe pas       
    - Créer un second script pour le renouvellement des certificats
        - En entrée : le couple .cer + .key OU .pfx OU .csr/.req (mais à éviter car SAN pas bien)
        - En sortie : juste le .cer car la clé privée ne doit pas être exportée
    - Faire du WorkingDirectory une variable globale
        - Pour le module, la valeur par défaut sera '.\'    
#>

param (    
    [Switch]
    $UseCSR,

    [Switch]
    $UsePublicDeposit,

    [String]
    $PublicDepositPath = "",

    [String]
    $ObjectName,
    
    [String]
    $CertificateTemplate,

    [String]
    $San,

    [Switch]
    $ExportableKey,

    [Int]
    $KeyLength = 2048,

    [String]
    $OrganizationalUnit,
    
    [String]
    $Email,

    [String]
    $Organisation,

    [String]
    $Localisation,

    [String]
    $Region,

    [String]
    $Pays,

    [String]
    $Thumbprint = $null,

    [Int]
    $RequestID,

    [Switch]
    $NoCertInstall,

    [Switch]
    $InstallMachine,

    [Switch]
    $ExportPrivateKey,

    [Switch]
    $ExportPEM,

    [Switch]
    $ExportPFX,

    [SecureString]
    $PfxPassword,
    
    [Switch]
    $ClearWorkingDirectory,

    [Boolean]
    $DeleteWorkingDirectory,

    [Switch]
    $DeleteCertFromStore,

    [Switch]
    $DeletePrivateKey,

    [Switch]
    $ExportResults,

    [String]
    $ResultsFilePath,

    [String]
    $TargetCA,

    [String]
    $WorkingDirectory,

    [String]
    $PolicyFileName,

    [String]
    $CSRFileName,

    [String]
    $CERFileName,

    [String]
    $PrivateKeyFileName,

    [String]
    $PEMFileName,

    [String]
    $PFXFileName,

    [String]
    $CertStore = "Cert:\CurrentUser\My"
)

if (-not (Get-Module -Name "PKIMAN")) {
    try {
        Import-Module "PKIMAN" -ErrorAction Stop
    } catch [System.IO.IOException] {
        Write-Host -ForegroundColor Red "Module PKIMAN non installé ou introuvable.`nFermeture du programme..."
        return
    }  catch {
        Write-Host -ForegroundColor Red "Erreur du chargement du module PKIMAN.`n$_`nFermeture du programme..."
        return
    }    
    Write-Host "Module PKIMAN importé."
} else {
    Write-Host "Module PKIMAN déjà importé."
}

try {
    @("Clear-WorkingDirectory.ps1", "Export-Results.ps1","Get-FilePath.ps1","Get-Step.ps1") | ForEach-Object { . "$PSScriptRoot`\Functions\$_" }
} catch {
    Write-Host -ForegroundColor Red "Fonctions non trouvées.`n$_`nFermeture du programme..."
    return
}

if ($UseCSR) {
    $ObjectName, $CSRFileName = Select-Csr -Path $CSRFileName -InitialDirectory $PublicDepositPath
    if (-not $ObjectName) { return }
    if (-not $PublicDepositPath) {
        $PublicDepositPath = Split-Path -Parent -Path $CSRFileName
    }

    $Validation = Read-Host "`nValidation de la requête ? (Y/N)"

    if ($Validation.ToLower() -eq 'n') {
        $Validation = Read-Host "`nRecréer la demande depuis le nom de l'objet ? (Y/N)"
        if ($Validation.ToLower() -eq 'y') {
            $CSRFileName = "$ObjectName`Request.req"
        } else {
            Write-Host "Fermeture du programme..."
            return
        }
    } else {
        $NoCertInstall, $UsePublicDeposit = $True, $True
    }
    
}

#Gestion des cas en reprenant le script à partir du Thumbprint ou sans nom de sujet
try {
    if (-not $ObjectName) {
        if ($Thumbprint) {
            try {
                $Certificate = Get-ChildItem "$Store\$CertificateThumbprint" -ErrorAction Stop
            } catch [System.Management.Automation.ItemNotFoundException] {
                Write-Host -ForegroundColor Yellow "Le chemin spécifié pour le certificat n'existe pas."
                return
            }

            $ObjectName = (((($Certificate.Subject) -split "\s") -match "CN=") -split '=')[-1] -replace ",$"
        }                              
    }    

    if (-not $WorkingDirectory) {
        $WorkingDirectory = "C:\Temp\Certificats\$ObjectName"
    }    
        
    #Si le répertoire de travail finit toujours par le caractère '\', on le supprime
    if ($WorkingDirectory[-1] -eq '\') {
        $WorkingDirectory = $WorkingDirectory -replace ".$"
    }  

    New-Item -ItemType "directory" -Path $WorkingDirectory -ErrorAction Stop > $null
    Write-Host "Dossier $WorkingDirectory créé."
} catch [System.UnauthorizedAccessException] {
    Write-Host -ForegroundColor Red "L'utilisateur en cours n'a pas les droits nécessaires pour créer le dossier à l'emplacement suivant : $WorkingDirectory"
    exit
} catch [System.IO.IOException] {
    Write-Host "Utilisation du répertoire $WorkingDirectory."
} catch {
    Write-Host -ForegroundColor Red $_
    exit
}

#Construction du nom des fichiers s'ils n'ont pas été donné
if (-not $CERFileName) {
    $CERFileName = "$ObjectName`Certificate.cer"
}
if (-not $CSRFileName) {
    $CSRFileName = "$ObjectName`Request.req"
}
if (-not $PolicyFileName) {
    $PolicyFileName = "$ObjectName`Policy.inf"
}
if (-not $PrivateKeyFileName) {
    $PrivateKeyFileName = "$ObjectName`PKey.key"
}
if (-not $PEMFileName) {
    $PEMFileName = "$ObjectName`Certificate.pem"
}
if (-not $PFXFileName) {
    $PFXFileName = "$ObjectName`Certificate.pfx"
}

#Copie en local du CSR dans le répertoire de travail
if ($UseCSR) {
    Copy-Item -Path $CSRFileName -Destination $WorkingDirectory -ErrorAction SilentlyContinue
}

#Construction des chemins des fichiers à l'aide de la fonction Get-FilePath
$CERFilePath = Get-FilePath -WD $WorkingDirectory -FileName $CERFileName
$CSRFilePath = Get-FilePath -WD $WorkingDirectory -FileName $CSRFileName
$PolicyFilePath = Get-FilePath -WD $WorkingDirectory -FileName $PolicyFileName
$PrivateKeyFilePath = Get-FilePath -WD $WorkingDirectory -FileName $PrivateKeyFileName
$PEMFilePath = Get-FilePath -WD $WorkingDirectory -FileName $PEMFileName
$PFXFilePath = Get-FilePath -WD $WorkingDirectory -FileName $PFXFileName

#S'il a été spécifié que l'installation devait se faire avec le contexte machine, le magasin d'installation du certificat est mis à jour si nécessaire.
if ($InstallMachine) {
    $CertStore = $CertStore.Replace('CurrentUser', 'LocalMachine')    
}

#Vérification de l'existance du magasin de certificat spécifié.
#Si ce n'est pas le cas, une valeur par défaut est choisie à la place.
try {
    Get-ChildItem -Path $CertStore -ErrorAction Stop > $null
} catch [System.Management.Automation.ItemNotFoundException] {
    Write-Host -ForegroundColor Yellow "Le magasin spécifié $CertStore n'existe pas."
    if ($InstallMachine) {
        $CertStore = "Cert:\LocalMachine\My"
    } else {
        $CertStore = "Cert:\CurrentUser\My"
    }
}
Write-Host -ForegroundColor Green "Magasin d'installation du certificat : $CertStore`n"

#D'après les paramètres passés en entrée du script, détermine à quelle étape il faut reprendre.
while ($Step -ne 4) {
    $Step = Get-Step -Thumbprint $Thumbprint `
        -CERFile (Get-Content -Path $CERFilePath -ErrorAction Ignore) `
        -RequestID $RequestID `
        -CSRFile (Get-Content -Path $CSRFilePath -ErrorAction Ignore) `
        -INFFile (Get-Content -Path $PolicyFilePath -ErrorAction Ignore)

    switch ($Step) {
        0 {
            $Sujet = Get-Subject -ObjectName $ObjectName -Organisation $Organisation -Localisation $Localisation -Region $Region -Pays $Pays
            New-Policy -FilePath $PolicyFilePath -San $San -Subject $Sujet -KeyLength $KeyLength -ExportableKey $ExportableKey -InstallMachine $InstallMachine
        }
        1 {
            New-CSR -PolicyFile $PolicyFilePath -CSRFile $CSRFilePath -UseMachine $InstallMachine
        }
        2 {
            $TargetCA = New-CER -CSRFile $CSRFilePath -TargetCA $TargetCA -CertificateTemplate $CertificateTemplate -CERFile $CERFilePath -RequestID $RequestID -UseMachine $InstallMachine
            if (-not $TargetCA) { return }
        }
        3 {
            if (-not $NoCertInstall) {
                $Thumbprint = Install-Cert -CERFile $CERFilePath -Store $CertStore
            } else {
                $RetrievedCertificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $CERFilePath
                $Thumbprint = $RetrievedCertificate.Thumbprint
            }
        }
    }
}

if ($UsePublicDeposit -and ($PublicDepositPath -ne $WorkingDirectory)) {
    $CopyDeposit = Read-Host "Copier le certificat dans le répertoire contenant le CSR ? (Y/N)"
    if ($CopyDeposit.ToLower() -eq 'y') {
        Copy-Item -Path $CERFilePath -Destination "$PublicDepositPath\$CERFileName"
        Write-Host "Certificat disponible à l'emplacement $PublicDepositPath\$CERFileName"
    }    
}

if (-not $NoCertInstall) {
    if (-not $ExportPrivateKey -and -not $InstallMachine) {
        $ExportPrivateKey = ((Read-Host "Exporter la clé privée ? (Y/N)").ToLower() -eq 'y')
    } 
    if ($ExportPrivateKey) {
        Export-PrivateKey -CertificateThumbprint $Thumbprint -Store $CertStore -PrivateKeyFilePath $PrivateKeyFilePath
    }
    
    if (-not $ExportPEM -and -not $InstallMachine) {
        $ExportPEM = ((Read-Host "Exporter en PEM ? (Y/N)").ToLower() -eq 'y')
    }
    if ($ExportPEM) {
        Export-Pem -CertificateThumbprint $Thumbprint -PEMPathFile $PEMFilePath -Store $CertStore
    }
    
    if (-not $ExportPFX -and -not $InstallMachine) {
        $ExportPFX = ((Read-Host "Exporter en PFX ? (Y/N)").ToLower() -eq 'y')
    }
    if ($ExportPFX) {
        Export-Pfx -CertificateThumbprint $Thumbprint -Store $CertStore -PFXFilePath $PFXFilePath -Password $PfxPassword
    }
    
    if ($ClearWorkingDirectory -or ((Read-Host "Faire le ménage dans le répertoire de travail ? (Y/N)").ToLower() -eq 'y')) {
        Clear-WorkingDirectory -Directory $WorkingDirectory -INFFilePath $PolicyFilePath -CSRFilePath $CSRFilePath -CERFilePath $CERFilePath
    }
    
    if (-not $DeletePrivateKey -and -not $InstallMachine) {
        $DeletePrivateKey = ((Read-Host "Supprimer la clé privée ? (Y/N)").ToLower() -eq 'y')
    }
    if ($DeletePrivateKey) {
        $DeletePrivateKey = Remove-PrivateKey -Thumbprint $Thumbprint -Store $CertStore
    }
    
    if ($ExportResults -or ((Read-Host "Exporter les résultats ? (Y/N)").ToLower() -eq 'y')) {
        Export-Results -Thumbprint $Thumbprint `
            -CSVFilePath $ResultsFilePath `
            -PFXFilePath $PFXFilePath `
            -PEMFilePath $PEMFilePath `
            -Store $CertStore `
            -DeletedPrivateKey $DeletePrivateKey `
            -WorkingDirectory $WorkingDirectory
    }
    
    if (-not $DeleteCertFromStore -and -not $InstallMachine) {
        $DeleteCertFromStore = ((Read-Host "Supprimer le certificat du magasin ? (Y/N)").ToLower() -eq 'y')
    }
    if ($DeleteCertFromStore) {
        Remove-CertFromStore -Thumbprint $Thumbprint -Store $CertStore
    }    
}

if ($DeleteWorkingDirectory) {
    try {
        Write-Host -ForegroundColor Yellow "Suppression du répertoire $WorkingDirectory"
        Remove-Item -Recurse -Path $WorkingDirectory -ErrorAction Stop
    } catch {
        Write-Host -ForegroundColor Red "Erreur lors de la suppression du répertoire.`n$_"        
    }
}

return