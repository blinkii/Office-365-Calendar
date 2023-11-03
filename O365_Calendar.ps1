$continue = $true
while ($continue){
  write-host “"
  write-host “-------------------- O365 Modif Calendrier ----------------------”
  write-host “1/ Connexion MS O365”
  write-host “2/ Voir droit utilisateur”
  write-host "3/ Ajouter droit utiliseur"
  write-host "4/ Supprimer droit utiliseur"
  write-host "X/ Quitter"
  write-host "--------------------------------------------------------"
  write-host “"
  $choix = read-host “Faites un choix ”
  switch ($choix){
      1{
            Connect-ExchangeOnline -Credential $M365credentials
        }
    2{
        $mail = Read-Host "Entrez le mail complet du collaborateur"
        Get-MailboxFolderPermission -Identity ${mail}:\Calendrier
        }
    3{
        $mail = Read-Host "Entrez le mail complet du collaborateur"
        $mail2 = Read-Host "Entrez le mail complet de la personne a ajouter en admin"
        Add-MailboxFolderPermission -Identity ${mail}:\Calendrier -User ${mail2} -AccessRights editor
        }
    4{
        $mail = Read-Host "Entrez le mail complet du collaborateur à modifier"
        $mail2 = Read-Host "Entrez le mail complet de la personne à retrier"
        Remove-MailboxFolderPermission -Identity ${mail}:\Calendrier -User ${mail2}
        }
    ‘x’ {$continue = $false}
    default {Write-Host "Choix invalide"-ForegroundColor Red}
  }
}
