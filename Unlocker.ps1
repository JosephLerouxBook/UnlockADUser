##############################################
##                                          ##
##  Nom du script : AD_Unlocker.ps1         ##
##  Version : 1.2                           ##
##  Auteur : Joseph Leroux                  ##
##  tmps d'exec :                           ##
##                                          ##
##  Utilisation : Interface proposant une   ##
##    liste dans laquelle on choisit le nom ##
##     d'une personne dont on veux          ##  
##     deverouiller le compte.              ##
##############################################

##################################################################################################
## Note de MAJ :  1.1 Ne refait pas le Microsoft ID avec une fonction. Recupere avec une requete##                                                                              ##
##                    Gere (ducoup) les noms composer                                           ##
##                1.2 Cache la console pendant que le programme tourne                          ##
##################################################################################################

##################################################################################################
## Pourrait etre fait : Réflechir a un lien ouvrant une GUI de reinitialisation de mdp          ##
##                                                                                              ##
##################################################################################################
PowerShell.exe -windowstyle hidden {
    #Import des biblioteques : WindowsForm, ActiveDirectory
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    import-module ActiveDirectory

    #Création du FORM
    $form = New-Object Windows.Forms.Form
    $form.MinimizeBox = $False
    $form.MaximizeBox = $False
    $form.Text = "Déverouiller le compte"
    $form.Size = New-Object System.Drawing.Size(420,125)

    #Creation liste deroulante
    $liste1 = New-Object System.Windows.Forms.Combobox
    $liste1.Location = New-Object Drawing.Point 20,20
    $liste1.Size = New-Object System.Drawing.Size(300,10)
    $liste1.DropDownStyle = "DropDownList"
    #Avec les propositions venant de l'AD
    $user_list = Get-ADUser  -filter * -Properties Name | Sort-Object
    foreach ($elem in $user_list) {
        $liste1.items.Add($elem.name)
    }
    #Affiche une personne dans la liste quand la fenettre s'ouvre
    $liste1.SelectedIndex = 0

    #Creation button 'déverouiller'
    $button1 = New-Object Windows.Forms.Button
    $button1.Text = "Déverouiller"
    $button1.Location = New-Object System.Drawing.Size(325,19)
    $button1.Size = New-Object System.Drawing.Size(60,23)

    #Creation du text (label) en dessous de la liste
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(30,50)
    $label1.Size = New-Object System.Drawing.Size(350,100)
    $label1.Text = "Choisissez une personne, puis, Appuyez sur déverouiller"

    #A l'activation du boutton, déverouille le compte de la personne ciblé dans la liste
    $button1.Add_click({
        $user_realname = $liste1.items.Item($liste1.SelectedIndex)
        $user = Get-ADUser -filter 'name -like $user_realname'
        Unlock-ADAccount -identity $user
        if ($?) {$label1.Text = "COMPTE DEVEROUILLER"} 
    })

    #Affichage
    $form.Controls.Add($liste1)
    $form.Controls.Add($button1)
    $form.controls.Add($label1)
    $form.ShowDialog()
}