Attribute VB_Name = "Module6"
'Salamata Nourou MBAYE

Sub ComparerProgrammes()

  '---------------------- Optimisation pour accélérer la macro --------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Dim wbClient As Workbook, wbISTA As Workbook, wbSortie As Workbook
    Dim wsClient As Worksheet, wsISTA As Worksheet, wsSortie As Worksheet
    Dim fdlg As FileDialog
    Dim fdlgDossier As FileDialog
    Dim dossierSauvegarde As String
    
    ' Initialiser les variables de classeur à Nothing
    Set wbClient = Nothing
    Set wbISTA = Nothing
    Set wbSortie = Nothing
    
    Dim cheminFichierClient As String, cheminFichierIsta As String
    Dim fichierSortie As String, cheminSortie As String
    Dim derniereLigneClient As Long, derniereLigneISTA As Long
    Dim i As Long, j As Long, ligneResultat As Long
    Dim dictProgrammes As Object, dictClient As Object, dictISTA As Object
    Dim dictAffaires As Object
    Dim programme As Variant, programmes As Variant
    Dim ptcClient As Long, ptcISTA As Long
    Dim codeAffaire As String, deltaPositif As Long, deltaNegatif As Long
    Dim cellValue As Variant  ' Variable pour stocker la valeur de la cellule
    
    ' Initialiser les dictionnaires
    Set dictProgrammes = CreateObject("Scripting.Dictionary")
    Set dictClient = CreateObject("Scripting.Dictionary")
    Set dictISTA = CreateObject("Scripting.Dictionary")
    Set dictAffaires = CreateObject("Scripting.Dictionary")
    
    On Error GoTo GestionErreur
    
     ' ------------- Sélection du premier fichier (CLIENT) ---------------
     MsgBox "Sélectionner le fichier client"
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Étape 1/2 : Choisir le fichier client obligatoirement"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers Excel", "*.xlsx;*.xls;*.xlsm"
    fdlg.AllowMultiSelect = False
    
    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        GoTo NettoyageEtSortie
    End If
    
    cheminFichierClient = fdlg.SelectedItems(1)
    
    ' ------------------ Sélection du deuxième fichier (ISTA) ------------
    MsgBox "Sélectionner le fichier ISTA "
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Étape 2/2 : Choisir le fichier d'extraction ISTA obligatoirement"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers Excel", "*.xlsx;*.xls;*.xlsm"
    fdlg.AllowMultiSelect = False
    
    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        GoTo NettoyageEtSortie
    End If
    
    cheminFichierIsta = fdlg.SelectedItems(1)
     
    ' --------------- Vérification des fichiers -------------
    If Dir(cheminFichierClient) = "" Then
        MsgBox "Le fichier Client n'existe pas : " & cheminFichierClient, vbCritical
        GoTo NettoyageEtSortie
    End If
    
    If Dir(cheminFichierIsta) = "" Then
        MsgBox "Le fichier ISTA n'existe pas : " & cheminFichierIsta, vbCritical
        GoTo NettoyageEtSortie
    End If
    
    ' Vérifier que les fichiers sélectionnés soient différents
    If cheminFichierClient = cheminFichierIsta Then
        If MsgBox("Attention ! Vous avez sélectionné le même fichier deux fois." & vbCrLf & _
                 "Voulez-vous continuer quand même ?", vbExclamation + vbYesNo) = vbNo Then
            GoTo NettoyageEtSortie
        End If
    End If
    
    ' Ouvrir les fichiers sources
    On Error Resume Next
    Set wbClient = Workbooks.Open(cheminFichierClient, ReadOnly:=True)
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'ouverture du fichier client : " & Err.Description, vbCritical
        GoTo NettoyageEtSortie
    End If
    wbClient.Windows(1).Visible = False
    
    Set wbISTA = Workbooks.Open(cheminFichierIsta, ReadOnly:=True)
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'ouverture du fichier ISTA : " & Err.Description, vbCritical
        GoTo NettoyageEtSortie
    End If
    wbISTA.Windows(1).Visible = False
    On Error GoTo GestionErreur
    
    ' Sélection du dossier de sauvegarde du fichier final
    MsgBox "Choisir le dossier dans lequel le fichier de résultat doit être enregistré"
    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlgDossier
        .Title = "Choisir le dossier où enregistrer le fichier de résultat"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
    End With
    
    If fdlgDossier.Show <> -1 Then
        MsgBox "Sélection du dossier annulée par l'utilisateur.", vbInformation
        GoTo NettoyageEtSortie
    End If
    
    dossierSauvegarde = fdlgDossier.SelectedItems(1)
    
    ' Vérifier que le dossier existe et est accessible
    If Dir(dossierSauvegarde, vbDirectory) = "" Then
        MsgBox "Le dossier sélectionné n'est pas accessible : " & dossierSauvegarde, vbCritical
        GoTo NettoyageEtSortie
    End If

    fichierSortie = "UEX_Cli_CDC_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"  '
    cheminSortie = dossierSauvegarde & "\" & fichierSortie

    ' Références aux feuilles
    Set wsClient = wbClient.Worksheets("PATISTA")
    Set wsISTA = wbISTA.Worksheets("LOT 1 après MAJ_BASE TRAVAIL")
    
    ' --------------------- Ouverture du fichier client ----------------------------------
    ' Trouver la dernière ligne avec des données dans la colonne E
    derniereLigneClient = wsClient.Cells(wsClient.Rows.Count, "E").End(xlUp).Row
    
     ' Compter les programmes du fichier client (colonne E)
    For i = 2 To derniereLigneClient ' Commencer à la ligne 2 pour éviter l'en-tête
        ' Gestion sécurisée de la valeur de cellule
        cellValue = wsClient.Cells(i, "E").Value
        If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
            If VarType(cellValue) = vbString Or VarType(cellValue) = vbDouble Or VarType(cellValue) = vbInteger Or VarType(cellValue) = vbLong Then
                programme = Trim(CStr(cellValue))
                If programme <> "" Then
                    If dictClient.Exists(programme) Then
                        dictClient(programme) = dictClient(programme) + 1
                    Else
                        dictClient.Add programme, 1
                    End If
                    ' Ajouter à la liste globale des programmes
                    If Not dictProgrammes.Exists(programme) Then
                        dictProgrammes.Add programme, True
                    End If
                End If
            End If
        End If
    Next i
    
    ' ---------------- Ouverture du fichier ISTA --------------------------------------------------------
    
    ' Trouver la dernière ligne avec des données dans la colonne H
    derniereLigneISTA = wsISTA.Cells(wsISTA.Rows.Count, "H").End(xlUp).Row
    
    ' Compter les programmes du fichier ISTA (colonne H)
    For i = 4 To derniereLigneISTA ' Commencer à la ligne 4 pour éviter l'en-tête
        ' Gestion sécurisée de la valeur de cellule H
        On Error Resume Next
        cellValue = wsISTA.Cells(i, "H").Value
        On Error GoTo GestionErreur
        
        If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
            If VarType(cellValue) = vbString Or VarType(cellValue) = vbDouble Or VarType(cellValue) = vbInteger Or VarType(cellValue) = vbLong Then
                programme = Trim(CStr(cellValue))
                If programme <> "" Then
                    ' Gestion sécurisée de la colonne F (code affaire)
                    On Error Resume Next
                    Dim codeAffaireTemp As Variant
                    codeAffaireTemp = wsISTA.Cells(i, "F").Value
                    On Error GoTo GestionErreur
                    
                    If Not IsEmpty(codeAffaireTemp) And Not IsNull(codeAffaireTemp) Then
                        codeAffaire = Trim(CStr(codeAffaireTemp))
                    Else
                        codeAffaire = ""
                    End If
                    
                    If dictISTA.Exists(programme) Then
                        dictISTA(programme) = dictISTA(programme) + 1
                    Else
                        dictISTA.Add programme, 1
                    End If
                    
                     ' Stocker l'affaire pour ce programme
                    If Not dictAffaires.Exists(programme) Then
                        dictAffaires.Add programme, codeAffaire
                    End If
                    
                    ' Ajouter à la liste globale des programmes
                    If Not dictProgrammes.Exists(programme) Then
                        dictProgrammes.Add programme, True
                    End If
                End If
            End If
        End If
    Next i
    
    ' -------------------------- Création de la feuille de résultats --------------------------------------
    ' Créer un nouveau classeur pour les résultats
    Set wbSortie = Workbooks.Add
    Set wsSortie = wbSortie.Worksheets(1)
    wsSortie.Name = "UEX CLI"
    
    ' Créer les en-têtes
    With wsSortie
        .Cells(1, 1).Value = "Programme"
        .Cells(1, 2).Value = "PTC ISTA"
        .Cells(1, 3).Value = "PTC CLI"
        .Cells(1, 4).Value = "Code Affaire"
        .Cells(1, 5).Value = "Delta positif"
        .Cells(1, 6).Value = "Delta négatif"
        
        ' Mettre en forme les en-têtes
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 200, 200)
        .Range("A1:F1").Borders.LineStyle = xlContinuous
        
        ' CORRECTION: Figer les volets après avoir créé les en-têtes
        .Cells(2, 1).Select
        ActiveWindow.FreezePanes = True
    End With
    
    ' Trier les programmes par ordre alphabétique
    programmes = dictProgrammes.Keys
    Call TrierTableau(programmes)
    
    ligneResultat = 2
    
    ' -------------------------- Génération du rapport de comparaison --------------------------------------
    
    ' Remplir les données pour chaque programme
    For Each programme In programmes
        ' Obtenir les comptages
        ptcClient = 0
        ptcISTA = 0
        codeAffaire = ""
        deltaPositif = 0
        deltaNegatif = 0
        
        If dictClient.Exists(programme) Then ptcClient = dictClient(programme)
        If dictISTA.Exists(programme) Then ptcISTA = dictISTA(programme)
        If dictAffaires.Exists(programme) Then codeAffaire = dictAffaires(programme)
        
           ' Calculer les deltas
        If ptcISTA < ptcClient Then
            deltaPositif = ptcClient - ptcISTA
        ElseIf ptcISTA > ptcClient Then
            deltaNegatif = ptcISTA - ptcClient
        End If
        
        ' Remplir la ligne de résultats
        With wsSortie
            .Cells(ligneResultat, 1).Value = programme
            .Cells(ligneResultat, 2).Value = ptcISTA
            .Cells(ligneResultat, 3).Value = ptcClient
            .Cells(ligneResultat, 4).Value = codeAffaire
            .Cells(ligneResultat, 5).Value = IIf(deltaPositif > 0, deltaPositif, "")
            .Cells(ligneResultat, 6).Value = IIf(deltaNegatif > 0, deltaNegatif, "")
            
            ' Colorier en vert si les nombres correspondent
            If ptcISTA = ptcClient And ptcISTA > 0 Then
                .Cells(ligneResultat, 2).Interior.Color = RGB(0, 255, 0)
                .Cells(ligneResultat, 3).Interior.Color = RGB(0, 255, 0)
            End If
            
            ' Colorier les deltas
            If deltaPositif > 0 Then
                .Cells(ligneResultat, 5).Interior.Color = RGB(144, 238, 144) ' Vert clair
                .Cells(ligneResultat, 5).Font.Color = RGB(0, 100, 0) ' Vert foncé
            End If
            
            If deltaNegatif > 0 Then
                .Cells(ligneResultat, 6).Interior.Color = RGB(255, 182, 193) ' Rouge clair
                .Cells(ligneResultat, 6).Font.Color = RGB(139, 0, 0) ' Rouge foncé
            End If
            
            ' Ajouter des bordures
            .Range(.Cells(ligneResultat, 1), .Cells(ligneResultat, 6)).Borders.LineStyle = xlContinuous
        End With
        
        ligneResultat = ligneResultat + 1
    Next programme
    
    ' Ajuster la largeur des colonnes
    wsSortie.Columns("A:A").ColumnWidth = 40
    wsSortie.Columns("B:F").ColumnWidth = 13
    
    ' Fermer les fichiers source
    wbClient.Close SaveChanges:=False
    wbISTA.Close SaveChanges:=False
    ' CORRECTION: Réinitialiser les variables pour éviter les erreurs dans le nettoyage
    Set wbClient = Nothing
    Set wbISTA = Nothing
    
    ' --------------------------------- Créer un résumé -----------------------------------------
    Dim totalProgrammes As Long, programmesIdentiques As Long, programmesDifferents As Long
    Dim totalDeltaPositif As Long, totalDeltaNegatif As Long
    totalProgrammes = dictProgrammes.Count
    
    ' Compter les programmes avec comptage identique et calculer les totaux des deltas
    For Each programme In programmes
        ptcClient = 0
        ptcISTA = 0
        If dictClient.Exists(programme) Then ptcClient = dictClient(programme)
        If dictISTA.Exists(programme) Then ptcISTA = dictISTA(programme)
        
        If ptcISTA = ptcClient And ptcISTA > 0 Then
            programmesIdentiques = programmesIdentiques + 1
        ElseIf ptcISTA < ptcClient Then
            totalDeltaPositif = totalDeltaPositif + (ptcClient - ptcISTA)
            programmesDifferents = programmesDifferents + 1
        ElseIf ptcISTA > ptcClient Then
            totalDeltaNegatif = totalDeltaNegatif + (ptcISTA - ptcClient)
            programmesDifferents = programmesDifferents + 1
        End If
    Next programme
    
    ' CORRECTION: Supprimé la ligne qui recalculait programmesDifferents incorrectement
    
    ' Résumé en bas du fichier
    With wsSortie
        .Cells(ligneResultat + 2, 1).Value = "RÉSUMÉ :"
        .Cells(ligneResultat + 2, 1).Font.Bold = True
        .Cells(ligneResultat + 3, 1).Value = "Total programmes :"
        .Cells(ligneResultat + 3, 2).Value = totalProgrammes
        .Cells(ligneResultat + 4, 1).Value = "Programmes avec comptage identique:"
        .Cells(ligneResultat + 4, 2).Value = programmesIdentiques
        .Cells(ligneResultat + 4, 2).Interior.Color = RGB(0, 255, 0)
        .Cells(ligneResultat + 5, 1).Value = "Programmes avec différences:"
        .Cells(ligneResultat + 5, 2).Value = programmesDifferents
        If programmesDifferents > 0 Then .Cells(ligneResultat + 5, 2).Interior.Color = RGB(255, 255, 0)
        .Cells(ligneResultat + 6, 1).Value = "Total Delta Positif:"
        .Cells(ligneResultat + 6, 2).Value = totalDeltaPositif
        .Cells(ligneResultat + 6, 2).Interior.Color = RGB(144, 238, 144)
        .Cells(ligneResultat + 7, 1).Value = "Total Delta Négatif:"
        .Cells(ligneResultat + 7, 2).Value = totalDeltaNegatif
        .Cells(ligneResultat + 7, 2).Interior.Color = RGB(255, 182, 193)
    End With
    
    ' -------------------------------- Création du rapport ANO ---------------------
    
    ' Réouvrir les fichiers pour l'analyse des anomalies
    Set wbClient = Workbooks.Open(cheminFichierClient, ReadOnly:=True)
    Set wsClient = wbClient.Worksheets("PATISTA")
    Set wbISTA = Workbooks.Open(cheminFichierIsta, ReadOnly:=True)
    Set wsISTA = wbISTA.Worksheets("LOT 1 après MAJ_BASE TRAVAIL")
    
    Dim wsRapportAffaire As Worksheet
    Set wsRapportAffaire = wbSortie.Sheets.Add(After:=wbSortie.Sheets(wbSortie.Sheets.Count))
    wsRapportAffaire.Name = "RAPPORT ANO"
    
    ' Dictionnaire : code affaire -> Collection de descriptions à afficher
    Dim dictAnomalies As Object
    Set dictAnomalies = CreateObject("Scripting.Dictionary")
    
    ' Structure : stocker la couleur à appliquer pour chaque description
    Dim dictTypesDelta As Object
    Set dictTypesDelta = CreateObject("Scripting.Dictionary")
    
    Dim desc As String
    For Each programme In programmes
        ptcClient = 0
        ptcISTA = 0
        codeAffaire = ""
        deltaPositif = 0
        deltaNegatif = 0
        desc = ""
    
        If dictClient.Exists(programme) Then ptcClient = dictClient(programme)
        If dictISTA.Exists(programme) Then ptcISTA = dictISTA(programme)
        If dictAffaires.Exists(programme) Then codeAffaire = dictAffaires(programme)
    
        ' Calculer les deltas
        If ptcISTA < ptcClient Then
            deltaPositif = ptcClient - ptcISTA
        ElseIf ptcISTA > ptcClient Then
            deltaNegatif = ptcISTA - ptcClient
        End If
        
        ' Traiter seulement les programmes avec des différences
        If deltaPositif > 0 Or deltaNegatif > 0 Then
            
            If deltaPositif > 0 Then
                ' Delta positif : chercher TOUTES les occurrences dans le fichier CLIENT
                For i = 2 To derniereLigneClient
                    On Error Resume Next
                    cellValue = wsClient.Cells(i, "E").Value
                    On Error GoTo GestionErreur
                    
                    If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
                        If Trim(CStr(cellValue)) = programme Then
                            Dim refLogement As String, ptcValue As String
                            
                            On Error Resume Next
                            refLogement = CStr(wsClient.Cells(i, "L").Value)
                            ptcValue = CStr(wsClient.Cells(i, "X").Value)
                            On Error GoTo GestionErreur
                            
                            desc = refLogement & "-" & ptcValue
                            
                            ' Ajouter cette occurrence
                            If Not dictAnomalies.Exists(codeAffaire) Then
                                dictAnomalies.Add codeAffaire, New Collection
                            End If
                            
                            ' Vérifier si cette description existe déjà pour éviter les doublons
                            Dim existeDeja As Boolean
                            existeDeja = False
                            Dim itemExistant As Variant
                            For Each itemExistant In dictAnomalies(codeAffaire)
                                If itemExistant = desc Then
                                    existeDeja = True
                                    Exit For
                                End If
                            Next itemExistant
                            
                            If Not existeDeja Then
                                dictAnomalies(codeAffaire).Add desc
                                dictTypesDelta.Add desc & "_" & codeAffaire, "positif"
                            End If
                        End If
                    End If
                Next i
            End If
            
            If deltaNegatif > 0 Then
                ' Delta négatif : chercher TOUTES les occurrences dans le fichier ISTA
                For i = 4 To derniereLigneISTA
                    On Error Resume Next
                    cellValue = wsISTA.Cells(i, "H").Value
                    On Error GoTo GestionErreur
                    
                    If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
                        If Trim(CStr(cellValue)) = programme Then
                            On Error Resume Next
                            refLogement = CStr(wsISTA.Cells(i, "W").Value)
                            ptcValue = CStr(wsISTA.Cells(i, "N").Value)
                            On Error GoTo GestionErreur
                            
                            desc = refLogement & "-" & ptcValue
                            
                            ' Ajouter cette occurrence
                            If Not dictAnomalies.Exists(codeAffaire) Then
                                dictAnomalies.Add codeAffaire, New Collection
                            End If
                            
                            ' Vérifier si cette description existe déjà pour éviter les doublons
                            Dim existeDeja2 As Boolean  ' C
                            existeDeja2 = False
                            Dim itemExistant2 As Variant
                            For Each itemExistant2 In dictAnomalies(codeAffaire)
                                If itemExistant2 = desc Then
                                    existeDeja2 = True
                                    Exit For
                                End If
                            Next itemExistant2
                            
                            If Not existeDeja2 Then
                                dictAnomalies(codeAffaire).Add desc
                                dictTypesDelta.Add desc & "_" & codeAffaire, "negatif"
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next programme
    
    ' Affichage dans la feuille
    Dim colIndex As Long: colIndex = 2
    Dim maxLignes As Long: maxLignes = 0
    Dim colAffaire As Variant, item As Variant
    
    For Each colAffaire In dictAnomalies.Keys
        With wsRapportAffaire
        
            .Cells(1, 1).Value = "Code Affaire"
            .Cells(2, 1).Value = "Ref LGT - PTC"
    
            ' En-tête
            .Cells(1, colIndex).Value = colAffaire
            .Cells(1, colIndex).Font.Bold = True
            .Cells(1, colIndex).Interior.Color = RGB(200, 200, 200)
    
            Dim rowIndex As Long: rowIndex = 2
            For Each item In dictAnomalies(colAffaire)
                .Cells(rowIndex, colIndex).Value = item
    
                ' Rechercher le type de delta avec la clé unique
                Dim cleRecherche As String
                cleRecherche = item & "_" & colAffaire
                
                If dictTypesDelta.Exists(cleRecherche) Then
                    If dictTypesDelta(cleRecherche) = "positif" Then
                        .Cells(rowIndex, colIndex).Interior.Color = RGB(144, 238, 144)
                        .Cells(rowIndex, colIndex).Font.Color = RGB(0, 100, 0)
                    ElseIf dictTypesDelta(cleRecherche) = "negatif" Then
                        .Cells(rowIndex, colIndex).Interior.Color = RGB(255, 182, 193)
                        .Cells(rowIndex, colIndex).Font.Color = RGB(139, 0, 0)
                    End If
                End If
    
                rowIndex = rowIndex + 1
            Next item
    
            If rowIndex > maxLignes Then maxLignes = rowIndex
            colIndex = colIndex + 1
            
        End With
    Next colAffaire
    
    ' Ajuster les colonnes du rapport d'anomalies
    With wsRapportAffaire
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:Z").ColumnWidth = 25
        .Range("A1:A2").Font.Bold = True
        .Range("A1:A2").Interior.Color = RGB(220, 220, 220)
        ' Ajout des bordures
        If maxLignes > 0 Then
            .Range("A1:" & Chr(65 + colIndex - 1) & maxLignes).Borders.LineStyle = xlContinuous
        End If
    End With
    
    ' Fermer les fichiers réouverts
    wbClient.Close SaveChanges:=False
    wbISTA.Close SaveChanges:=False
    Set wbClient = Nothing
    Set wbISTA = Nothing
    
    '  Enregistrer le fichier
    On Error Resume Next
    wbSortie.SaveAs Filename:=cheminSortie, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la sauvegarde du fichier : " & Err.Description & vbCrLf & _
               "Chemin : " & cheminSortie, vbCritical
        GoTo GestionErreur
    End If
    On Error GoTo GestionErreur
    
    ' --------------------------------Fin du programme ----------------------------------------------------
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox " Analyse terminée avec succès !" & vbCrLf & vbCrLf & _
           " Total programmes analysés: " & totalProgrammes & vbCrLf & _
           " Programmes identiques: " & programmesIdentiques & vbCrLf & _
           " Programmes avec différences: " & programmesDifferents & vbCrLf & vbCrLf & _
           " Fichier sauvegardé : " & cheminSortie, vbInformation, "Comparaison terminée"
    
    ' Ouvrir le dossier contenant le fichier créé
    Shell "explorer.exe /select,""" & cheminSortie & """", vbNormalFocus
    
    Exit Sub

GestionErreur:
    ' Restaurer les paramètres même en cas d'erreur
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    MsgBox " Erreur lors de l'exécution: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           " Vérifiez que:" & vbCrLf & _
           " - Les fichiers existent et sont accessibles" & vbCrLf & _
           " - Les feuilles 'PATISTA' et 'LOT 1 après MAJ_BASE TRAVAIL' existent" & vbCrLf & _
           " - Les colonnes E et H contiennent les données de programmes" & vbCrLf & _
           " - Vous avez les droits d'écriture dans le dossier de destination", vbCritical, "Erreur"

NettoyageEtSortie:
    ' Fermer les fichiers en cas d'erreur
    On Error Resume Next
    If Not (wbClient Is Nothing) Then
        wbClient.Close SaveChanges:=False
        Set wbClient = Nothing
    End If
    If Not (wbISTA Is Nothing) Then
        wbISTA.Close SaveChanges:=False
        Set wbISTA = Nothing
    End If
    If Not (wbSortie Is Nothing) Then
        Set wbSortie = Nothing
    End If
    On Error GoTo 0
    
    ' Restaurer les paramètres
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub




' Fonction pour trier un tableau
Sub TrierTableau(ByRef arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub


