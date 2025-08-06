Attribute VB_Name = "Module4"
Sub ComparerProgrammes()

  '---------------------- Optimisation pour acc�l�rer la macro --------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Dim wbClient As Workbook, wbISTA As Workbook, wbSortie As Workbook
    Dim wsClient As Worksheet, wsISTA As Worksheet, wsSortie As Worksheet
    
    ' Initialiser les variables de classeur � Nothing
    Set wbClient = Nothing
    Set wbISTA = Nothing
    Set wbSortie = Nothing
    Dim fichierClient As String, fichierISTA As String
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
    
    ' Demander les fichiers � l'utilisateur
    fichierClient = Application.GetOpenFilename("Fichiers Excel (*.xlsx;*.xls), *.xlsx;*.xls", , "S�lectionner le fichier CLIENT")
    If fichierClient = "False" Then Exit Sub
    
    fichierISTA = Application.GetOpenFilename("Fichiers Excel (*.xlsx;*.xls), *.xlsx;*.xls", , "S�lectionner le fichier ISTA")
    If fichierISTA = "False" Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ____________  S�lection du dossier de sauvegarde du fichier final csv ET du rapport d'anomalie ______________

    MsgBox "S�lectionner l'emplacement du nouveau fichier"
    fichierSortie = "UEX_Cli_CDC" & ".xlsx"
 
    'Choix du r�pertoire d'enregistrement du fichier trait�
    cheminSortie = Application.GetSaveAsFilename( _
        InitialFileName:=fichierSortie, _
        FileFilter:="Fichiers Excel (*.xlsx), *.xlsx", _
        Title:="Sauvegarder le fichier")
        
        ' V�rifier si l'utilisateur a annul� l'enregistrement
        If cheminSortie = False Then
            MsgBox "Enregistrement annul� par l'utilisateur !", vbInformation
            Exit Sub
        End If
    
    ' ---------------- Ouverture du fichier client -----------------------------------
    Set wbClient = Workbooks.Open(fichierClient)
    Set wsClient = wbClient.Worksheets("PATISTA")
    
    ' Trouver la derni�re ligne avec des donn�es dans la colonne E
    derniereLigneClient = wsClient.Cells(wsClient.Rows.Count, "E").End(xlUp).Row
    
     ' Compter les programmes du fichier client (colonne E)
    For i = 2 To derniereLigneClient ' Commencer � la ligne 2 pour �viter l'en-t�te
        ' Gestion s�curis�e de la valeur de cellule
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
                    ' Ajouter � la liste globale des programmes
                    If Not dictProgrammes.Exists(programme) Then
                        dictProgrammes.Add programme, True
                    End If
                End If
            End If
        End If
    Next i
    
    ' ---------------- Ouverture du fichier ISTA --------------------------------------------------------
    Set wbISTA = Workbooks.Open(fichierISTA)
    On Error Resume Next
    Set wsISTA = wbISTA.Worksheets("LOT 1 apr�s MAJ_BASE TRAVAIL")
    On Error GoTo GestionErreur
    
    If wsISTA Is Nothing Then
        MsgBox "La feuille 'LOT 1 apr�s MAJ_BASE TRAVAIL' est introuvable.", vbCritical
        GoTo NettoyageEtSortie
    End If
    
    ' Trouver la derni�re ligne avec des donn�es dans la colonne H
    derniereLigneISTA = wsISTA.Cells(wsISTA.Rows.Count, "H").End(xlUp).Row
    
    ' Compter les programmes du fichier ISTA (colonne H) - VERSION CORRIG�E
    For i = 4 To derniereLigneISTA ' Commencer � la ligne 4 pour �viter l'en-t�te
        ' Gestion s�curis�e de la valeur de cellule H
        On Error Resume Next
        cellValue = wsISTA.Cells(i, "H").Value
        On Error GoTo GestionErreur
        
        If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
            If VarType(cellValue) = vbString Or VarType(cellValue) = vbDouble Or VarType(cellValue) = vbInteger Or VarType(cellValue) = vbLong Then
                programme = Trim(CStr(cellValue))
                If programme <> "" Then
                    ' Gestion s�curis�e de la colonne F (code affaire)
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
                    
                    ' Ajouter � la liste globale des programmes
                    If Not dictProgrammes.Exists(programme) Then
                        dictProgrammes.Add programme, True
                    End If
                End If
            End If
        End If
    Next i
    
    ' -------------------------- Cr�ation de la feuille de r�sultats --------------------------------------
    ' Cr�er un nouveau classeur pour les r�sultats
    Set wbSortie = Workbooks.Add
    Set wsSortie = wbSortie.Worksheets(1)
    wsSortie.Name = "UEX CLI"
    
    ' Cr�er les en-t�tes
    With wsSortie
        .Cells(1, 1).Value = "Programme"
        .Cells(1, 2).Value = "PTC ISTA"
        .Cells(1, 3).Value = "PTC CLI"
        .Cells(1, 4).Value = "Code Affaire"    'UEX
        .Cells(1, 5).Value = "Delta positif"
        .Cells(1, 6).Value = "Delta n�gatif"
        
        ' Mettre en forme les en-t�tes
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 200, 200)
        .Range("A1:F1").Borders.LineStyle = xlContinuous
    End With
    
    ' Trier les programmes par ordre alphab�tique
    programmes = dictProgrammes.Keys
    Call TrierTableau(programmes)
    
    ligneResultat = 2
    
    ' -------------------------- G�n�ration du rapport de comparaison --------------------------------------
    
    ' Remplir les donn�es pour chaque programme
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
        
        ' Remplir la ligne de r�sultats
        With wsSortie
            .Cells(ligneResultat, 1).Value = programme
            .Cells(ligneResultat, 2).Value = ptcISTA
            .Cells(ligneResultat, 3).Value = ptcClient
            .Cells(ligneResultat, 4).Value = codeAffaire
            .Cells(ligneResultat, 5).Value = IIf(deltaPositif > 0, deltaPositif, "")
            .Cells(ligneResultat, 6).Value = IIf(deltaNegatif > 0, deltaNegatif, "")
            
             ' Figer les volets sous la ligne d'en-t�te (ligne 1)
            .Cells(2, 1).Select
            ActiveWindow.FreezePanes = True
            
            ' Colorier en vert si les nombres correspondent
            If ptcISTA = ptcClient And ptcISTA > 0 Then
                .Cells(ligneResultat, 2).Interior.Color = RGB(0, 255, 0)
                .Cells(ligneResultat, 3).Interior.Color = RGB(0, 255, 0)
            End If
            
            ' Colorier les deltas
            If deltaPositif > 0 Then
                .Cells(ligneResultat, 5).Interior.Color = RGB(144, 238, 144) ' Vert clair
                .Cells(ligneResultat, 5).Font.Color = RGB(0, 100, 0) ' Vert fonc�
            End If
            
            If deltaNegatif > 0 Then
                .Cells(ligneResultat, 6).Interior.Color = RGB(255, 182, 193) ' Rouge clair
                .Cells(ligneResultat, 6).Font.Color = RGB(139, 0, 0) ' Rouge fonc�
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
    
    ' --------------------------------- Cr�er un r�sum� -----------------------------------------
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
        ElseIf ptcISTA > ptcClient Then
            totalDeltaNegatif = totalDeltaNegatif + (ptcISTA - ptcClient)
        End If
    Next programme
    
    programmesDifferents = totalProgrammes - programmesIdentiques
    
    ' R�sum� en bas du fichier
    With wsSortie
        .Cells(ligneResultat + 2, 1).Value = "R�SUM� :"
        .Cells(ligneResultat + 2, 1).Font.Bold = True
        .Cells(ligneResultat + 3, 1).Value = "Total programmes :"
        .Cells(ligneResultat + 3, 2).Value = totalProgrammes
        .Cells(ligneResultat + 4, 1).Value = "Programmes avec comptage identique:"
        .Cells(ligneResultat + 4, 2).Value = programmesIdentiques
        .Cells(ligneResultat + 4, 2).Interior.Color = RGB(0, 255, 0)
        .Cells(ligneResultat + 5, 1).Value = "Programmes avec diff�rences:"
        .Cells(ligneResultat + 5, 2).Value = programmesDifferents
        If programmesDifferents > 0 Then .Cells(ligneResultat + 5, 2).Interior.Color = RGB(255, 255, 0)
        .Cells(ligneResultat + 6, 1).Value = "Total Delta Positif:"
        .Cells(ligneResultat + 6, 2).Value = totalDeltaPositif
        .Cells(ligneResultat + 6, 2).Interior.Color = RGB(144, 238, 144)
        .Cells(ligneResultat + 7, 1).Value = "Total Delta N�gatif:"
        .Cells(ligneResultat + 7, 2).Value = totalDeltaNegatif
        .Cells(ligneResultat + 7, 2).Interior.Color = RGB(255, 182, 193)
    End With
    
    ' -------------------------------- Cr�ation du rapport ANO ---------------------
    
    ' R�ouvrir les fichiers pour l'analyse des anomalies
    Set wbClient = Workbooks.Open(fichierClient)
    Set wsClient = wbClient.Worksheets("PATISTA")
    Set wbISTA = Workbooks.Open(fichierISTA)
    Set wsISTA = wbISTA.Worksheets("LOT 1 apr�s MAJ_BASE TRAVAIL")
    
    Dim wsRapportAffaire As Worksheet
    Set wsRapportAffaire = wbSortie.Sheets.Add(After:=wbSortie.Sheets(wbSortie.Sheets.Count))
    wsRapportAffaire.Name = "RAPPORT ANO"
    
    ' Dictionnaire : code affaire -> Collection de descriptions � afficher
    Dim dictAnomalies As Object
    Set dictAnomalies = CreateObject("Scripting.Dictionary")
    
    ' Structure : stocker la couleur � appliquer pour chaque description
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
        
        ' Traiter seulement les programmes avec des diff�rences
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
                            
                            ' V�rifier si cette description existe d�j� pour �viter les doublons
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
                                dictTypesDelta.Add desc & "_" & codeAffaire, "positif"  ' Cl� unique
                            End If
                        End If
                    End If
                Next i
            End If
            
            If deltaNegatif > 0 Then
                ' Delta n�gatif : chercher TOUTES les occurrences dans le fichier ISTA
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
                            
                            ' V�rifier si cette description existe d�j� pour �viter les doublons
                            existeDeja = False
                            For Each itemExistant In dictAnomalies(codeAffaire)
                                If itemExistant = desc Then
                                    existeDeja = True
                                    Exit For
                                End If
                            Next itemExistant
                            
                            If Not existeDeja Then
                                dictAnomalies(codeAffaire).Add desc
                                dictTypesDelta.Add desc & "_" & codeAffaire, "negatif"  ' Cl� unique
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
    
            ' En-t�te
            .Cells(1, colIndex).Value = colAffaire
            .Cells(1, colIndex).Font.Bold = True
            .Cells(1, colIndex).Interior.Color = RGB(200, 200, 200)
    
            Dim rowIndex As Long: rowIndex = 2
            For Each item In dictAnomalies(colAffaire)
                .Cells(rowIndex, colIndex).Value = item
    
                ' Rechercher le type de delta avec la cl� unique
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
            
            ' Ajouter des bordures
            .Range(.Cells(1, 1), .Cells(1, 10000)).Borders.LineStyle = xlContinuous
            
        End With
    Next colAffaire
    
    ' Ajuster les colonnes du rapport d'anomalies
    With wsRapportAffaire
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:Z").ColumnWidth = 25
        .Range("A1:A2").Font.Bold = True
        .Range("A1:A2").Interior.Color = RGB(220, 220, 220)
    End With
    
    ' Fermer les fichiers r�ouverts
    wbClient.Close SaveChanges:=False
    wbISTA.Close SaveChanges:=False
    
    ' --------------------------------Fin du programme ----------------------------------------------------
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    MsgBox " Analyse termin�e avec succ�s !" & vbCrLf & vbCrLf & _
           " Total programmes analys�s: " & totalProgrammes & vbCrLf & _
           " Programmes identiques: " & programmesIdentiques & vbCrLf & _
           " Programmes avec diff�rences: " & programmesDifferents & vbCrLf & vbCrLf & _
           " Feuille 'UEX CLI' cr��e avec succ�s!", vbInformation, "Comparaison termin�e"
    
           ' Enregistrer le fichier g�n�r�
        wbSortie.SaveAs Filename:=cheminSortie, FileFormat:=xlOpenXMLWorkbook
        
    ' Ouvrir le dossier contenant les fichier cr��s
    Shell "explorer.exe /select,""" & cheminSortie & """", vbNormalFocus
    
    Exit Sub

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
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Exit Sub

GestionErreur:
'    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True
'    Application.StatusBar = False

'    ' Fermer les fichiers en cas d'erreur
'    On Error Resume Next
'    If Not (wbClient Is Nothing) Then
'        wbClient.Close SaveChanges:=False
'        Set wbClient = Nothing
'    End If
'    If Not (wbISTA Is Nothing) Then
'        wbISTA.Close SaveChanges:=False
'        Set wbISTA = Nothing
'    End If
'    If Not (wbSortie Is Nothing) Then
'        Set wbSortie = Nothing
'    End If
'    On Error GoTo 0
'
'    MsgBox " Erreur lors de l'ex�cution: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
'           " V�rifiez que:" & vbCrLf & _
'           " - Les fichiers existent et sont accessibles" & vbCrLf & _
'           " - Les feuilles 'PATISTA' et 'LOT 1 apr�s MAJ_BASE TRAVAIL' existent" & vbCrLf & _
'           " - Les colonnes E et H contiennent les donn�es de programmes", vbCritical, "Erreur"    ' ----------------------------------- Restautrer les param�tres --------------------------------
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

