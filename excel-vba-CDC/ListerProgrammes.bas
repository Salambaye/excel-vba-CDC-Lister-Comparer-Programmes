Attribute VB_Name = "Module1"
Sub ComparerProgrammes()

    Dim wbClient As Workbook, wbISTA As Workbook, wbSortie As Workbook
    Dim wsClient As Worksheet, wsISTA As Worksheet, wsSortie As Worksheet
    Dim fichierClient As String, fichierISTA As String
    Dim derniereLigneClient As Long, derniereLigneISTA As Long
    Dim i As Long, j As Long, ligneResultat As Long
    Dim dictProgrammes As Object, dictClient As Object, dictISTA As Object
    Dim dictAffaires As Object
    Dim programme As Variant, programmes As Variant
    Dim ptcClient As Long, ptcISTA As Long
    Dim codeAffaire As String, deltaPositif As Long, deltaNegatif As Long
    'Dim programmesManquants As String
    
    ' Initialiser les dictionnaires
    Set dictProgrammes = CreateObject("Scripting.Dictionary")
    Set dictClient = CreateObject("Scripting.Dictionary")
    Set dictISTA = CreateObject("Scripting.Dictionary")
    Set dictAffaires = CreateObject("Scripting.Dictionary")
    
    On Error GoTo GestionErreur
    
    ' Demander les fichiers à l'utilisateur
    fichierClient = Application.GetOpenFilename("Fichiers Excel (*.xlsx;*.xls), *.xlsx;*.xls", , "Sélectionner le fichier CLIENT")
    If fichierClient = "False" Then Exit Sub
    
    fichierISTA = Application.GetOpenFilename("Fichiers Excel (*.xlsx;*.xls), *.xlsx;*.xls", , "Sélectionner le fichier ISTA")
    If fichierISTA = "False" Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ---------------- Ouverture du fichier client -----------------------------------
    Set wbClient = Workbooks.Open(fichierClient)
    Set wsClient = wbClient.Worksheets("PATISTA")
    
    ' Trouver la dernière ligne avec des données dans la colonne E
    derniereLigneClient = wsClient.Cells(wsClient.Rows.Count, "E").End(xlUp).Row
    
     ' Compter les programmes du fichier client (colonne E)
    For i = 2 To derniereLigneClient ' Commencer à la ligne 2 pour éviter l'en-tête
        If wsClient.Cells(i, "E").Value <> "" Then
            programme = Trim(wsClient.Cells(i, "E").Value)
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
    Next i
    
    
     ' ---------------- Ouverture du fichier ISTA --------------------------------------------------------
     
    Set wbISTA = Workbooks.Open(fichierISTA)
    Set wsISTA = wbISTA.Worksheets("LOT 1 après MAJ_BASE TRAVAIL")
    
    ' Trouver la dernière ligne avec des données dans la colonne H
    derniereLigneISTA = wsISTA.Cells(wsISTA.Rows.Count, "H").End(xlUp).Row
    
    ' Compter les programmes du fichier ISTA (colonne H)
    For i = 4 To derniereLigneISTA ' Commencer à la ligne 4 pour éviter l'en-tête
        If wsISTA.Cells(i, "H").Value <> "" Then
            programme = Trim(wsISTA.Cells(i, "H").Value)
            codeAffaire = Trim(wsISTA.Cells(i, "F").Value) ' Colonne F pour les affaires
            
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
        .Cells(1, 4).Value = "Code Affaire"    'UEX
        .Cells(1, 5).Value = "Delta positif"
        .Cells(1, 6).Value = "Delta négatif"
        
        ' Mettre en forme les en-têtes
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 200, 200)
        .Range("A1:F1").Borders.LineStyle = xlContinuous
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
            
             ' Figer les volets sous la ligne d'en-tête (ligne 1)
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
    wsSortie.Columns("B:F").ColumnWidth = 13  '
    
    ' Fermer les fichiers source
    wbClient.Close SaveChanges:=False
    wbISTA.Close SaveChanges:=False
    
    
    ' --------------------------------- Créer un résumé
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
    
    ' Ajouter le résumé en bas
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
    
    
        ' -------------------------------- Création du rapport d'anomalies --------------------------------
    Dim wsAnomalies As Worksheet
    Set wsAnomalies = wbSortie.Sheets.Add(After:=wbSortie.Sheets(wbSortie.Sheets.Count))
    wsAnomalies.Name = "RAPPORT ANO"

    ' En-têtes
    With wsAnomalies
        .Cells(1, 1).Value = "Code Affaire"
        .Cells(1, 2).Value = "Programme"
        .Cells(1, 3).Value = "Delta positif"
        .Cells(1, 4).Value = "Delta négatif"
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(200, 200, 200)
        .Range("A1:D1").Borders.LineStyle = xlContinuous
    End With

    Dim ligneAno As Long
    ligneAno = 2

    ' Remplissage du rapport d'anomalies
    For Each programme In programmes
        ptcClient = 0
        ptcISTA = 0
        codeAffaire = ""
        deltaPositif = 0
        deltaNegatif = 0

        If dictClient.Exists(programme) Then ptcClient = dictClient(programme)
        If dictISTA.Exists(programme) Then ptcISTA = dictISTA(programme)
        If dictAffaires.Exists(programme) Then codeAffaire = dictAffaires(programme)

        If ptcISTA < ptcClient Then
            deltaPositif = ptcClient - ptcISTA
        ElseIf ptcISTA > ptcClient Then
            deltaNegatif = ptcISTA - ptcClient
        End If

        If deltaPositif > 0 Or deltaNegatif > 0 Then
            With wsAnomalies
                .Cells(ligneAno, 1).Value = codeAffaire
                .Cells(ligneAno, 2).Value = programme
                .Cells(ligneAno, 3).Value = IIf(deltaPositif > 0, deltaPositif, "")
                .Cells(ligneAno, 4).Value = IIf(deltaNegatif > 0, deltaNegatif, "")

                ' Colorier les deltas
                If deltaPositif > 0 Then
                    .Cells(ligneAno, 3).Interior.Color = RGB(144, 238, 144)
                    .Cells(ligneAno, 3).Font.Color = RGB(0, 100, 0)
                End If
                If deltaNegatif > 0 Then
                    .Cells(ligneAno, 4).Interior.Color = RGB(255, 182, 193)
                    .Cells(ligneAno, 4).Font.Color = RGB(139, 0, 0)
                End If

                ' Bordures
                .Range(.Cells(ligneAno, 1), .Cells(ligneAno, 4)).Borders.LineStyle = xlContinuous
            End With

            ligneAno = ligneAno + 1
        End If
    Next programme

    ' Ajustement des colonnes
    wsAnomalies.Columns("A:D").AutoFit



    
    
    
    ' --------------------------------Fin du programme ----------------------------------------------------
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    MsgBox " Analyse terminée avec succès !" & vbCrLf & vbCrLf & _
           " Total programmes analysés: " & totalProgrammes & vbCrLf & _
           " Programmes identiques: " & programmesIdentiques & vbCrLf & _
           " Programmes avec différences: " & programmesDifferents & vbCrLf & vbCrLf & _
           " Feuille 'UEX CLI' créée avec succès!", vbInformation, "Comparaison terminée"
    
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    ' Fermer les fichiers en cas d'erreur
    On Error Resume Next
    If Not wbClient Is Nothing Then wbClient.Close SaveChanges:=False
    If Not wbISTA Is Nothing Then wbISTA.Close SaveChanges:=False
    On Error GoTo 0
    
    MsgBox " Erreur lors de l'exécution: " & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           " Vérifiez que:" & vbCrLf & _
           " Les fichiers existent et sont accessibles" & vbCrLf & _
           " Les feuilles 'PATISTA' et 'LOT 1 après MAJ_BASE TRAVAIL' existent" & vbCrLf & _
           "Les colonnes E et H contiennent les données de programmes", vbCritical, "Erreur"

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
