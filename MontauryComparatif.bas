Attribute VB_Name = "MontauryComparatif"
Option Explicit

' ============================================================
' MODULE : MontauryComparatif
' Comparaison annuelle N / N-1 des fichiers RPS - Montaury M2
' ------------------------------------------------------------
' Fichier 2026 (N)   : D:\ASIA\asia-tim\rimp\MONTAURY\M2\Montaury-2026M2-rps.txt
' Fichier 2025 (N-1) : D:\ASIA\rimp_2025\MONTAURY\M2\Montaury-2025M2-rps.txt
' ------------------------------------------------------------
' Fonctionnalités :
'   - Lecture et agrégation par (Activité, Établissement, Mois)
'   - Tri des clés par tableau (aucun Collection.Add positionnel)
'   - Tableau comparatif complet même si un fichier est vide
'   - Mise en page : titres, couleurs, bordures, colonnes figées
'   - Module auto-suffisant : aucune dépendance externe
' ============================================================

' --- Chemins des deux fichiers à comparer ---
Private Const FICHIER_2026 As String = "D:\ASIA\asia-tim\rimp\MONTAURY\M2\Montaury-2026M2-rps.txt"
Private Const FICHIER_2025 As String = "D:\ASIA\rimp_2025\MONTAURY\M2\Montaury-2025M2-rps.txt"

' --- Libellés des années ---
Private Const ANNEE_N   As String = "2026"
Private Const ANNEE_N1  As String = "2025"
Private Const ETAB_NOM  As String = "Montaury"
Private Const PERIODE   As String = "M2"

' --- Séparateur de champ dans les fichiers RPS ---
Private Const SEP As String = ";"

' --- Indices de colonnes dans le fichier RPS (base 1) ---
' Adaptez ces constantes si la structure de vos fichiers diffère.
Private Const COL_FINESS   As Integer = 1   ' Code FINESS / établissement
Private Const COL_ACTIVITE As Integer = 2   ' Code activité (ex : 1, 2, 3...)
Private Const COL_MOIS     As Integer = 3   ' Mois (1..12)
Private Const COL_ACTE     As Integer = 4   ' Code acte / prestation
Private Const COL_VALEUR   As Integer = 5   ' Valeur / nombre

' --- Couleurs de la mise en page ---
Private Const COULEUR_TITRE      As Long = 1644825    ' Bleu marine foncé  (#191970)
Private Const COULEUR_EN_TETE    As Long = 5065970    ' Bleu moyen         (#4D77C2)
Private Const COULEUR_LIGNE_PAIR As Long = 13561798   ' Bleu très clair    (#CFDFF6)
Private Const COULEUR_LIGNE_IMP  As Long = 16777215   ' Blanc              (#FFFFFF)
Private Const COULEUR_TOTAL      As Long = 16763904   ' Jaune pâle         (#FFD700 approx)
Private Const COULEUR_VARIATION_POS As Long = 13434828 ' Vert clair        (#CCFFCC)
Private Const COULEUR_VARIATION_NEG As Long = 16753920 ' Rouge/orange pâle (#FFAA00 approx)
Private Const COULEUR_TEXTE_BLANC As Long = 16777215  ' Blanc

' ============================================================
' POINT D'ENTRÉE PRINCIPAL
' ============================================================
Sub LancerComparatifMontaury()

    ' ---- Vérification de l'existence des fichiers ----
    Dim ok2026 As Boolean, ok2025 As Boolean
    ok2026 = (Dir(FICHIER_2026) <> "")
    ok2025 = (Dir(FICHIER_2025) <> "")

    If Not ok2026 And Not ok2025 Then
        MsgBox "Aucun des deux fichiers n'a été trouvé :" & vbCrLf & vbCrLf & _
               "  2026 : " & FICHIER_2026 & vbCrLf & _
               "  2025 : " & FICHIER_2025, vbCritical, "Fichiers introuvables"
        Exit Sub
    End If

    Dim avertissements As String
    avertissements = ""
    If Not ok2026 Then
        avertissements = avertissements & "ATTENTION : fichier 2026 introuvable, la colonne sera vide." & vbCrLf
    End If
    If Not ok2025 Then
        avertissements = avertissements & "ATTENTION : fichier 2025 introuvable, la colonne sera vide." & vbCrLf
    End If

    ' ---- Chargement et agrégation ----
    Dim dict2026 As Object, dict2025 As Object
    Set dict2026 = CreateObject("Scripting.Dictionary")
    Set dict2025 = CreateObject("Scripting.Dictionary")

    If ok2026 Then Call ChargerFichierRPS(FICHIER_2026, dict2026)
    If ok2025 Then Call ChargerFichierRPS(FICHIER_2025, dict2025)

    ' ---- Union des clés ----
    Dim toutesLesClés() As String
    toutesLesClés = UnionClés(dict2026, dict2025)

    ' ---- Génération de la feuille comparative ----
    Call GenererFeuille(toutesLesClés, dict2026, dict2025)

    ' ---- Message de fin ----
    Dim nbClesFinal As Long
    nbClesFinal = IIf(UBound(toutesLesClés) >= LBound(toutesLesClés), _
                      UBound(toutesLesClés) - LBound(toutesLesClés) + 1, 0)
    Dim msg As String
    msg = "Tableau comparatif généré avec succès." & vbCrLf & vbCrLf & _
          "  Lignes 2026 agrégées : " & dict2026.Count & vbCrLf & _
          "  Lignes 2025 agrégées : " & dict2025.Count & vbCrLf & _
          "  Clés totales : " & nbClesFinal
    If avertissements <> "" Then
        msg = msg & vbCrLf & vbCrLf & avertissements
    End If
    MsgBox msg, IIf(avertissements <> "", vbExclamation, vbInformation), _
           "Comparatif " & ETAB_NOM & " " & PERIODE & " " & ANNEE_N & " / " & ANNEE_N1

End Sub

' ============================================================
' CHARGEMENT ET AGRÉGATION D'UN FICHIER RPS
' ============================================================
' Agrège les valeurs par clé composite :
'   "Activité|Établissement|Mois|Acte"
' La valeur stockée est la somme numérique de COL_VALEUR.
' ============================================================
Private Sub ChargerFichierRPS(ByVal cheminFichier As String, ByRef dico As Object)

    Dim numFichier As Integer
    numFichier = FreeFile()

    On Error GoTo ErreurLecture
    Open cheminFichier For Input As #numFichier
    On Error GoTo 0

    Dim ligne       As String
    Dim champs()    As String
    Dim cle         As String
    Dim valeur      As Double
    Dim numLigne    As Long
    numLigne = 0

    Do While Not EOF(numFichier)
        Line Input #numFichier, ligne
        numLigne = numLigne + 1

        ' Ignorer les lignes vides et les en-têtes (première ligne ou ligne commençant par #)
        ligne = Trim(ligne)
        If Len(ligne) = 0 Then GoTo SuiteLigne
        If Left(ligne, 1) = "#" Then GoTo SuiteLigne
        If numLigne = 1 And Not IsNumeric(Left(ligne, 1)) Then GoTo SuiteLigne

        champs = Split(ligne, SEP)

        ' Vérification du nombre minimal de colonnes
        If UBound(champs) < (COL_VALEUR - 1) Then GoTo SuiteLigne

        ' Extraction sécurisée de chaque champ
        Dim finess   As String
        Dim activite As String
        Dim mois     As String
        Dim acte     As String
        Dim valStr   As String

        finess   = Trim(champs(COL_FINESS   - 1))
        activite = Trim(champs(COL_ACTIVITE - 1))
        mois     = Trim(champs(COL_MOIS     - 1))
        acte     = Trim(champs(COL_ACTE     - 1))
        valStr   = Trim(champs(COL_VALEUR   - 1))

        ' Remplacer virgule décimale française par point
        valStr = Replace(valStr, ",", ".")
        If Not IsNumeric(valStr) Then valStr = "0"
        valeur = CDbl(valStr)

        ' Construction de la clé composite
        cle = activite & "|" & finess & "|" & mois & "|" & acte

        ' Agrégation (somme)
        If dico.Exists(cle) Then
            dico(cle) = dico(cle) + valeur
        Else
            dico.Add cle, valeur
        End If

SuiteLigne:
    Loop

    Close #numFichier
    Exit Sub

ErreurLecture:
    MsgBox "Erreur lors de la lecture du fichier :" & vbCrLf & cheminFichier & vbCrLf & vbCrLf & _
           "Erreur n°" & Err.Number & " : " & Err.Description, vbCritical, "Erreur de lecture"
    On Error GoTo 0
    If numFichier > 0 Then Close #numFichier

End Sub

' ============================================================
' UNION ET TRI DES CLÉS DES DEUX DICTIONNAIRES
' ============================================================
' Retourne un tableau trié sans doublon combinant les clés
' des deux dictionnaires.
' Tri effectué via tableau (pas de Collection.Add positionnel).
' ============================================================
Private Function UnionClés(ByRef d1 As Object, ByRef d2 As Object) As String()

    ' --- Constitution de l'union ---
    Dim dicoUnion As Object
    Set dicoUnion = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In d1.Keys
        If Not dicoUnion.Exists(CStr(k)) Then dicoUnion.Add CStr(k), 0
    Next k
    For Each k In d2.Keys
        If Not dicoUnion.Exists(CStr(k)) Then dicoUnion.Add CStr(k), 0
    Next k

    Dim n As Long
    n = dicoUnion.Count

    If n = 0 Then
        ' Retourner un tableau vide (UBound < LBound) pour signaler l'absence de clés
        Dim vide() As String
        ReDim vide(0 To -1)
        UnionClés = vide
        Exit Function
    End If

    ' --- Copie dans un tableau ---
    Dim tbl() As String
    ReDim tbl(n - 1)
    Dim i As Long
    i = 0
    For Each k In dicoUnion.Keys
        tbl(i) = CStr(k)
        i = i + 1
    Next k

    ' --- Tri par insertion (stable, pas de Collection positionnel) ---
    Dim j As Long
    Dim tmp As String
    For i = 1 To n - 1
        tmp = tbl(i)
        j = i - 1
        Do While j >= 0 And tbl(j) > tmp
            tbl(j + 1) = tbl(j)
            j = j - 1
        Loop
        tbl(j + 1) = tmp
    Next i

    UnionClés = tbl

End Function

' ============================================================
' GÉNÉRATION DE LA FEUILLE COMPARATIVE
' ============================================================
Private Sub GenererFeuille(ByRef clés() As String, ByRef d2026 As Object, ByRef d2025 As Object)

    ' ---- Créer ou réutiliser la feuille ----
    Dim nomFeuille As String
    nomFeuille = ETAB_NOM & "_" & PERIODE & "_Comp"

    Dim ws As Worksheet
    Dim wsExiste As Boolean
    wsExiste = False
    Dim wsTmp As Worksheet
    For Each wsTmp In ThisWorkbook.Worksheets
        If wsTmp.Name = nomFeuille Then
            wsExiste = True
            Set ws = wsTmp
            Exit For
        End If
    Next wsTmp

    If wsExiste Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nomFeuille

    ' ---- Désactiver le rafraîchissement pendant la construction ----
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo FinGeneration

    ' ========================================================
    ' LIGNE 1 : Grand titre
    ' ========================================================
    Dim ligneActuelle As Long
    ligneActuelle = 1

    With ws.Range("A1:J1")
        .Merge
        .Value = ETAB_NOM & " — Comparatif " & ANNEE_N & " / " & ANNEE_N1 & " — Période " & PERIODE
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COULEUR_TEXTE_BLANC
        .Interior.Color = COULEUR_TITRE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With

    ' ========================================================
    ' LIGNE 2 : Sous-titre (informations fichiers)
    ' ========================================================
    ligneActuelle = 2
    With ws.Range("A2:J2")
        .Merge
        .Value = "Fichier " & ANNEE_N & " : " & FICHIER_2026 & "   |   " & _
                 "Fichier " & ANNEE_N1 & " : " & FICHIER_2025
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(230, 235, 245)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 18
        .WrapText = False
    End With

    ' ========================================================
    ' LIGNE 3 : Vide (séparateur)
    ' ========================================================
    ligneActuelle = 3
    ws.Rows(3).RowHeight = 6

    ' ========================================================
    ' LIGNE 4 : En-têtes de colonnes
    ' ========================================================
    ligneActuelle = 4
    Dim colonnes(9) As String
    colonnes(0) = "Activité"
    colonnes(1) = "Établissement"
    colonnes(2) = "Mois"
    colonnes(3) = "Code Acte"
    colonnes(4) = ANNEE_N & " (N)"
    colonnes(5) = ANNEE_N1 & " (N-1)"
    colonnes(6) = "Écart (N - N-1)"
    colonnes(7) = "Variation (%)"
    colonnes(8) = "Présent " & ANNEE_N
    colonnes(9) = "Présent " & ANNEE_N1

    Dim c As Integer
    For c = 0 To 9
        With ws.Cells(ligneActuelle, c + 1)
            .Value = colonnes(c)
            .Font.Bold = True
            .Font.Color = COULEUR_TEXTE_BLANC
            .Interior.Color = COULEUR_EN_TETE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next c
    ws.Rows(ligneActuelle).RowHeight = 36

    ' ========================================================
    ' LIGNES DE DONNÉES
    ' ========================================================
    Dim totalN   As Double
    Dim totalN1  As Double
    Dim nbCles   As Long
    totalN  = 0
    totalN1 = 0

    ' Vérification que le tableau de clés n'est pas vide (UBound < LBound = 0)
    If UBound(clés) < LBound(clés) Then GoTo LigneTotal

    nbCles = UBound(clés) + 1

    Dim idx As Long
    For idx = 0 To nbCles - 1
        Dim cle As String
        cle = clés(idx)
        If Len(cle) = 0 Then GoTo SuiteData

        ' Décomposer la clé
        Dim parties() As String
        parties = Split(cle, "|")
        If UBound(parties) < 3 Then GoTo SuiteData

        Dim sActivite As String, sFiness As String, sMois As String, sActe As String
        sActivite = parties(0)
        sFiness   = parties(1)
        sMois     = parties(2)
        sActe     = parties(3)

        ' Valeurs
        Dim vN  As Double, vN1 As Double
        vN  = IIf(d2026.Exists(cle), CDbl(d2026(cle)), 0)
        vN1 = IIf(d2025.Exists(cle), CDbl(d2025(cle)), 0)

        Dim ecart     As Double
        Dim variation As Variant
        ecart = vN - vN1

        If vN1 <> 0 Then
            variation = (ecart / Abs(vN1)) * 100
        ElseIf vN <> 0 Then
            variation = "N/A (nouveau)"
        Else
            variation = 0
        End If

        totalN  = totalN  + vN
        totalN1 = totalN1 + vN1

        ligneActuelle = ligneActuelle + 1
        Dim r As Long
        r = ligneActuelle

        ' Couleur alternée des lignes
        Dim coulLigne As Long
        coulLigne = IIf((idx Mod 2) = 0, COULEUR_LIGNE_IMP, COULEUR_LIGNE_PAIR)

        ' --- Remplissage des cellules ---
        With ws.Cells(r, 1)
            .Value = sActivite
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlCenter
        End With
        With ws.Cells(r, 2)
            .Value = sFiness
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlCenter
        End With
        With ws.Cells(r, 3)
            If IsNumeric(sMois) Then
                .Value = CInt(sMois)
                .NumberFormat = "00"
            Else
                .Value = sMois
            End If
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlCenter
        End With
        With ws.Cells(r, 4)
            .Value = sActe
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlLeft
        End With
        With ws.Cells(r, 5)
            .Value = vN
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00"
        End With
        With ws.Cells(r, 6)
            .Value = vN1
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00"
        End With
        With ws.Cells(r, 7)
            .Value = ecart
            .HorizontalAlignment = xlRight
            .NumberFormat = "#,##0.00"
            If ecart > 0 Then
                .Interior.Color = COULEUR_VARIATION_POS
                .Font.Color = RGB(0, 100, 0)
            ElseIf ecart < 0 Then
                .Interior.Color = COULEUR_VARIATION_NEG
                .Font.Color = RGB(150, 0, 0)
            Else
                .Interior.Color = coulLigne
                .Font.Color = RGB(0, 0, 0)
            End If
        End With
        With ws.Cells(r, 8)
            If IsNumeric(variation) Then
                .Value = CDbl(variation) / 100
                .NumberFormat = "0.0%"
                If CDbl(variation) > 0 Then
                    .Interior.Color = COULEUR_VARIATION_POS
                    .Font.Color = RGB(0, 100, 0)
                ElseIf CDbl(variation) < 0 Then
                    .Interior.Color = COULEUR_VARIATION_NEG
                    .Font.Color = RGB(150, 0, 0)
                Else
                    .Interior.Color = coulLigne
                    .Font.Color = RGB(0, 0, 0)
                End If
            Else
                .Value = CStr(variation)
                .Interior.Color = coulLigne
                .HorizontalAlignment = xlCenter
                .Font.Italic = True
            End If
        End With
        With ws.Cells(r, 9)
            .Value = IIf(d2026.Exists(cle), "Oui", "Non")
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlCenter
            If d2026.Exists(cle) Then
                .Font.Color = RGB(0, 128, 0)
                .Font.Bold = True
            Else
                .Font.Color = RGB(180, 0, 0)
                .Font.Italic = True
            End If
        End With
        With ws.Cells(r, 10)
            .Value = IIf(d2025.Exists(cle), "Oui", "Non")
            .Interior.Color = coulLigne
            .HorizontalAlignment = xlCenter
            If d2025.Exists(cle) Then
                .Font.Color = RGB(0, 128, 0)
                .Font.Bold = True
            Else
                .Font.Color = RGB(180, 0, 0)
                .Font.Italic = True
            End If
        End With

SuiteData:
    Next idx

    ' ========================================================
    ' LIGNE TOTAL
    ' ========================================================
LigneTotal:
    ligneActuelle = ligneActuelle + 1
    Dim rTot As Long
    rTot = ligneActuelle

    Dim ecartTotal As Double
    ecartTotal = totalN - totalN1

    Dim varTotalPct As Variant
    If totalN1 <> 0 Then
        varTotalPct = (ecartTotal / Abs(totalN1)) * 100
    ElseIf totalN <> 0 Then
        varTotalPct = "N/A"
    Else
        varTotalPct = 0
    End If

    With ws.Range(ws.Cells(rTot, 1), ws.Cells(rTot, 4))
        .Merge
        .Value = "TOTAL"
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = COULEUR_TOTAL
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Cells(rTot, 5)
        .Value = totalN
        .Font.Bold = True
        .Interior.Color = COULEUR_TOTAL
        .NumberFormat = "#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(rTot, 6)
        .Value = totalN1
        .Font.Bold = True
        .Interior.Color = COULEUR_TOTAL
        .NumberFormat = "#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(rTot, 7)
        .Value = ecartTotal
        .Font.Bold = True
        .Interior.Color = COULEUR_TOTAL
        .NumberFormat = "#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(rTot, 8)
        If IsNumeric(varTotalPct) Then
            .Value = CDbl(varTotalPct) / 100
            .NumberFormat = "0.0%"
        Else
            .Value = CStr(varTotalPct)
            .HorizontalAlignment = xlCenter
        End If
        .Font.Bold = True
        .Interior.Color = COULEUR_TOTAL
    End With
    With ws.Range(ws.Cells(rTot, 9), ws.Cells(rTot, 10))
        .Interior.Color = COULEUR_TOTAL
    End With

    ws.Rows(rTot).RowHeight = 22

    ' ========================================================
    ' MISE EN PAGE GÉNÉRALE
    ' ========================================================

    ' Largeurs de colonnes
    ws.Columns(1).ColumnWidth  = 12   ' Activité
    ws.Columns(2).ColumnWidth  = 18   ' Établissement
    ws.Columns(3).ColumnWidth  = 8    ' Mois
    ws.Columns(4).ColumnWidth  = 22   ' Code Acte
    ws.Columns(5).ColumnWidth  = 14   ' N
    ws.Columns(6).ColumnWidth  = 14   ' N-1
    ws.Columns(7).ColumnWidth  = 16   ' Écart
    ws.Columns(8).ColumnWidth  = 14   ' Variation %
    ws.Columns(9).ColumnWidth  = 12   ' Présent N
    ws.Columns(10).ColumnWidth = 12   ' Présent N-1

    ' Bordures sur la zone de données
    Dim derniereColonne As Integer
    derniereColonne = 10
    Dim plage As Range
    Set plage = ws.Range(ws.Cells(4, 1), ws.Cells(rTot, derniereColonne))
    AppliquerBordures plage

    ' Figer les volets à partir de la ligne de données (ligne 5)
    ws.Activate
    ws.Cells(5, 1).Select
    ActiveWindow.FreezePanes = False
    ws.Cells(5, 5).Select
    ActiveWindow.FreezePanes = True

    ' Zoom
    ActiveWindow.Zoom = 85

    ' Retour en A1
    ws.Cells(1, 1).Select

    GoTo FinOK

FinGeneration:
    MsgBox "Erreur lors de la génération de la feuille :" & vbCrLf & _
           "Erreur n°" & Err.Number & " : " & Err.Description, vbCritical, "Erreur"

FinOK:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' ============================================================
' APPLIQUE DES BORDURES FINES SUR UNE PLAGE
' ============================================================
Private Sub AppliquerBordures(ByRef plage As Range)
    Dim bordures As Variant
    Dim bd As Variant
    bordures = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                     xlInsideVertical, xlInsideHorizontal)
    For Each bd In bordures
        With plage.Borders(bd)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 200)
        End With
    Next bd
    ' Bordure extérieure plus épaisse
    With plage.Borders(xlEdgeLeft)
        .Weight = xlMedium
        .Color = RGB(80, 80, 120)
    End With
    With plage.Borders(xlEdgeRight)
        .Weight = xlMedium
        .Color = RGB(80, 80, 120)
    End With
    With plage.Borders(xlEdgeTop)
        .Weight = xlMedium
        .Color = RGB(80, 80, 120)
    End With
    With plage.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .Color = RGB(80, 80, 120)
    End With
End Sub
