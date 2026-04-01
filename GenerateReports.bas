Attribute VB_Name = "GenerateReports"
Option Explicit

' ============================================================
' MODULE : GenerateReports
' ============================================================
' BLOC CONFIGURATION UTILISATEUR
' ============================================================
' Pour ajouter un établissement :
'   Décommenter la ligne correspondante (supprimer le ' devant)
'   ou ajouter une nouvelle ligne :
'       Call AjouterEtab(listeEtab, idx, "CODE")
'
' Pour retirer un établissement :
'   Commenter la ligne (ajouter un ' devant) ou la supprimer.
'
' Vous pouvez comparer 2, 3, 4 établissements ou plus,
' selon les lignes actives dans la liste.
' ============================================================

' --- Chemins de base ---
Private Const CHEMIN_N   As String = "D:\ASIA\asia-tim\rimp\"  ' Chemin pour l'année N (2026)
Private Const CHEMIN_N1  As String = "D:\ASIA\rimp_2025\"      ' Chemin alternatif pour l'année N-1

' --- Période ---
Private Const MOIS       As String = "M2"   ' Mois à traiter (ex : M2 = mois 2)
Private Const ANNEE_N    As Integer = 2026  ' Année courante (N)
Private Const ANNEE_N1   As Integer = 2025  ' Année précédente (N-1)

' ============================================================
' PROCÉDURE PRINCIPALE
' ============================================================
Sub GenerateReports()

    ' ============================================================
    ' LISTE DES ÉTABLISSEMENTS À TRAITER
    ' - Pour inclure un établissement : ligne décommentée (active)
    ' - Pour l'exclure : commenter la ligne (ajouter ' devant)
    ' ============================================================
    Dim listeEtab() As String
    Dim idx As Integer
    idx = 0

    ' --- LISTE À ÉDITER ---
    Call AjouterEtab(listeEtab, idx, "DC")
    Call AjouterEtab(listeEtab, idx, "RC")
    Call AjouterEtab(listeEtab, idx, "EC")
    Call AjouterEtab(listeEtab, idx, "EU")
    Call AjouterEtab(listeEtab, idx, "GD")
    Call AjouterEtab(listeEtab, idx, "MA")
    Call AjouterEtab(listeEtab, idx, "VAL")
    Call AjouterEtab(listeEtab, idx, "EB")
    Call AjouterEtab(listeEtab, idx, "VA")
    Call AjouterEtab(listeEtab, idx, "Montaury")
    Call AjouterEtab(listeEtab, idx, "ET")
    ' ----------------------
    ' Pour ajouter un nouvel établissement, ajoutez une ligne ici :
    ' Call AjouterEtab(listeEtab, idx, "NOUVEAU_CODE")

    ' Vérifie qu'au moins un établissement est configuré
    If idx = 0 Then
        MsgBox "Aucun établissement configuré dans la liste." & vbCrLf & _
               "Veuillez décommenter au moins une ligne dans la liste.", _
               vbExclamation, "Configuration vide"
        Exit Sub
    End If

    ReDim Preserve listeEtab(idx - 1)

    ' ============================================================
    ' TRAITEMENT DES ÉTABLISSEMENTS
    ' ============================================================
    Dim codeEtab  As String
    Dim pathN     As String
    Dim pathN1    As String
    Dim msgOK     As String
    Dim msgManquants As String
    Dim compteur  As Integer

    compteur     = 0
    msgOK        = ""
    msgManquants = ""

    Dim i As Integer
    For i = 0 To UBound(listeEtab)
        codeEtab = listeEtab(i)

        pathN  = TrouverFichier(codeEtab, ANNEE_N)
        pathN1 = TrouverFichier(codeEtab, ANNEE_N1)

        If pathN <> "" And pathN1 <> "" Then
            ' Les deux fichiers existent : traitement
            msgOK = msgOK & "  " & Chr(10003) & " " & codeEtab & vbCrLf & _
                    "      N   : " & pathN & vbCrLf & _
                    "      N-1 : " & pathN1 & vbCrLf
            Call TraiterEtablissement(codeEtab, pathN, pathN1)
            compteur = compteur + 1
        Else
            ' Au moins un fichier manquant : signalement
            Dim detailManquant As String
            detailManquant = ""
            If pathN = "" Then
                detailManquant = detailManquant & _
                    "      N   (" & ANNEE_N & ") : INTROUVABLE" & vbCrLf
            Else
                detailManquant = detailManquant & _
                    "      N   (" & ANNEE_N & ") : " & pathN & vbCrLf
            End If
            If pathN1 = "" Then
                detailManquant = detailManquant & _
                    "      N-1 (" & ANNEE_N1 & ") : INTROUVABLE" & vbCrLf
            Else
                detailManquant = detailManquant & _
                    "      N-1 (" & ANNEE_N1 & ") : " & pathN1 & vbCrLf
            End If
            msgManquants = msgManquants & "  ! " & codeEtab & vbCrLf & detailManquant
        End If
    Next i

    ' ============================================================
    ' MESSAGE DE SYNTHÈSE FINAL
    ' ============================================================
    Dim msgFinal As String
    Dim icone    As Integer

    msgFinal = "=== RAPPORT DE TRAITEMENT ===" & vbCrLf & vbCrLf

    If msgOK <> "" Then
        msgFinal = msgFinal & "ETABLISSEMENTS TRAITES (" & compteur & ") :" & vbCrLf & msgOK & vbCrLf
    End If

    If msgManquants <> "" Then
        msgFinal = msgFinal & "FICHIERS MANQUANTS :" & vbCrLf & msgManquants & vbCrLf
    End If

    If compteur = 0 And msgManquants = "" Then
        msgFinal = msgFinal & "Aucun établissement à traiter."
        icone = vbInformation
    ElseIf msgManquants <> "" Then
        msgFinal = msgFinal & "(" & compteur & " établissement(s) traité(s), " & _
                   (UBound(listeEtab) + 1 - compteur) & " avec fichier(s) manquant(s).)"
        icone = vbExclamation
    Else
        msgFinal = msgFinal & Chr(10003) & " " & compteur & " établissement(s) traité(s) avec succès."
        icone = vbInformation
    End If

    MsgBox msgFinal, icone, "Generation des rapports " & MOIS & " " & ANNEE_N

End Sub

' ============================================================
' Ajoute un code établissement dans le tableau dynamique
' ============================================================
Private Sub AjouterEtab(ByRef liste() As String, ByRef idx As Integer, ByVal codeEtab As String)
    If idx = 0 Then
        ReDim liste(0)
    ElseIf idx > UBound(liste) Then
        ReDim Preserve liste(idx)
    End If
    liste(idx) = Trim(codeEtab)
    idx = idx + 1
End Sub

' ============================================================
' Retourne le préfixe utilisé dans le nom de fichier pour un
' établissement donné. Pour la plupart, c'est le code lui-même ;
' pour Montaury, le fichier porte le préfixe "MA".
' ============================================================
Private Function ObtenirPrefixeFichier(ByVal codeEtab As String) As String
    Select Case codeEtab
        Case "Montaury": ObtenirPrefixeFichier = "MA"
        Case Else:       ObtenirPrefixeFichier = codeEtab
    End Select
End Function

' ============================================================
' Recherche automatique du fichier d'un établissement pour
' une année donnée. Essaie d'abord le chemin principal (N),
' puis le chemin alternatif (N-1). Teste aussi avec/sans .txt.
' Retourne le chemin complet si trouvé, sinon "".
' ============================================================
Private Function TrouverFichier(ByVal codeEtab As String, ByVal annee As Integer) As String

    Dim nomBase     As String
    Dim cheminBase  As String
    Dim chemin      As String

    ' Nom du fichier sans extension (le préfixe peut différer du code dossier, ex. Montaury → MA)
    nomBase = ObtenirPrefixeFichier(codeEtab) & "-" & annee & MOIS & "-rps"

    ' Chemins à tester dans l'ordre de priorité
    Dim chemins(1) As String
    chemins(0) = CHEMIN_N  & codeEtab & "\" & MOIS & "\" & nomBase
    chemins(1) = CHEMIN_N1 & codeEtab & "\" & MOIS & "\" & nomBase

    Dim j As Integer
    For j = 0 To UBound(chemins)
        ' Test sans extension
        chemin = chemins(j)
        If Dir(chemin) <> "" Then
            TrouverFichier = chemin
            Exit Function
        End If
        ' Test avec extension .txt
        chemin = chemins(j) & ".txt"
        If Dir(chemin) <> "" Then
            TrouverFichier = chemin
            Exit Function
        End If
    Next j

    TrouverFichier = ""

End Function

' ============================================================
' TraiterEtablissement : à compléter avec la logique existante
' (lecture du fichier TXT, alimentation de la feuille Excel,
'  calculs, mise en forme, etc.)
' NE PAS MODIFIER la signature de cette procédure.
' ============================================================
'
' Private Sub TraiterEtablissement(ByVal codeEtab As String, _
'                                   ByVal pathN As String, _
'                                   ByVal pathN1 As String)
'     ' ... votre logique existante ici ...
' End Sub
