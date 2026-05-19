Attribute VB_Name = "GenerateReports"
'
' Module  : GenerateReports
' Auteur  : (à renseigner)
' Date    : 2026
' Objet   : Génération des rapports comparatifs N / N-1 pour plusieurs établissements.
'
' ============================================================
'  POINT DE CONFIGURATION UTILISATEUR
' ============================================================
'  Modifiez uniquement la fonction GetEtablissementsSelectionnes()
'  pour choisir quels établissements doivent être traités.
'  Mettez True pour inclure un établissement, False pour l'exclure.
'
'  Pour ajouter un nouvel établissement, ajoutez une ligne dans
'  GetEtablissementsSelectionnes() ET une entrée correspondante
'  dans GetCheminFichier().
' ============================================================

Option Explicit

' ============================================================
'  CONSTANTES DE CHEMINS (racines)
' ============================================================
Private Const RACINE_2026    As String = "D:\ASIA\asia-tim\rimp\"
Private Const RACINE_2025    As String = "D:\ASIA\rimp_2025\"
Private Const SUFFIXE        As String = "\M2\"
Private Const ANNEE_N        As String = "2026"
Private Const ANNEE_N1       As String = "2025"
Private Const EXT_FICHIER    As String = "M2-rps.txt"
Private Const CHAR_CROIX     As String = "❌"   ' Chr(10060) – fichier manquant
Private Const CHAR_COCHE     As String = "✔"    ' Chr(10004) – succès

' ============================================================
'  POINT DE CONFIGURATION UTILISATEUR
'  Listez les codes établissement à traiter, séparés par des virgules.
'  Supprimez ou commentez un code pour l'exclure du traitement.
'  Aucun ajustement de borne de tableau requis.
' ============================================================
Private Function GetEtablissementsSelectionnes() As String()
    Const LISTE As String = "DC,GD,MA"   ' <-- Modifiez ici : ex. "DC" ou "DC,GD" ou "DC,GD,MA"
    '                                         MA = Montaury : cas particulier géré automatiquement
    GetEtablissementsSelectionnes = Split(LISTE, ",")
End Function

' ============================================================
'  Retourne le chemin complet d'un fichier pour un établissement
'  et une année données.
'
'  Cas particulier Montaury (code "MA") :
'    - 2026 : préfixe "MA"  dans RACINE_2026\MA\
'    - 2025 : préfixe "Montaury" dans RACINE_2025\Montaury\
' ============================================================
Private Function GetCheminFichier(ByVal codeEtab As String, ByVal annee As String) As String
    Dim dossierRacine As String
    Dim dossierEtab   As String
    Dim prefixeFichier As String

    Select Case annee
        Case ANNEE_N   ' 2026
            dossierRacine  = RACINE_2026
            If UCase(codeEtab) = "MA" Then
                dossierEtab    = "MA"
                prefixeFichier = "MA"
            Else
                dossierEtab    = codeEtab
                prefixeFichier = codeEtab
            End If

        Case ANNEE_N1  ' 2025
            dossierRacine  = RACINE_2025
            If UCase(codeEtab) = "MA" Then
                dossierEtab    = "Montaury"
                prefixeFichier = "Montaury"
            Else
                dossierEtab    = codeEtab
                prefixeFichier = codeEtab
            End If

        Case Else
            GetCheminFichier = ""
            Exit Function
    End Select

    GetCheminFichier = dossierRacine & dossierEtab & SUFFIXE & _
                       prefixeFichier & "-" & annee & EXT_FICHIER
End Function

' ============================================================
'  Signature de TraiterEtablissement
'  À COMPLÉTER par l'utilisateur avec la logique métier.
'
'  Paramètres :
'    codeEtab  : code court de l'établissement (ex. "DC", "MA")
'    cheminN   : chemin complet du fichier 2026
'    cheminN1  : chemin complet du fichier 2025
' ============================================================
Public Sub TraiterEtablissement(ByVal codeEtab As String, _
                                 ByVal cheminN  As String, _
                                 ByVal cheminN1 As String)
    ' TODO : insérer ici la logique de traitement pour un établissement
    '        (lecture des fichiers, calculs, écriture du rapport, etc.)
    MsgBox "Traitement de " & codeEtab & " : non encore implémenté.", vbInformation
End Sub

' ============================================================
'  Procédure principale
' ============================================================
Public Sub GenerateReports()

    '--- Récupération de la liste des établissements sélectionnés ---
    Dim listeEtabs() As String
    listeEtabs = GetEtablissementsSelectionnes()

    '--- Robustesse : tableau vide ou liste vide ---
    Dim nbEtabs As Long
    On Error Resume Next
    nbEtabs = UBound(listeEtabs) - LBound(listeEtabs) + 1
    On Error GoTo 0

    If nbEtabs <= 0 Then
        MsgBox "Aucun établissement sélectionné." & vbCrLf & _
               "Veuillez configurer GetEtablissementsSelectionnes().", _
               vbExclamation, "GenerateReports"
        Exit Sub
    End If

    '--- Détection des fichiers manquants ---
    Dim manquants()    As String
    Dim nbManquants    As Long
    ReDim manquants(0 To nbEtabs * 2 - 1)   ' au plus 2 fichiers manquants par étab.
    nbManquants = 0

    '--- Listes de traitement effectif ---
    Dim codesTraites() As String
    Dim cheminsN()     As String
    Dim cheminsN1()    As String
    Dim nbTraites      As Long
    ReDim codesTraites(0 To nbEtabs - 1)
    ReDim cheminsN(0 To nbEtabs - 1)
    ReDim cheminsN1(0 To nbEtabs - 1)
    nbTraites = 0

    Dim i         As Long
    Dim codeEtab  As String
    Dim pathN     As String
    Dim pathN1    As String
    Dim okN       As Boolean
    Dim okN1      As Boolean

    For i = LBound(listeEtabs) To UBound(listeEtabs)
        codeEtab = Trim(listeEtabs(i))
        If Len(codeEtab) = 0 Then GoTo SuivantEtab    ' robustesse : entrée vide

        pathN  = GetCheminFichier(codeEtab, ANNEE_N)
        pathN1 = GetCheminFichier(codeEtab, ANNEE_N1)

        okN  = (Len(pathN) > 0)  And (Dir(pathN) <> "")
        okN1 = (Len(pathN1) > 0) And (Dir(pathN1) <> "")

        If Not okN Then
            manquants(nbManquants) = codeEtab & " [" & ANNEE_N & "] : " & pathN
            nbManquants = nbManquants + 1
        End If
        If Not okN1 Then
            manquants(nbManquants) = codeEtab & " [" & ANNEE_N1 & "] : " & pathN1
            nbManquants = nbManquants + 1
        End If

        If okN And okN1 Then
            codesTraites(nbTraites) = codeEtab
            cheminsN(nbTraites)     = pathN
            cheminsN1(nbTraites)    = pathN1
            nbTraites = nbTraites + 1
        End If

SuivantEtab:
    Next i

    '--- Synthèse des fichiers manquants ---
    If nbManquants > 0 Then
        Dim synthese As String
        synthese = "=== FICHIERS MANQUANTS (" & nbManquants & ") ===" & vbCrLf & vbCrLf
        Dim j As Long
        For j = 0 To nbManquants - 1
            synthese = synthese & CHAR_CROIX & " " & manquants(j) & vbCrLf
        Next j
        synthese = synthese & vbCrLf & _
                   "Les établissements avec fichiers manquants ne seront PAS traités."
        MsgBox synthese, vbExclamation, "GenerateReports – Fichiers manquants"
    End If

    '--- Robustesse : aucun établissement traitable ---
    If nbTraites = 0 Then
        MsgBox "Aucun établissement ne peut être traité (fichiers introuvables)." & vbCrLf & _
               "Vérifiez les chemins et relancez.", vbCritical, "GenerateReports"
        Exit Sub
    End If

    '--- Traitement effectif ---
    Dim compteur As Long
    compteur = 0

    For i = 0 To nbTraites - 1
        Call TraiterEtablissement(codesTraites(i), cheminsN(i), cheminsN1(i))
        compteur = compteur + 1
    Next i

    '--- Bilan final ---
    Dim bilan As String
    bilan = "=== BILAN GÉNÉRATION ===" & vbCrLf & vbCrLf & _
            CHAR_COCHE & " " & compteur & " rapport(s) généré(s) sur " & nbEtabs & " demandé(s)."
    If nbManquants > 0 Then
        bilan = bilan & vbCrLf & "⚠ " & (nbEtabs - nbTraites) & " établissement(s) ignoré(s) (fichiers manquants)."
    End If
    MsgBox bilan, vbInformation, "GenerateReports – Terminé"

End Sub
