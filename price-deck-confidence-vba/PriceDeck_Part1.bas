Attribute VB_Name = "Module1"
Dim Fichier As String
Dim Wb1 As Workbook
Dim Wb2 As Workbook
Dim type_version As Long

Sub code2()
    
       
Dim resultat As String
Dim tab_onglets As Variant
Dim tab_noms_plages As Variant
Dim plage As String
Dim lettre As String
Dim onglet_actu As String
Dim vecteur_noms_plages As String
Dim plage_actu As String
Dim arg1 As Double
Dim arg2 As Double
Dim arg3 As Double
Dim arg4 As Double
Dim nom_fichier As String

UserForm_demo.CommandButton1.Visible = True
Application.ScreenUpdating = True
MSG1 = MsgBox("Is the Input (Price End) updated by yourself?", vbYesNo, "Price End Update")

UserForm_demo.CommandButton1.Caption = "...Updating the prices..."
progression = 0
UserForm_demo.Label_barre.Caption = progression & "%"

If MSG1 = vbNo Then
    'Valeurs actuels
    AECO_GAS = Application.WorksheetFunction.Max(Wb1.Sheets("ConfidenceLevels").Range("Fwd_AKO"))
    US_Gas = Application.WorksheetFunction.Max(Wb1.Sheets("ConfidenceLevels").Range("Fwd_NG"))
    Crude_Oil = Application.WorksheetFunction.Max(Wb1.Sheets("ConfidenceLevels").Range("Fwd_CL"))
    Brent_Oil = Application.WorksheetFunction.Max(Wb1.Sheets("ConfidenceLevels").Range("Fwd_BL"))
    UK_Gas = Application.WorksheetFunction.Max(Wb1.Sheets("ConfidenceLevels").Range("Fwd_NBP"))
    
    Wb2.Sheets("Input_VBA").Range("C2").Value = Round(2 * Crude_Oil, 2)
    Wb2.Sheets("Input_VBA").Range("C3").Value = Round(2 * Crude_Oil, 2)
    Wb2.Sheets("Input_VBA").Range("C4").Value = Round(2 * US_Gas, 2)
    Wb2.Sheets("Input_VBA").Range("C5").Value = Round(2 * US_Gas, 2)
    Wb2.Sheets("Input_VBA").Range("C6").Value = Round(2 * AECO_GAS, 2)
    Wb2.Sheets("Input_VBA").Range("C7").Value = Round(2 * AECO_GAS, 2)
    Wb2.Sheets("Input_VBA").Range("C8").Value = Round(2 * Brent_Oil, 2)
    Wb2.Sheets("Input_VBA").Range("C9").Value = Round(2 * Brent_Oil, 2)
    
    Wb2.Sheets("Input_VBA").Range("C10").Value = Round(2 * UK_Gas, 2)
    Wb2.Sheets("Input_VBA").Range("C11").Value = Round(2 * UK_Gas, 2)
End If


progression = progression + 10
UserForm_demo.Label_barre.Caption = progression & "%"
DoEvents


If UserForm_demo.CheckBox_Crude.Value = True Then

        'Crude Base Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B2").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C2").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D2").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E2").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "Crude Base Case", "BC_CL_Asia", "ConfLev_BC_CL_Asia")
        'Efface ancienne valeurs
        Sheets("Crude").Select
        Range("D3:AF65000").Select
        Selection.ClearContents
        
        Call copie_colle_valeur_tableau_commun("Crude Base Case", "Crude", 13)
        Wb2.Sheets("Crude").Cells(2, 13).Value = "Crude Base Case " & arg4 & "%"
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
        
        'Crude Sensitivity Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B3").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C3").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D3").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E3").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "Crude Sensitivity Case", "SC_CL_Asia", "ConfLev_SC_CL_Asia")
        Call copie_colle_valeur_tableau_commun("Crude Sensitivity Case", "Crude", 14)
        Wb2.Sheets("Crude").Cells(2, 14).Value = "Crude Sensitivity Case " & arg4 & "%"
        
        Call copie_colle_concurence_price("Crude", 2, 4)
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
End If

If UserForm_demo.CheckBox_USGAS.Value = True Then
        'US GAS Base Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B4").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C4").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D4").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E4").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "US GAS Base Case ", "BC_NG_Asia", "ConfLev_BC_NG_Asia")
        'Efface ancienne valeurs
        Sheets("US GAS").Select
        Range("D3:AF65000").Select
        Selection.ClearContents
        
        Call copie_colle_valeur_tableau_commun("US GAS Base Case ", "US GAS", 13)
        Wb2.Sheets("US GAS").Cells(2, 13).Value = "US GAS Base Case " & arg4 & "%"
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
        
        'US GAS Sensitivity Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B5").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C5").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D5").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E5").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "US GAS Sensitivity Case", "SC_NG_Asia", "ConfLev_SC_NG_Asia")
        Call copie_colle_valeur_tableau_commun("US GAS Sensitivity Case", "US GAS", 14)
        Wb2.Sheets("US GAS").Cells(2, 14).Value = "US GAS Sensitivity Case" & arg4 & "%"
        
        Call copie_colle_concurence_price("US GAS", 14, 4)
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
End If

If UserForm_demo.CheckBox_AECOGAS.Value = True Then
        'AECO GAS Base Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B6").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C6").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D6").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E6").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "AECO GAS Base Case", "BC_AKO_Asia", "ConfLev_BC_AKO_Asia")
        'Efface ancienne valeurs
        Sheets("AECO GAS").Select
        Range("D3:AF65000").Select
        Selection.ClearContents
        
        Call copie_colle_valeur_tableau_commun("AECO GAS Base Case", "AECO GAS", 13)
        Wb2.Sheets("AECO GAS").Cells(2, 13).Value = "AECO GAS Base Case " & arg4 & "%"
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
        
        'AECO GAS Sensitivity Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B7").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C7").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D7").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E7").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "AECO GAS Sensitivity Case", "SC_AKO_Asia", "ConfLev_SC_AKO_Asia")
        Call copie_colle_valeur_tableau_commun("AECO GAS Sensitivity Case", "AECO GAS", 14)
        Wb2.Sheets("AECO GAS").Cells(2, 14).Value = "AECO GAS Sensitivity Case " & arg4 & "%"
        
        Call copie_colle_concurence_price("AECO GAS", 18, 4)
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
End If

If UserForm_demo.CheckBox_Brent.Value = True Then
        'Brent Base Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B8").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C8").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D8").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E8").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "Brent Base Case", "BC_BL_Asia", "ConfLev_BC_BL_Asia")
        
        'Efface ancienne valeurs
        Sheets("Brent").Select
        Range("D3:AF65000").Select
        Selection.ClearContents
        
        Call copie_colle_valeur_tableau_commun("Brent Base Case", "Brent", 13)
        Wb2.Sheets("Brent").Cells(2, 13).Value = "Brent Base Case " & arg4 & "%"
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
        
        'Brent Sensitivity Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B9").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C8").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D9").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E9").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "Brent Sensitivity Case", "SC_BL_Asia", "ConfLev_SC_BL_Asia")
        Call copie_colle_valeur_tableau_commun("Brent Sensitivity Case", "Brent", 14)
        Wb2.Sheets("Brent").Cells(2, 14).Value = "Brent Sensitivity Case " & arg4 & "%"
        
        Call copie_colle_concurence_price("Brent", 6, 4)
        
        progression = progression + 10
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
End If

If UserForm_demo.CheckBox_UKGAS.Value = True Then

        'UK Gas Base Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B10").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C10").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D10").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E10").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "UK Gas Base Case", "BC_NBP_Asia", "ConfLev_BC_NBP_Asia")
        'Efface ancienne valeurs
        Sheets("NBP").Select
        Range("D3:AF65000").Select
        Selection.ClearContents
        
        
        Call copie_colle_valeur_tableau_commun("UK Gas Base Case", "NBP", 13)
        Wb2.Sheets("NBP").Cells(2, 13).Value = "UK Gas Base Case " & arg4 & "%"
        
        progression = progression + 5
        UserForm_demo.Label_barre.Caption = progression & "%"
        DoEvents
        
        'UK Gas Sensitivity Case
        arg1 = CDbl(Wb2.Sheets("Input_VBA").Range("B11").Value)
        arg2 = CDbl(Wb2.Sheets("Input_VBA").Range("C11").Value)
        arg3 = CDbl(Wb2.Sheets("Input_VBA").Range("D11").Value)
        arg4 = CDbl(Wb2.Sheets("Input_VBA").Range("E11").Value)
        Call maj_tableur_avec_vals_new(arg1, arg2, arg3, arg4, "UK Gas Sensitivity Case", "SC_NBP_Asia", "ConfLev_SC_NBP_Asia")
        Call copie_colle_valeur_tableau_commun("UK Gas Sensitivity Case", "NBP", 14)
        Wb2.Sheets("NBP").Cells(2, 14).Value = "UK Gas Sensitivity Case " & arg4 & "%"
        
        Call copie_colle_concurence_price("NBP", 10, 4)

End If

Call main_get_pourcentage_from_porposition_SG

progression = 100
UserForm_demo.Label_barre.Caption = progression & "%"
DoEvents

    If UserForm_demo.CheckBox_Crude.Value = True Then Sheets("Crude Graph").Activate
    If UserForm_demo.CheckBox_Brent.Value = True Then Sheets("Brent Graph").Activate
    If UserForm_demo.CheckBox_USGAS.Value = True Then Sheets("US GAS Graph").Activate
    If UserForm_demo.CheckBox_AECOGAS.Value = True Then Sheets("AECO GAS Graph").Activate
    If UserForm_demo.CheckBox_UKGAS.Value = True Then Sheets("NBP Graph").Activate

    
'Sauvegarde résultat
ActiveWorkbook.Save

'Save le même fichier sans formules Bloomberg
Sheets("Crude").Activate
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Sheets("US GAS").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Sheets("AECO GAS").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Sheets("Brent").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Sheets("NBP").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Call copie_colle_year_now(Sheets("Resultats").Range("A4").Value, Sheets("Resultats").Range("A2").Value, Sheets("Resultats").Range("B4").Value, Sheets("Resultats").Range("C4").Value)
Call copie_colle_year_now(Sheets("Resultats").Range("A21").Value, Sheets("Resultats").Range("A19").Value, Sheets("Resultats").Range("B21").Value, Sheets("Resultats").Range("C21").Value)
Call copie_colle_year_now(Sheets("Resultats").Range("A38").Value, Sheets("Resultats").Range("A36").Value, Sheets("Resultats").Range("B38").Value, Sheets("Resultats").Range("C38").Value)
Call copie_colle_year_now(Sheets("Resultats").Range("A55").Value, Sheets("Resultats").Range("A53").Value, Sheets("Resultats").Range("B55").Value, Sheets("Resultats").Range("C55").Value)

Call maj_tab_resultat_global

On Error Resume Next
Dim nom_nouveau_sans_Bb_formula As String
nom_nouveau_sans_Bb_formula = ThisWorkbook.Path & "\" & ThisWorkbook.Name & " sans formule BDH.xlsm"
ActiveWorkbook.SaveAs Filename:=nom_nouveau_sans_Bb_formula, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
  
'Application.ScreenUpdating = True
UserForm_demo.Hide
MsgBox "Sheets updated"
Sheets("Resultats").Select


End Sub

Sub copie_colle_concurence_price(onglet_copie As String, position_debut_tableau As Integer, nb_bank_concurence As Integer)

fin = position_debut_tableau + nb_bank_concurence - 1
For j = position_debut_tableau To fin

    name_bank = Sheets("Input_VBA").Cells(32, j).Value

    colonne_bank = Cellule_recherchee(name_bank, Sheets(onglet_copie).Range("a2:z2"))
    colonne_bank = Range(colonne_bank).Column
    
    num_ligne_start_tab = 33
    end_ligne_tab = 39
    
    For i = num_ligne_start_tab To end_ligne_tab

    
        valeur_a_tester = CInt(Sheets("Input_VBA").Cells(i, 1).Value)
        valeur_a_copier = Sheets("Input_VBA").Cells(i, j).Value
            
        fin_1 = Findutableau_ligne(onglet_copie, "A")
        fin_2 = Findutableau_ligne(onglet_copie, "B")
        
        If fin_1 > fin_2 Then
        fin = fin_1
        Else
        fin = fin_2
        End If

                 For h = 3 To fin
                        
                    val_feuille = Sheets(onglet_copie).Cells(h, 2).Value
          If (IsError(val_feuille) = False) Then
                    If (Len(val_feuille) <= 6) Then
                    val_actu = CInt(2000 + Right(Sheets(onglet_copie).Cells(h, 2).Value, 2))
                    Else
                    val_actu = CInt(Right(Sheets(onglet_copie).Cells(h, 2).Value, 4))
                     End If
          End If
                
                    If val_actu = valeur_a_tester Then
                    Sheets(onglet_copie).Cells(h, colonne_bank).Value = valeur_a_copier
                    End If
                    
                Next h
                
    Next i



Next j


End Sub
Sub import()

    Dim NomFeuille As String
    Dim chemin As String
    Dim plage_cellules As String
    Dim endroit_a_coller As String
    Dim options_entete As Boolean
    
    
    NomFeuille = "ConfidenceLevels"
    Sheets("Import").Activate
    Call RequeteClasseurFerme_valeur(Fichier, NomFeuille, "E8:E17", "A2", False)
    Call RequeteClasseurFerme_valeur(Fichier, NomFeuille, "G8:G17", "B2", False)
    Call RequeteClasseurFerme_valeur(Fichier, NomFeuille, "I8:I17", "C2", False)
    Call RequeteClasseurFerme_valeur(Fichier, NomFeuille, "M41:M50", "D2", False)
    
    MsgBox "Table importée"
End Sub


Function fin_colonne(onglet_actu As String) As Integer

fin_colonne = Range("A16").End(xlToRight).Column

        
End Function
Function LetCol(NoCol)
LetCol = Split(Cells(1, NoCol).Address, "$")(1)
End Function



Function Browse(Optional ByVal rep_defaut As String = "", Optional ByVal filtres As String = "*.xls; *.xlsx; *.xlsm ; *.csv ; *.txt", Optional ByRef type_boite As MsoFileDialogType = msoFileDialogFilePicker) As Variant
    If rep_defaut = "" Then rep_defaut = ThisWorkbook.Path
    If Len(rep_defaut) > 0 And InStr(1, rep_defaut, ".") > 0 Then rep_defaut = Left(rep_defaut, InStrRev(rep_defaut, "\") - 1)
   

    'Declare a variable as a FileDialog object.
    Dim fd As FileDialog

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(type_boite)

    'Declare a variable to contain the path
    'of each selected item. Even though the path is aString,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant

    'Use a With...End With block to reference the FileDialog object.
    With fd
            
         'Allow the selection of multiple file.
        .AllowMultiSelect = False
        .Title = "Sélectionner le fichier"
        Select Case type_boite
         Case msoFileDialogFilePicker
            .InitialFileName = rep_defaut
            .Filters.Clear
            .Filters.Add "Fichiers autorisés", filtres, 1
        Case msoFileDialogFolderPicker
            .InitialFileName = rep_defaut
        
        Case Else
            .FilterIndex = 12
            .InitialFileName = rep_defaut & "\MatriceCompetences_Save_" & Format(Date, "yyyymmdd") & ".txt"
            
        End Select
       
        If .Show = -1 Then Browse = .SelectedItems(1)

    End With


    'Set the object variable to Nothing.
    Set fd = Nothing

End Function

Function FichierEstOuvert(ByRef FichierTeste As String) As Boolean
    Dim Fichier As Long
    On Error GoTo Erreur
    Fichier = FreeFile
    Open FichierTeste For Input Lock Read As #Fichier
    Close #Fichier
    FichierEstOuvert = False
    Exit Function
Erreur:
    FichierEstOuvert = True
End Function




Sub copie_colle_valeur_tableau_commun(onglet_actu As String, onglet_copie As String, num_colonne As Integer)


If onglet_copie = "NBP" Then
colonne_BB = Cellule_recherchee("ICE UK Natural Gas", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "Brent" Then
colonne_BB = Cellule_recherchee("ICE Brent Crude Oil ", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "US GAS" Then
colonne_BB = Cellule_recherchee("ICE US GAS ", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "AECO GAS" Then
colonne_BB = Cellule_recherchee("ICE AECO GAS", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "Crude" Then
colonne_BB = Cellule_recherchee("ICE Crude", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If



colonne_BC = Cellule_recherchee("SG Base Case", Sheets(onglet_copie).Range("a2:z2"))
colonne_BC = Range(colonne_BC).Column

colonne_SC = Cellule_recherchee("SG Sens Case", Sheets(onglet_copie).Range("a2:z2"))
colonne_SC = Range(colonne_SC).Column


For i = 32 To 41

If type_version = 33 Then
valeur_a_tester = CInt(Right(Sheets(onglet_actu).Cells(i, 2).Value, 4))
End If

If type_version = 1 Then
valeur_a_tester = CInt(2000 + Right(Sheets(onglet_actu).Cells(i, 2).Value, 2))
End If



valeur_a_copier = Sheets(onglet_actu).Cells(i, 3).Value
valeur_a_copier_BC = Sheets(onglet_actu).Cells(i, 4).Value
valeur_a_copier_SC = Sheets(onglet_actu).Cells(i, 5).Value


        fin_1 = Findutableau_ligne(onglet_copie, "A")
        fin_2 = Findutableau_ligne(onglet_copie, "B")
        
        If fin_1 > fin_2 Then
        fin = fin_1
        Else
        fin = fin_2
        End If
        
        
        For j = 3 To fin
        On Error Resume Next
        If Len(Sheets(onglet_copie).Cells(j, 2).Value) < 2 Then
            Sheets(onglet_copie).Cells(j, 2).EntireRow.Delete
        End If
        
        If val_feuille = Len(Sheets(onglet_copie).Cells(j, 2).Value) <= 6 Then
        val_actu = CInt(2000 + Right(Sheets(onglet_copie).Cells(j, 2).Value, 2))
        Else
        val_actu = CInt(Right(Sheets(onglet_copie).Cells(j, 2).Value, 4))
        End If
        

                If val_actu = valeur_a_tester Then
                 n = n + 1
                 somme_BB = somme_BB + Sheets(onglet_copie).Cells(j, colonne_BB)
                 Sheets(onglet_copie).Cells(j, num_colonne).Value = valeur_a_copier
                 Sheets(onglet_copie).Cells(j, colonne_BC).Value = valeur_a_copier_BC
                 Sheets(onglet_copie).Cells(j, colonne_SC).Value = valeur_a_copier_SC
                End If
                
        Next j

If somme_BB <> 0 Then
moyenne_BB = somme_BB / n
Sheets(onglet_actu).Cells(i, 6).Value = Round(moyenne_BB)
Else
Sheets(onglet_actu).Cells(i, 6).Value = "N/A"
End If


  
'Remise à 0
moyenne_BB = 0
somme_BB = 0
n = 0

Next i

End Sub
Public Function Cellule_recherchee(ByVal val_cherchee As String, ByRef plage As Range)
On Error Resume Next
Dim cellule As Range

 
For Each cellule In plage.Cells

    If cellule.Value = val_cherchee Then
    Cellule_recherchee = cellule.Address: Exit Function
    End If

Next cellule

Cellule_recherchee = Null

End Function

Public Function Findutableau_ligne(ByVal num_feuille As String, ByVal num_colonne As String)

Sheets(num_feuille).Activate
ThisWorkbook.Sheets(num_feuille).Select
Set myrange = Cells(1048576, num_colonne)
myrange.Select
Selection.End(xlUp).Select
Findutableau_ligne = Selection.Row

End Function

Sub maj_tableur_avec_vals_new(debut As Double, fin As Double, pas As Double, pourcentage As Double, onglet_actu As String, nom_plage_actu As String, plage_a_recup As String)


Wb2.Sheets(onglet_actu).Activate
Range("A16").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.ClearContents
Selection.ClearFormats
Selection.Interior.ColorIndex = xlNone

Call initialisation


For i = 2 To 1000

    If debut > fin Then Exit For
    
    Cells(16, i).Value = debut
    debut = debut + pas
    

Next i
Dim nombre_de_colonne As Integer
nombre_de_colonne = fin_colonne(onglet_actu)
For i = 2 To nombre_de_colonne
                
                'Argument à mettre dans la plage
                argument = Wb2.Sheets(onglet_actu).Cells(16, i).Value
                Wb2.Sheets(onglet_actu).Cells(16, i).Value = argument
                
                
                'MAJ
                Wb1.Sheets("ConfidenceLevels").Range(nom_plage_actu).Value = argument
                                
                lettre = LetCol(i)
                plage = lettre & "17:" & lettre & "26"

                'Trouver comment modifier la seconde plage
                Wb2.Sheets(onglet_actu).Range(plage).Value = Wb1.Sheets("ConfidenceLevels").Range(plage_a_recup).Value
        
Next i

Range("B17").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Style = "Percent"
Selection.NumberFormat = "0.00%"
Cells.Select
Cells.EntireColumn.AutoFit

Call parcours_prix(onglet_actu, pourcentage, nombre_de_colonne)



End Sub

Sub parcours_prix(onglet_actu As String, pourcentage As Double, nombre_de_colonne As Integer)

Dim difference As Double
Dim difference_mini As Double
Dim position As Integer


Range("B30:H43").Select
Selection.ClearContents

For j = 17 To 26
difference_mini = 100
        For i = 2 To nombre_de_colonne
            
            val_actu = 100 * Cells(j, i).Value
            difference = val_actu - pourcentage
            If ((difference > 0) And (difference < difference_mini)) Then
            difference_mini = difference
            position = i
            End If
        Next i
        
        On Error Resume Next
        val_potentiel = 100 * (Wb2.Sheets(onglet_actu).Cells(j, position).Value)
        
        If val_potentiel > pourcentage Then
        Wb2.Sheets(onglet_actu).Cells(j, position).Select
        Selection.Interior.Color = 255
        
        On Error GoTo 0
        'Tableau
        Cells(j, i + 2).Value = CDate(Left(Date, 5) & "/" & Cells(j, 1).Value)
            If (onglet_actu = "UK Gas Base Case" Or onglet_actu = "UK Gas Sensitivity Case") Then
            Cells(j, i + 3).Value = 100 * Cells(16, position).Value
            Else
            Cells(j, i + 3).Value = Cells(16, position).Value
            End If
            
        Else
        Cells(j, i + 2).Value = CDate(Left(Date, 5) & "/" & Cells(j, 1).Value)
        Cells(j, i + 3).Value = 0
        End If
Next j

'Change position du tableau
Cells(17, i + 2).Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Cut
Range("B32").Select
ActiveSheet.Paste
Columns("B:B").EntireColumn.AutoFit

Cells(30, 2).Value = "%"
Cells(30, 3).Value = pourcentage
Cells(42, 2).Value = "Mean"
Cells(42, 3).FormulaR1C1 = "=AVERAGE(R[-10]C:R[-1]C)"
Cells(43, 2).Value = "Std dev"
Cells(43, 3).FormulaR1C1 = "=ROUND(STDEVA(R[-11]C:R[-2]C),2)"

Cells(31, 2).Value = "Year"
Cells(31, 3).Value = "Confidence Price"
Cells(31, 4).Value = "SG Base Case"
Cells(31, 5).Value = "SG Sens Case"
Cells(31, 6).Value = "Mean Forward"
Cells(31, 7).Value = "Discount SG base Case"
Cells(31, 8).Value = "Discount SG Sens Case"
Columns("D:D").EntireColumn.AutoFit

'Formules tableau
Range("G32").FormulaR1C1 = "=IFERROR(ABS(RC[-3]-RC[-1])/RC[-1],""N/A"")"
Range("G32").Select
Selection.AutoFill Destination:=Range("G32:G43")

Range("H32").FormulaR1C1 = "=IFERROR(ABS(RC[-3]-RC[-2])/RC[-2],""N/A"")"
Range("H32").Select
Selection.AutoFill Destination:=Range("H32:H43")

If (onglet_actu = "Brent Base Case" Or onglet_actu = "Brent Sensitivity Case") Then
Range("D32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[4]),"""",Input_VBA!R[-29]C[4])"
Range("D32").Select
Selection.AutoFill Destination:=Range("D32:D43")

Range("E32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[4]),"""",Input_VBA!R[-29]C[4])"
Range("E32").Select
Selection.AutoFill Destination:=Range("E32:E43")
End If

If (onglet_actu = "UK Gas Base Case" Or onglet_actu = "UK Gas Sensitivity Case") Then
    Range("D32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[6]),"""",Input_VBA!R[-29]C[6]*100)"
    Range("D32").Select
    Selection.AutoFill Destination:=Range("D32:D43")
    
    Range("E32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[6]),"""",Input_VBA!R[-29]C[6]*100)"
    Range("E32").Select
    Selection.AutoFill Destination:=Range("E32:E43")
End If

If (onglet_actu = "Crude Base Case" Or onglet_actu = "Crude Sensitivity Case") Then
    Range("D32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[8]),"""",Input_VBA!R[-29]C[8])"
    Range("D32").Select
    Selection.AutoFill Destination:=Range("D32:D43")

    Range("E32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[8]),"""",Input_VBA!R[-29]C[8])"
    Range("E32").Select
    Selection.AutoFill Destination:=Range("E32:E43")
End If

If (onglet_actu = "US GAS Base Case " Or onglet_actu = "US GAS Sensitivity Case") Then
    Range("D32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[10]),"""",Input_VBA!R[-29]C[10])"
    Range("D32").Select
    Selection.AutoFill Destination:=Range("D32:D43")
 
    Range("E32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[10]),"""",Input_VBA!R[-29]C[10])"
    Range("E32").Select
    Selection.AutoFill Destination:=Range("E32:E43")
End If

If (onglet_actu = "AECO GAS Base Case" Or onglet_actu = "AECO GAS Sensitivity Case") Then
    Range("D32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[12]),"""",Input_VBA!R[-29]C[12])"
    Range("D32").Select
    Selection.AutoFill Destination:=Range("D32:D43")
    
    Range("E32").FormulaR1C1 = "=IF(ISBLANK(Input_VBA!R[-29]C[12]),"""",Input_VBA!R[-29]C[12])"
    Range("E32").Select
    Selection.AutoFill Destination:=Range("E32:E43")
End If



'Trace graphique
For Each Legraph In ActiveSheet.ChartObjects
    Legraph.Delete
Next

Dim abscisses As String
Dim ordonnees As String

abscisses = "='" & onglet_actu & "'!" & "$B$32:$B$41"
ordonnees = "='" & onglet_actu & "'!" & "$C$32:$C$41"
Charts.Add
ActiveChart.Location Where:=xlLocationAsObject, Name:=onglet_actu

On Error Resume Next

With ActiveChart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = abscisses
        .SeriesCollection(1).Values = ordonnees
        .ChartType = xlLine
        .HasTitle = False
        .HasTitle = True
        .ChartTitle.Characters.Text = "Price of " & onglet_actu & " at " & pourcentage & "%"
End With
ActiveChart.Legend.Select
Selection.Delete
ActiveSheet.ChartObjects(1).Left = Range("B46").Left
ActiveSheet.ChartObjects(1).Top = Range("B46").Top
ActiveSheet.ChartObjects(1).Height = 450
ActiveSheet.ChartObjects(1).Width = 800

Call mise_en_forme_tab






End Sub

Sub initialisation()


type_version = Application.International(xlCountryCode)
If type_version = 33 Then
Cells(16, 1).Value = "Year"
Cells(17, 1).Value = CStr(CInt(Right(Date, 4)) + 1)
Cells(18, 1).Value = CStr(CInt(Right(Date, 4)) + 2)
Cells(19, 1).Value = CStr(CInt(Right(Date, 4)) + 3)
Cells(20, 1).Value = CStr(CInt(Right(Date, 4)) + 4)
Cells(21, 1).Value = CStr(CInt(Right(Date, 4)) + 5)
Cells(22, 1).Value = CStr(CInt(Right(Date, 4)) + 6)
Cells(23, 1).Value = CStr(CInt(Right(Date, 4)) + 7)
Cells(24, 1).Value = CStr(CInt(Right(Date, 4)) + 8)
Cells(25, 1).Value = CStr(CInt(Right(Date, 4)) + 9)
Cells(26, 1).Value = CStr(CInt(Right(Date, 4)) + 10)
ElseIf type_version = 1 Then
Cells(16, 1).Value = "Year"
Cells(17, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 1)
Cells(18, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 2)
Cells(19, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 3)
Cells(20, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 4)
Cells(21, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 5)
Cells(22, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 6)
Cells(23, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 7)
Cells(24, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 8)
Cells(25, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 9)
Cells(26, 1).Value = CStr(2000 + CInt(Right(Date, 2)) + 10)
End If





End Sub

Sub code()
    
'Fonction savoir le chemin du fichier référence
Fichier = Browse(, "*.xlsb;*.xls; *.xlsx")
If Len(Fichier) = 0 Then Exit Sub

'test si fichier ouvert
If FichierEstOuvert(Fichier) = True Then
    a = Split(Fichier, "\")
    nom_fichier = a(UBound(a))
    Windows(nom_fichier).Close False
End If

Workbooks.Open Filename:=Fichier
Set Wb1 = ActiveWorkbook
Set Wb2 = ThisWorkbook

Wb2.Sheets("Input_VBA").Activate
UserForm_demo.CommandButton1.Caption = "Click for updating the prices"
UserForm_demo.Show


End Sub




Sub mise_en_forme_tab()
 Range("B30:H43").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
End Sub

Sub maj_tab_resultat_global()
    Sheets("Resultats").Select
    
    Range("D4").Value = mean_future(Range("A4"), "Crude")
    Range("D21").Value = mean_future(Range("A21"), "US GAS")
    Range("D38").Value = mean_future(Range("A38"), "Brent")
    Range("D55").Value = mean_future(Range("A55"), "NBP")
   
    
    
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "='Crude Base Case'!R[27]C[-4]"
    Range("G5").Select
    Selection.AutoFill Destination:=Range("G5:G14"), Type:=xlFillDefault
    Range("G5:G14").Select
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "='Crude Sensitivity Case'!R[27]C[-5]"
    Range("H5").Select
    Selection.AutoFill Destination:=Range("H5:H14"), Type:=xlFillDefault
    Range("H5:H14").Select
    Range("G15").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-10]C:R[-1]C)"
    Range("H15").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-10]C:R[-1]C)"
    Range("G16").Select
    ActiveCell.FormulaR1C1 = "=STDEV(R[-11]C:R[-1]C)"
    Range("H16").Select
    ActiveCell.FormulaR1C1 = "=STDEV(R[-11]C:R[-1]C)"
    Range("G22").Select
    ActiveCell.FormulaR1C1 = "='US GAS Base Case '!R[10]C[-4]"
    Range("G22").Select
    Selection.AutoFill Destination:=Range("G22:G31"), Type:=xlFillDefault
    Range("G22:G31").Select
    Range("H22").Select
    ActiveCell.FormulaR1C1 = "='US GAS Sensitivity Case'!R[10]C[-5]"
    Range("H22").Select
    Selection.AutoFill Destination:=Range("H22:H31"), Type:=xlFillDefault
    Range("H22:H31").Select
    Range("G32").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-10]C:R[-1]C)"
    Range("G33").Select
    ActiveCell.FormulaR1C1 = "=STDEV(R[-11]C:R[-2]C)"
    Range("G32:G33").Select
    Selection.AutoFill Destination:=Range("G32:H33"), Type:=xlFillDefault
    Range("G32:H33").Select
    Range("G39").Select
    ActiveCell.FormulaR1C1 = "='Brent Base Case'!R[-7]C[-4]"
    Range("G39").Select
    Selection.AutoFill Destination:=Range("G39:G48"), Type:=xlFillDefault
    Range("G39:G48").Select
    Range("H39").Select
    ActiveCell.FormulaR1C1 = "='Brent Sensitivity Case'!R[-7]C[-5]"
    Range("H39").Select
    Selection.AutoFill Destination:=Range("H39:H48"), Type:=xlFillDefault
    Range("H39:H48").Select
    Range("G49").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-10]C:R[-1]C)"
    Range("G50").Select
    ActiveCell.FormulaR1C1 = "=STDEV(R[-11]C:R[-1]C)"
    Range("G51").Select
    ActiveWindow.SmallScroll Down:=12
    Range("G49:G50").Select
    Selection.AutoFill Destination:=Range("G49:H50"), Type:=xlFillDefault
    Range("G49:H50").Select
    Range("J49").Select
    ActiveWindow.SmallScroll Down:=15
    Range("G56").Select
    ActiveCell.FormulaR1C1 = "='UK Gas Base Case'!R[-24]C[-4]"
    Range("G56").Select
    Selection.AutoFill Destination:=Range("G56:G65"), Type:=xlFillDefault
    Range("G56:G65").Select
    Range("G66").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-10]C:R[-1]C)"
    Range("G67").Select
    ActiveCell.FormulaR1C1 = "=STDEV(R[-11]C:R[-2]C)"
    Range("G66:G67").Select
    Selection.AutoFill Destination:=Range("G66:H67"), Type:=xlFillDefault
    Range("G66:H67").Select
    Range("H56").Select
    ActiveCell.FormulaR1C1 = "='UK Gas Sensitivity Case'!R[-24]C[-5]"
    Range("H56").Select
    Selection.AutoFill Destination:=Range("H56:H65"), Type:=xlFillDefault
    Range("H56:H65").Select
    Range("K63").Select
End Sub

Function mean_future(year As String, onglet_copie As String) As Double

If onglet_copie = "NBP" Then
colonne_BB = Cellule_recherchee("ICE UK Natural Gas", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "Brent" Then
colonne_BB = Cellule_recherchee("ICE Brent Crude Oil ", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "US GAS" Then
colonne_BB = Cellule_recherchee("ICE US GAS ", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "AECO GAS" Then
colonne_BB = Cellule_recherchee("ICE AECO GAS", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "Crude" Then
colonne_BB = Cellule_recherchee("ICE Crude", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

colonne_BC = Cellule_recherchee("SG Base Case", Sheets(onglet_copie).Range("a2:z2"))
colonne_BC = Range(colonne_BC).Column

colonne_SC = Cellule_recherchee("SG Sens Case", Sheets(onglet_copie).Range("a2:z2"))
colonne_SC = Range(colonne_SC).Column

valeur_a_tester = CInt(year)
n = 0
val_date_today = CDate(Left(Now(), 10))

Dim vec_passage(3000) As Boolean

For j = 3 To 3000


        If val_feuille = Len(Sheets(onglet_copie).Cells(j, 2).Value) <= 6 Then
            val_actu = CInt(2000 + Right(Sheets(onglet_copie).Cells(j, 2).Value, 2))
            val_actu_full = Sheets(onglet_copie).Cells(j, 2).Value
        Else
            val_actu = CInt(Right(Sheets(onglet_copie).Cells(j, 2).Value, 4))
            val_actu_full = Sheets(onglet_copie).Cells(j, 2).Value
        End If



                        If (val_actu = valeur_a_tester) And (val_actu_full > val_date_today) Then
                         n = n + 1
                         new_val_to_add = Sheets(onglet_copie).Cells(j, colonne_BB)
                         somme_BB = somme_BB + new_val_to_add
                         vec_passage(j) = True
                         Else
                        vec_passage(j) = False
                        End If


        If (vec_passage(j) = False And vec_passage(j - 1) = True) Then
        Exit For
        End If


Next j


moyenne_BB = somme_BB / n
mean_future = Round(moyenne_BB)

End Function

Sub copie_colle_year_now(year As String, onglet_copie As String, valeur_BC As Double, valeur_SC As Double)

If onglet_copie = "NBP" Then
colonne_BB = Cellule_recherchee("ICE UK Natural Gas", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "Brent" Then
colonne_BB = Cellule_recherchee("ICE Brent Crude Oil ", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "US GAS" Then
colonne_BB = Cellule_recherchee("ICE US GAS ", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "AECO GAS" Then
colonne_BB = Cellule_recherchee("ICE AECO GAS", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

If onglet_copie = "Crude" Then
colonne_BB = Cellule_recherchee("ICE Crude", Sheets(onglet_copie).Range("a2:z2"))
colonne_BB = Range(colonne_BB).Column
End If

colonne_BC = Cellule_recherchee("SG Base Case", Sheets(onglet_copie).Range("a2:z2"))
colonne_BC = Range(colonne_BC).Column

colonne_SC = Cellule_recherchee("SG Sens Case", Sheets(onglet_copie).Range("a2:z2"))
colonne_SC = Range(colonne_SC).Column

valeur_a_copier_BC = valeur_BC
valeur_a_copier_SC = valeur_SC

valeur_a_tester = CInt(year)
n = 0
val_date_today = CDate(Left(Now(), 10))
Dim vec_passage(3000) As Boolean

For j = 3 To 3000


    val_feuille = Sheets(onglet_copie).Cells(j, 2).Value
    If (IsError(val_feuille) = False) Then
            If val_feuille = Len(Sheets(onglet_copie).Cells(j, 2).Value) <= 6 Then
            val_actu = CInt(2000 + Right(Sheets(onglet_copie).Cells(j, 2).Value, 2))
            val_actu_full = Sheets(onglet_copie).Cells(j, 2).Value
            Else
            val_actu = CInt(Right(Sheets(onglet_copie).Cells(j, 2).Value, 4))
            val_actu_full = Sheets(onglet_copie).Cells(j, 2).Value
            End If
    Else
    Exit For
    End If
     
     
    If val_actu = valeur_a_tester Then
        Sheets(onglet_copie).Cells(j, colonne_BC).Value = valeur_a_copier_BC
        Sheets(onglet_copie).Cells(j, colonne_SC).Value = valeur_a_copier_SC
        vec_passage(j) = True
    Else
        vec_passage(j) = False
    End If

        If (vec_passage(j) = False And vec_passage(j - 1) = True) Then
        Exit For
        End If


Next j

End Sub

  
