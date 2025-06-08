Attribute VB_Name = "Module2"
Public x() As Double
Dim plage_to_search As Variant



Sub get_pourcentage_from_porposition_SG(feuille As String)

Sheets(feuille).Activate
Range("A16").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Font.ColorIndex = xlNone

Range("C2").Select
Selection.End(xlDown).Select
fin = Selection.Row

ReDim x(7)
Sheets(feuille).Activate
Range("E3:E" & fin).Value = ""
j = 17
position_vec = 0



For i = 3 To fin

Price_to_find = Cells(i, 4).Value
        If IsNumeric(Price_to_find) = True Then
                Date_to_find = Cells(i, 3).Value
                    
                    adress_val = val_recherchee(Price_to_find, Range("B16:ZZ16"))
                    num_colonne = Range(adress_val).Column
                    
                    a = 100 * Cells(j, num_colonne).Value
                    Cells(j, num_colonne).Select
                    Selection.Font.ColorIndex = 8
                    Cells(i, 5).Value = a
                    
                  
                    j = j + 1
                    position_vec = position_vec + 1
        End If

    
Next i
   
Range("E3").Select
   


End Sub

Public Function val_recherchee(ByVal val_cherchee As Double, ByRef plage As Range)
Dim cellule As Range

 
For Each cellule In plage.Cells


    If Abs(cellule.Value - val_cherchee) < 0.0001 Then
    
    val_recherchee = cellule.Address: Exit Function
    End If

Next cellule

val_recherchee = Null

End Function

Sub main_get_pourcentage_from_porposition_SG()

'US
Call get_pourcentage_from_porposition_SG("Crude Base Case")
Call get_pourcentage_from_porposition_SG("Crude Sensitivity Case")
Call get_pourcentage_from_porposition_SG("US GAS Base Case ")
Call get_pourcentage_from_porposition_SG("US GAS Sensitivity Case")
Call get_pourcentage_from_porposition_SG("AECO GAS Base Case")
Call get_pourcentage_from_porposition_SG("AECO GAS Sensitivity Case")

'UE
Call get_pourcentage_from_porposition_SG("Brent Base Case")
Call get_pourcentage_from_porposition_SG("Brent Sensitivity Case")
Call get_pourcentage_from_porposition_SG("UK Gas Sensitivity Case")
Call get_pourcentage_from_porposition_SG("UK Gas Base Case")

End Sub





