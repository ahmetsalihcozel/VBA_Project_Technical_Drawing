Public Sub Symbol_Switcher()

Dim SecInCorrect As AcadSelectionSet
Dim BlRef As AcadBlockReference
Dim varAttributes As Variant
Dim InnerObjects As AcadObject
Dim AttVarmi As Boolean
Dim nameofInner As Variant
Dim explodedObjects As Variant
Dim explodedObjects1 As Variant
Dim explodedObjects2 As Variant
Dim xNesneSayac(0 To 100) As Variant
Dim deneme As Integer
Dim CompareArrayIndex(100, 100) As Variant
Dim HerhangiObje As AcadEntity
Dim InsertCorrectBlocks As AcadBlockReference
Dim explodedInfosArray(100) As Variant
Dim EqualityCounter As Integer
Dim Rotasyon As String
Dim InsertNok0 As String
Dim InsertNok1 As String
Dim InsertNoktasi As String
Dim TorF As Integer
Dim InsertNok0Str As String
Dim InsertNok1Str As String
Dim Noktalar As Variant
Dim NumaraSirala(0 To 100) As Variant
Dim InsertNok As Double
Dim SiralamaObje As AcadEntity
Dim BlockSayi As Variant
Dim RotasyonStr As String
Const PiNum As Double = 3.14159265358979
Dim OldText As String
Dim NewText As String


On Error Resume Next

BlockSayi = ThisDrawing.Utility.GetInteger

ThisDrawing.SelectionSets("SecInCorrect").Delete

For i = 0 To BlockSayi - 1
Set SecInCorrect = ThisDrawing.SelectionSets.Add("SecInCorrect")
SecInCorrect.SelectOnScreen
Next i





For i = 0 To SecInCorrect.Count - 1

If SecInCorrect(i).HasAttributes Then

explodedObjects = SecInCorrect(i).Explode

explodedInfosArray(i) = explodedObjects

xNesneSayac(i) = 0

For k = 0 To UBound(explodedObjects) - 1
If Not TypeOf explodedObjects(k) Is IAcadAttribute Then

CompareArrayIndex(i, xNesneSayac(i)) = k

xNesneSayac(i) = xNesneSayac(i) + 1

If k = UBound(explodedObjects) - 1 Then
CompareArrayIndex(i, xNesneSayac(i)) = UBound(explodedObjects)
End If
End If
Next k



End If

Next i

BlockSayi = ThisDrawing.Utility.GetInteger

ThisDrawing.SelectionSets("SecCorrect").Delete

For i = 0 To BlockSayi - 1
Set SecCorrect = ThisDrawing.SelectionSets.Add("SecCorrect")
SıralamaObje = SecCorrect.SelectOnScreen
SecCorrect.AddItems SıralamaObje
Next i


For Each HerhangiObje In ThisDrawing.ModelSpace
If TypeOf HerhangiObje Is IAcadBlockReference Then


If HerhangiObje.HasAttributes Then

TorF = 1

Rotasyon = (HerhangiObje.Rotation / PiNum) * 180
RotasyonStr = CStr(Rotasyon)
RotasyonStr = Replace(RotasyonStr, ",", ".")

Noktalar = HerhangiObje.InsertionPoint

InsertNok0 = Noktalar(0)
InsertNok0Str = CStr(InsertNok0)
InsertNok0Str = Replace(InsertNok0Str, ",", ".")

InsertNok1 = Noktalar(1)
InsertNok1Str = CStr(InsertNok1)
InsertNok1Str = Replace(InsertNok1Str, ",", ".")


explodedObjects1 = HerhangiObje.Explode

ThisDrawing.SendCommand "_U" & vbCrLf


If explodedObjects1(0) Is Nothing Then
TorF = 0
End If


If TorF = 1 Then


EqualityCounter = 0

For i = 0 To SecInCorrect.Count - 1
explodedObjects2 = SecInCorrect(i).Explode

If UBound(explodedObjects1) = UBound(explodedObjects2) Then
For x = 0 To xNesneSayac(i) - 1

If explodedObjects1(CompareArrayIndex(i, x)) = explodedObjects2(CompareArrayIndex(i, x)) Then

If (TypeOf explodedObjects1(CompareArrayIndex(i, x)) Is IAcadCircle And TypeOf explodedObjects2(CompareArrayIndex(i, x)) Is IAcadCircle) Then
EqualityCounter = EqualityCounter + 1

ElseIf (TypeOf explodedObjects1(CompareArrayIndex(i, x)) Is IAcadLine And TypeOf explodedObjects2(CompareArrayIndex(i, x)) Is IAcadLine) Then
EqualityCounter = EqualityCounter + 1

ElseIf (TypeOf explodedObjects1(CompareArrayIndex(i, x)) Is IAcadHatch And TypeOf explodedObjects2(CompareArrayIndex(i, x)) Is IAcadHatch) Then
EqualityCounter = EqualityCounter + 1

ElseIf (TypeOf explodedObjects1(CompareArrayIndex(i, x)) Is IAcadArc And TypeOf explodedObjects2(CompareArrayIndex(i, x)) Is IAcadArc) Then
EqualityCounter = EqualityCounter + 1

End If
End If

Next x
End If

If EqualityCounter = xNesneSayac(i) Then

OldText = OldTAG1.Text

varAttributes = HerhangiObje.GetAttributes   'Old Block Attributes
GeometryAtt = HerhangiObje.GetConstantAttributes

For j = LBound(varAttributes) To UBound(varAttributes)

If varAttributes(j).TagString = OldTAG1.Text Then

OldBlockAttText = varAttributes(j).TextString

End If

Next j



NewText = NewTAG1.Text

varAttributes = SecCorrect.Item(i).GetAttributes   'Old Block Attributes

For j = LBound(varAttributes) To UBound(varAttributes)

If varAttributes(j).TagString = NewTAG1.Text Then

varAttributes(j).TextString = OldBlockAttText

End If

Next j






StrLspEnt = Ent2lspEnt(SecCorrect.Item(i))

Noktalar = SecCorrect.Item(i).InsertionPoint
copynok1 = Noktalar(0)
copynok2 = Noktalar(1)

copynokstr1 = CStr(copynok1)
copynokstr1 = Replace(copynokstr1, ",", ".")

copynokstr2 = CStr(copynok2)
copynokstr2 = Replace(copynokstr2, ",", ".")

ThisDrawing.SendCommand "_COPYBASE" & vbCrLf & copynokstr1 & "," & copynokstr2 & vbLf & StrLspEnt & vbCrLf

ThisDrawing.SendCommand "_PASTECLIP" & vbCrLf & "r" & vbLf & RotasyonStr & vbLf & InsertNok0Str & "," & InsertNok1Str & vbLf

End If

Next i

End If
End If
End If

Next HerhangiObje


End Sub
