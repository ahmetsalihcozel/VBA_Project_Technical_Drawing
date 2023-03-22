Public Function Ent2lspEnt(entObj As AcadEntity) As String
'Designed to work with SendCommand, which can't pass objects.
'This gets an objects handle and converts it to a string
'of lisp commands that returns an entity name when run in SendCommand.
Dim entHandle As String

entHandle = entObj.Handle
Ent2lspEnt = "(handent " & Chr(34) & entHandle & Chr(34) & ")"
End Function
