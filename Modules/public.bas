Attribute VB_Name = "public"
Option Compare Database

Function ObjectExists(strObjectType As String, strObjectName As String) As Boolean
     Dim db As Database
     Dim tbl As TableDef
     Dim qry As QueryDef
     Dim i As Integer
     
     Set db = CurrentDb()
     ObjectExists = False
     
     If strObjectType = "Table" Then
          For Each tbl In db.TableDefs
               If tbl.Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next tbl
     ElseIf strObjectType = "Query" Then
          For Each qry In db.QueryDefs
               If qry.Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next qry
     ElseIf strObjectType = "Form" Or strObjectType = "Report" Or strObjectType = "Module" Then
          For i = 0 To db.Containers(strObjectType & "s").Documents.Count - 1
               If db.Containers(strObjectType & "s").Documents(i).Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next i
     ElseIf strObjectType = "Macro" Then
          For i = 0 To db.Containers("Scripts").Documents.Count - 1
               If db.Containers("Scripts").Documents(i).Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next i
     Else
          MsgBox "Invalid Object Type passed, must be Table, Query, Form, Report, Macro, or Module"
     End If
     
End Function
'gets value from table based on id
Public Function getvalue(field1 As Variant, idvar As Variant)
Dim valph As Variant

valph = DLookup(field1, "TREE", "Point_ID=" & idvar)
If InStr(1, valph, " ") <= 5 Then
getvalue = StrConv(Mid(valph, (InStr(1, valph, " ") + 1)), vbProperCase)
Else
getvalue = StrConv(valph, vbProperCase)
End If
End Function
Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function
'populates values with extra "other comment" field
Public Function popval(fieldnamevar As String, tablevar As String, othval As Variant, idvar As Variant)
Dim othchk As Variant

If IsNull(othval) = False Then
    othchk = getvalue(othval, idvar)
    If IsNull(othchk) = False Then
        popval = othchk
        Exit Function
    End If
End If
    
    Dim getval As Variant
    Dim geteval As Variant
    'getval = getvalue(fieldnamevar)
    getval = Nz(DLookup(fieldnamevar, "TREE", "Point_ID=" & idvar), "     ")
    geteval = DLookup("desc", tablevar, "code='" & Left(getval, (InStr(1, getval, " ") - 1)) & "'")

If geteval = Null Then
    popval = "Unspecified"
Else
    popval = geteval
End If

End Function



Public Function processdate(datevar As String)
Dim dateyear As String
Dim datemonth As String
Dim dateday As String
Dim forcedatevar As String

processdate = Format(datevar, "YYYYMMDD")
End Function

' evaluate "z other" value in fields
Public Function evalother(currentval As Variant, otherval As Variant)
If currentval = Null Then
evalother = "Unspecified"
End If

If Left(currentval, 1) = "Z" Or Left(currentval, 5) = "Other" Then
evalother = otherval
Else
evalother = currentval
End If

End Function
