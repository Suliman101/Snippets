Attribute VB_Name = "updatemod"
Option Compare Database

Public Function updater(vername As Variant, vervar As Variant)
Dim pname As String
Dim pdir As String
Dim pfile As String
Dim newver As Double
Dim fupdate As Boolean
Dim messagestr As String

'find version information from PUT db


If ObjectExists("Table", "kdgimport_updater") Then
    DoCmd.DeleteObject acTable, "kdgimport_updater"
End If

If Dir("\\NWATER-SERVER1\Apps\DATA\Kapil\Projects\2014 projects\Project Update Tracker\PUT.accdb") <> "" Then
    DoCmd.TransferDatabase acImport, "Microsoft Access", _
    "\\NWATER-SERVER1\Apps\DATA\Kapil\Projects\2014 projects\Project Update Tracker\PUT.accdb", _
    acTable, "base", _
    "kdgimport_updater"

    pname = DLookup("pname", "kdgimport_updater", "pname = '" & vername & "'")
    pdir = DLookup("pdir", "kdgimport_updater", "pname = '" & vername & "'")
    pfile = DLookup("pfilename", "kdgimport_updater", "pname = '" & vername & "'")
    newver = CDbl(DLookup("newestbuild", "kdgimport_updater", "pname = '" & vername & "'"))
    fupdate = CBool(DLookup("forceupdate", "kdgimport_updater", "pname = '" & vername & "'"))
    messagestr = DLookup("messagestr", "kdgimport_updater", "pname = '" & vername & "'")
Else
MsgBox "local"
    Exit Function
End If

If newver > vervar And fupdate = False Then
    Dim LResponse As Integer
    
    LResponse = MsgBox("New version of " & pname & " found: " & "V. " & newver & _
        " Do you wish to update?", vbYesNo, "Continue")
End If

If LResponse = vbYes Or (newver > vervar And fupdate = True) Then
        MsgBox "New version will be downloaded: " & pname & " Version " & newver & _
        vbNewLine & messagestr
        Dim filenamevar As String
        'Dim verid As String
        'verid = "Build=0.1.2" & vbCrLf & "Force Update=-1" & vbCrLf & "Update Source=test" & vbCrLf & "Current Source=" & CurrentProject.FullName
        
        filenamevar = CurrentProject.Path & "\firstrun.kdg"
        Dim FILENAME
        FILENAME = filenamevar
        Dim My_filenumber As Integer
        My_filenumber = FreeFile
        Open FILENAME For Output As #My_filenumber
        Print #My_filenumber, CurrentProject.FullName
        Close #My_filenumber
            
        Dim SourceFile, DestinationFile As String
        SourceFile = pdir  ' Define source file name.
        DestinationFile = CurrentProject.Path & "\" & pname & " " & newver & ".accdb" ' Define target file name.
        FileCopy SourceFile, DestinationFile  ' Copy source to target.
            
        Application.FollowHyperlink DestinationFile
            
        DoCmd.Quit
End If


End Function

Public Function getdata(var1 As Variant)
Dim lenvar As Integer
lenvar = InStr(1, var1, "=") + 1
getdata = Mid(var1, lenvar, Len(var1) - lenvar + 1)
End Function

Public Function virginapp()

End Function

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function
