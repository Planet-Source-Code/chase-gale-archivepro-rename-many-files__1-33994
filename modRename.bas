Attribute VB_Name = "modRename"
Option Explicit

Public Sub RenameFiles(ByVal vDir As Variant, FileName As String, Pad As Integer, _
    StartNumber As Integer, TypeOfFiles As String)

Dim vFile As Variant
Dim Padded As String
Dim OrigPadded As String
Dim FullFile As String
Dim NewFile As String
Dim FileExt As String

Select Case Pad
    Case 0
        OrigPadded = ""
    Case 1
        OrigPadded = "0"
    Case 2
        OrigPadded = "00"
    Case 3
        OrigPadded = "000"
    Case 4
        OrigPadded = "0000"
End Select

Padded = OrigPadded
vFile = Dir(vDir, vbDirectory)

If vFile = "" Then
    Exit Sub
End If

vFile = Dir(vDir & "\", vbDirectory)

Do Until vFile = ""
    If vFile = "." Or vFile = ".." Then
        vFile = Dir
    ElseIf (GetAttr(vDir & "\" & vFile) And vbDirectory) = vbDirectory Then
        vFile = Dir
    Else
    
        Select Case Len(OrigPadded)
        Case 1
            If StartNumber = 10 Then Padded = ""
        Case 2
            If StartNumber = 10 Then Padded = "0"
            If StartNumber = 100 Then Padded = ""
        Case 3
            If StartNumber = 10 Then Padded = "00"
            If StartNumber = 100 Then Padded = "0"
            If StartNumber = 1000 Then Padded = ""
        Case 4
            If StartNumber = 10 Then Padded = "000"
            If StartNumber = 100 Then Padded = "00"
            If StartNumber = 1000 Then Padded = "0"
            If StartNumber = 10000 Then Padded = ""
        End Select
        
        FullFile = vDir & "\" & vFile
        FileExt = Right(vFile, 3)
        NewFile = vDir & "\" & FileName & Padded & StartNumber & "." & FileExt
        
        Select Case TypeOfFiles
            Case "All"
                Name FullFile As NewFile
                StartNumber = StartNumber + 1
            Case "Graphic"
                If (FileExt = "jpg") Or (FileExt = "gif") Or (FileExt = "bmp") Then
                    Name FullFile As NewFile
                    StartNumber = StartNumber + 1
                End If
            Case "Text"
                If (FileExt = "txt") Or (FileExt = "doc") Then
                    Name FullFile As NewFile
                    StartNumber = StartNumber + 1
                End If
        End Select
        
        vFile = Dir
        frmMain.Bar.Value = frmMain.Bar.Value + 1
    End If
Loop

frmMain.Bar.Value = 0
End Sub
