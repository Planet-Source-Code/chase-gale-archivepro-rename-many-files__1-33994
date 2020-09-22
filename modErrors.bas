Attribute VB_Name = "modErrors"
Option Explicit

Public Sub HandleError(ByVal CurrentModule As String, ByVal CurrentProcedure As String, _
                        ByVal ErrNum As Long, ByVal ErrDescription As String)

Select Case ErrNum
    Case 68
        MsgBox "That drive is unavailable; Try inserting a CD, disk, or restoring the network connection. Thank you.", vbOKOnly, "Whoops!"
        frmMain.DriveList.Refresh
    Case 58
        Resume Next
    Case Else
        MsgBox "Error in module: " & CurrentModule & ", Procedure: " & CurrentProcedure & ".  " _
                & "Error " & ErrNum & ": " & ErrDescription, vbCritical, "Error."
End Select
End Sub


