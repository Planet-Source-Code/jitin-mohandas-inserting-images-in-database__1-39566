Attribute VB_Name = "Module1"
Option Explicit
Public objDB
Public lkey
Public rsImage
Public gsFileName As String
Public gsDrive As String
Public gsPath As String
Public gErrFormName As String

Sub main()
    On Error GoTo main_Error
    Set objDB = New ADODB.Connection
    Dim sPath As String
     
    Dim str As String
    sPath = App.Path
    str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & "\db1.mdb"
    
    objDB.Open str
    
    testimage.Show
    
    
Exit_main:
        
   Exit Sub
    
main_Error:
    
    #If gnDebug Then
        Stop
        Resume
    #End If

    HandleError "main", Err.Description, Err.Number, gErrFormName
    Resume Exit_main
    
End Sub

Sub HandleError(strLoc As String, strError$, lError As Long, varModule As Variant)

    Dim nCursorType As Integer

    nCursorType = Screen.MousePointer

    Screen.MousePointer = vbNormal
    MsgBox strLoc & ": " & strError & " (" & lError & ")", vbExclamation, varModule
    Screen.MousePointer = nCursorType

End Sub


