VERSION 5.00
Begin VB.Form testimage 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   795
   ClientTop       =   6285
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   10995
   Begin VB.TextBox txtpno 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txttitle 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8760
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtdesig 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6120
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Insert Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   840
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   6120
      ScaleHeight     =   2955
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3600
      TabIndex        =   0
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Reports To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "DESIGNATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "testimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const conChunkSize = 16384

Private Sub cmdadd_Click()
Dim rsImage
Dim ssql1 As String
Dim chunk() As Byte
Dim lOffset1 As Long
Dim nFragmentOffset1 As Integer
If txtName.Text = "" Or txtdesig.Text = "" Or txttitle.Text = "" Or txtpno.Text = "" Then
MsgBox "Please Enter the appropriate values in text field"
txtName.SetFocus
Exit Sub
End If
cmdload.Enabled = True
frmFileDialog.Show vbModal
If gsFileName <> "" Then
 
  

 ssql1 = "select * from images "
 Set rsImage = New ADODB.Recordset
 
 rsImage.CursorType = adOpenKeyset
 rsImage.LockType = adLockOptimistic
 rsImage.Open ssql1, objDB
 
 
 
 Set Picture1.Picture = LoadPicture(gsFileName)
 
 nHandle = FreeFile
 Open gsFileName For Binary Access Read As nHandle
 lsize = LOF(nHandle)
 If nHandle = 0 Then
 Close nHandle
 End If
 lchunks = lsize / conChunkSize
 nFragmentOffset1 = lsize Mod conChunkSize
 
 rsImage.AddNew
 
 rsImage("Designation") = txtdesig.Text
 rsImage("name") = txtName.Text
 
 rsImage("Title") = txttitle.Text
 rsImage("Phoneno") = txtpno.Text
 
 
 ReDim chunk(nFragmentOffset1)
 Get nHandle, , chunk()
 rsImage("a_image").AppendChunk chunk()
 ReDim chunk(conChunkSize)
 
 For i = 1 To lchunks
 Get nHandle, , chunk()
 rsImage("a_image").AppendChunk chunk()
 DoEvents
 Next
 rsImage.Update
 
 End If
 cmdnext.Enabled = False
 cmdprev.Enabled = False
txtName.Text = ""
txtdesig.Text = ""
txttitle.Text = ""
txtpno.Text = ""
txtName.SetFocus
 
End Sub

Private Sub cmdload_Click()

Dim fld As ADODB.Field

Dim sSQL As String
Dim lkey As Long
Dim varChunk() As Byte
Dim lOffset As Long
    Dim sPath As String
    Dim nHandle As Integer
    Dim iChunks As Integer
    Dim nFragmentOffset As Integer
    Dim i As Integer
    Dim sFile As String
    
    


        cmdnext.Enabled = True
        sSQL = "select * from images"
    
        Set rsImage = New ADODB.Recordset
        
        rsImage.CursorType = adOpenDynamic
        rsImage.LockType = adLockOptimistic
        
        rsImage.Open sSQL, objDB
        If rsImage.BOF Then
        MsgBox ("No records in Database")
        cmdnext.Enabled = False
        txtName.SetFocus
        Exit Sub
        End If
        txtName.Text = rsImage("name")
        txtdesig.Text = rsImage("Designation")
        txttitle.Text = rsImage("Title")
            txtpno.Text = rsImage("Phoneno")
        
        
        If Not rsImage.EOF Then
            nHandle = FreeFile
            sPath = App.Path
            sFile = sPath & "\output.bin"
            
            
            Open sFile For Binary Access Write As nHandle
            Set fld = rsImage("a_image")
           lsize = rsImage("a_image").ActualSize
           iChunks = lsize \ conChunkSize
           nFragmentOffset = lsize Mod conChunkSize
           
           
           varChunk() = rsImage("a_image").GetChunk(nFragmentOffset)
           Put nHandle, , varChunk()
           lOffset = nFragmentOffset
           For i = 1 To iChunks
                ReDim varChunk(conChunkSize) As Byte
                varChunk() = rsImage("a_image").GetChunk(conChunkSize)
                Put nHandle, , varChunk()
                             
                DoEvents
           Next
        
        


        
End If
Close nHandle
Set Picture1.Picture = LoadPicture(sFile, , vbLPColor)


cmdadd.Enabled = False

cmdload.Enabled = False


    
End Sub





Private Sub cmdnext_Click()
Dim varChunk() As Byte
cmdadd.Enabled = False
cmdprev.Enabled = True
rsImage.MoveNext
If rsImage.EOF Then

cmdnext.Enabled = False
cmdload.Enabled = False
txtName.Text = ""
txtdesig.Text = ""
txttitle.Text = ""
txtpno.Text = ""
rsImage.MoveLast
cmdadd.Enabled = True
cmdprev.Enabled = True
MsgBox "You have reached the last record"
Exit Sub
End If


txtName.Text = rsImage("name")
txtdesig.Text = rsImage("Designation")
txttitle.Text = rsImage("Title")
txtpno.Text = rsImage("Phoneno")
nHandle = FreeFile
            sPath = App.Path
            sFile = sPath & "\output.bin"
            
            
            Open sFile For Binary Access Write As nHandle
            Set fld = rsImage("a_image")
           lsize = rsImage("a_image").ActualSize
           iChunks = lsize \ conChunkSize
           nFragmentOffset = lsize Mod conChunkSize
           
           
           varChunk() = rsImage("a_image").GetChunk(nFragmentOffset)
           Put nHandle, , varChunk()
           lOffset = nFragmentOffset
           For i = 1 To iChunks
                ReDim varChunk(conChunkSize) As Byte
                varChunk() = rsImage("a_image").GetChunk(conChunkSize)
                Put nHandle, , varChunk()
                             
                DoEvents
           Next
       
       Close nHandle
Set Picture1.Picture = LoadPicture(sFile, , vbLPColor)

End Sub

Private Sub cmdprev_Click()
Dim varChunk() As Byte
cmdadd.Enabled = False
cmdnext.Enabled = True
rsImage.MovePrevious
If rsImage.BOF Then
rsImage.MoveFirst
cmdnext.Enabled = True
cmdload.Enabled = False
txtName.Text = ""
txtdesig.Text = ""
txttitle.Text = ""
txtpno.Text = ""
cmdadd.Enabled = True
cmdprev.Enabled = False
MsgBox "You have reached the first record"
txtName.SetFocus
Exit Sub
End If


txtName.Text = rsImage("name")
txtdesig.Text = rsImage("Designation")
txttitle.Text = rsImage("Title")
txtpno.Text = rsImage("Phoneno")
nHandle = FreeFile
            sPath = App.Path
            sFile = sPath & "\output.bin"
            
            
            Open sFile For Binary Access Write As nHandle
            Set fld = rsImage("a_image")
           lsize = rsImage("a_image").ActualSize
           iChunks = lsize \ conChunkSize
           nFragmentOffset = lsize Mod conChunkSize
           
           
           varChunk() = rsImage("a_image").GetChunk(nFragmentOffset)
           Put nHandle, , varChunk()
           lOffset = nFragmentOffset
           For i = 1 To iChunks
                ReDim varChunk(conChunkSize) As Byte
                varChunk() = rsImage("a_image").GetChunk(conChunkSize)
                Put nHandle, , varChunk()
                             
                DoEvents
           Next
       
       Close nHandle
Set Picture1.Picture = LoadPicture(sFile, , vbLPColor)

End Sub

Private Sub Form_Load()



txtName.Text = ""

cmdnext.Enabled = False
cmdadd.Enabled = True
cmdprev.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not objDB Is Nothing Then
        objDB.Close
    End If
    Set objDB = Nothing
    

  End Sub





Private Sub txtName_KeyPress(KeyAscii As Integer)
cmdadd.Enabled = True
End Sub
