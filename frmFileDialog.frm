VERSION 5.00
Begin VB.Form frmFileDialog 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   2100
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   4800
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
gsFileName = ""
Unload Me
End Sub

Private Sub cmdok_Click()
gsFileName = txtfile
Unload Me
End Sub

Private Sub Dir1_Change()
txtfile.Text = Dir1.Path
File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
gsDrive = Drive1.Drive


End Sub

Private Sub File1_Click()
txtfile = File1.Path + "\" + File1.FileName

End Sub
