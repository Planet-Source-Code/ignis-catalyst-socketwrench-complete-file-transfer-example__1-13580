VERSION 5.00
Begin VB.Form frmBrowseDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse Directory"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2490
      TabIndex        =   3
      Top             =   3135
      Width           =   915
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   330
      Left            =   3450
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.DirListBox Directory 
      Height          =   2565
      Left            =   90
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4335
   End
End
Attribute VB_Name = "frmBrowseDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChoose_Click()
    'Add and select the new path
    frmMain.cmbRecvDirectory.AddItem Directory.Path
    frmMain.cmbRecvDirectory.Text = Directory.Path
    
    'Close dialog
    Unload Me
End Sub

Private Sub Drive_Change()
    'Change the Directory to reflect the Drive
    On Error GoTo HandleError
    Me.MousePointer = vbHourglass
    Directory.Path = Drive.Drive
    Me.MousePointer = vbNormal
    Exit Sub
    
HandleError:
    MsgBox "Error: Either the drive is inaccessible, or drive is corrupt.", vbExclamation, "Drive Error"
    Drive.Drive = "C:"
    Me.MousePointer = vbNormal
    Exit Sub
End Sub


