VERSION 5.00
Begin VB.Form EnterDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your software title goes here."
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton CmdReg 
      Caption         =   "&Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.TextBox txtGen 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Username:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Registration Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "EnterDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDS2 As New clsDS2

'*********************************************'
'                                             '
' SimpleTrial                                 '
' Feel free to re-distrubute this code, since '
' this code is freeware :).                   '
'                                             '
' Please vote for me.                         '
'                                             '
'*********************************************'

Private Sub cmdQuit_Click()

    'Quit the form when a user decides to.
        Unload Me

End Sub

Private Sub CmdReg_Click()
    
    'Check to see if the user input matches correct information.
        If KeyGen(txtUsername, "589501026005156", 3) = txtGen Then
            MsgBox "Registration successfull, thank you for purchasing this product, you will need to re-launch this program for the changes to take effect.", vbInformation, "Registration Complete!"
    
    'Encrypt the file to stop people from looking at this hidden info.
        txtUsername.Text = clsDS2.EncryptString(txtUsername.Text, "589501068402658", True)
        txtGen.Text = clsDS2.EncryptString(txtGen.Text, "589501068402658", True)
    
    'Write the details to file, if they are correct then the software will be registered.
    Open "C:\WINDOWS\system32\hlgxu.002" For Output As #1
        Print #1, txtUsername.Text
        Print #1, txtGen.Text
    Close
    
    'Copy the temp file to the trial config file.
    FileCopy "C:\WINDOWS\system32\hlgxu.002", "C:\WINDOWS\system32\hlgxu.001"
    Kill "C:\WINDOWS\system32\hlgxu.002"
    
    Unload Me

    End If
End Sub

