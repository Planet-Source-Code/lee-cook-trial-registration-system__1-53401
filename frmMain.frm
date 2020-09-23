VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Software title goes here."
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrMain 
      Interval        =   10000
      Left            =   240
      Top             =   720
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton CmdEnter 
      Caption         =   "Enter Trial"
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdKGen 
      Caption         =   "Key Generator"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CmdEntSerial 
      Caption         =   "Enter Serial"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************'
'                                             '
' SimpleTrial                                 '
' Feel free to re-distrubute this code, since '
' this code is freeware :).                   '
'                                             '
' Please vote for me.                         '
'                                             '
'*********************************************'

Dim clsDS2 As New clsDS2

Private Sub CmdAbout_Click()

    'Show details about your software.
        MsgBox "Company Name: " & App.CompanyName & vbCrLf & "Product Name: " & App.ProductName & vbCrLf & "Version: " & App.Major & "." & App.Revision & "." & App.Minor & vbCrLf & vbCrLf & "Little message about your product here.."

End Sub

Private Sub cmdEnter_Click()

    'Load the software.
        frmSoftware.Show
        
    'Add the unregistered status to the software.
        frmSoftware.Caption = "" & App.ProductName & " (Unregistered Version)"

End Sub

Private Sub cmdExit_Click()

    'Terminate the program if the user decides to.
        Unload Me

End Sub

Private Sub CmdKGen_Click()

    'Load the Key Generator form.
        frmKeyGen.Show

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Dim Line01 As String
    Dim Line02 As String
    
    'Open trial config file to check if the software is registered or not.
    Open "C:\WINDOWS\system32\hlgxu.001" For Input As #1
    
    'Grab details from config file.
    Line Input #1, Line01
    Line Input #1, Line02
    Close #1
    
    'Decrypt the text using DS2 Cipher decryption.
    Line01 = clsDS2.DecryptString(Line01, "589501068402658", True)
    Line02 = clsDS2.DecryptString(Line02, "589501068402658", True)
    
    'Check to see if the text matches a valid registration code.
    If KeyGen(Line01, "589501026005156", 3) = Line02 Then frmSoftware.Show: Unload Me
    
End Sub

Private Sub CmdEntSerial_Click()

    'Load the details entry form.
        EnterDetails.Show
End Sub

Private Sub TmrMain_Timer()
    
    'Delay the enter button.
        CmdEnter.Enabled = True

End Sub
