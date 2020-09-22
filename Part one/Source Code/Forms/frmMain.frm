VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Command Line Example"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Keep the variables controlled
Option Explicit

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
' Here are the 3 easy steps to be able to use this program:
'   1) Compile and make it a .exe file.
'   2) Place it in your Windows folder (base directory for
'      "Run", DOS and shortcuts).
'   3) Launch the program through "Run", through DOS or
'      through creating a shortcut.
'
' The format for this program to function correctly is:
'   CL <MESSAGE>
'
' Here is an example of how it could be used:
'   CL Welcome to this program, I hope you enjoy it...
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub Form_Load()

    'Dim the strParas String (String Parameters)
    'This string will store the information passed through Command
    Dim strParas As String
    
    'Set strParas value to the current "Command" value
    strParas = Command

    'Display the Parameters contained in the given Command Line
    lblWelcome.Caption = strParas

    'Resize the Form to fit the dimensions of the Label (lblWelcome)
    'This code is not required, though it makes the program nicer looking
    Me.Width = lblWelcome.Width + 360

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Securely unload and terminate the program
    Unload Me
    End

End Sub
