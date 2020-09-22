VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sample application - Test your Keys"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test it!!"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"sample.frx":0000
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your registration key"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number1, Number2

Private Sub Command1_Click()
Dim vKey
If Text1 <> "" Then
    vKey = Text1
    If VerifyKey(vKey) = True Then
        MsgBox "Key accepted", vbInformation
    Else
        MsgBox "Key rejected", vbCritical
    End If
End If
End Sub

Private Sub Form_Load()

'#' When the user has successfully entered the registration key, your application
'#' should store it someplace (registry, maybe?) and verify it each time it starts.

'#' change if you want to test other numbers
Number1 = 97
Number2 = 17

End Sub


Public Function VerifyKey(vKey)
Dim counter, keyTotal

'#' sum up the ascii values of all characters of the key
keyTotal = 0
For counter = 1 To Len(vKey)
    keyTotal = keyTotal + Asc(Mid(vKey, counter, 1))
Next counter

'#' verify against the two numbers. See Key Generator code for more details
If keyTotal <> 0 And keyTotal Mod Number1 = 0 And keyTotal Mod Number2 = 0 Then
    VerifyKey = True
Else
    VerifyKey = False
End If

End Function

