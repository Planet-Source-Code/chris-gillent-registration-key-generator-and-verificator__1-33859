VERSION 5.00
Begin VB.Form frmKeyGenerator 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Generator - Carriage Return software"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "keyGenerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "keyGenerator.frx":030A
   ScaleHeight     =   6450
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmTimer 
      Left            =   120
      Top             =   5760
   End
   Begin VB.CommandButton btnStopGenerate 
      BackColor       =   &H0080C0FF&
      Caption         =   "Stop"
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtKeyCopy 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5280
      Width           =   4935
   End
   Begin VB.CommandButton btnGenerateKeys 
      BackColor       =   &H0080C0FF&
      Caption         =   "Generate Keys"
      Height          =   255
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ListBox lstGeneratedKeys 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2730
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   4935
   End
   Begin VB.TextBox txtKeyNumber2 
      Alignment       =   2  'Center
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      ToolTipText     =   "number between 5 and 500"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtKeyNumber1 
      Alignment       =   2  'Center
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "number between 5 and 500"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtKeyLength 
      Alignment       =   2  'Center
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "number between 5 and 25"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cmbSoftwareList 
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "copyright (C) 2002 - Carriage Return software"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4755
      Width           =   4935
   End
   Begin VB.Label lblNumberOfGeneratedKeys 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "x keys generated"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Key length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   4935
   End
End
Attribute VB_Name = "frmKeyGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#' Copyright (C) 2002 - Carriage Return software
'#'
'#' This Key Generator is made to help protect your shareware applications
'#' with an easy Registration Key system. Although very simple, the protection
'#' system is not that easy to figure out.
'#'
'#' Key Generator creates (alpha-numeric) Keys that your
'#' application will verify. Each application you create can have a different
'#' coding scheme.
'#'
'#' The principle is that each character composing the key is transformed into
'#' is Ascii value (from 0 to 255) and these values are summed up. You also supply
'#' two number ranging between 5 and 500. The sum calculated will be divided by
'#' those two numbers. If both divisions give entire results (which means a modal
'#' division with 0 as result), the key is accepted.
'#'
'#' For example, the key  y9vMP9ifC619ouul3PjMn4Fsh  gives a total of 2162.
'#' If the two numbers you specify are 47 and 23 see what happens:
'#'   2162 / 47 = 46 with no rest
'#'   2162 / 23 = 94 with no rest
'#' This key is accepted.
'#'
'#' But, take the key  sQuExEebT4ObTiG8H9wNLBnHu  which gives a total of 2166
'#' and the same numbers 47 and 23.
'#'  2166 / 47 gives a rest of 4
'#'  2166 / 23 gives a rest of 4 as well.
'#' The key is rejected.
'#'
'#' So, in your application, you just provide a way for the user to enter a
'#' registration key and you include a routine to verify it (see Sample.vbp)
'#'
'#' NOTE some numbers simply don't match together and you cannot create keys with
'#' them. After 10 seconds, key Generator will tell you it can't create keys with
'#' the specified data. Usually, changing one or both numbers and/or the key length
'#' wil help. Try using numbers smaller than 100 and key lengths between 15 and 25
'#' for best results.
'#'
'#' You can store the key data for your different applications in a plain text file
'#' named keyGenerator.dat and placed in the same directory than the source code
'#' /compiled exe. The format is (one line per application):
'#'     nom de l'application@key_data
'#' example:
'#'     Sample application version 2.1@25078064
'#' (see sample file + rest of the code for details about the structure of the key data)
'#'

Option Explicit

Dim appDirectory As String
Dim varStop As Boolean
Dim timeElapsed As Long

Const keySigns = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"


Private Sub btnGenerateKeys_Click()
Dim keyLen, keyNumber1, keyNumber2

'#' let's see if everything is there first
keyLen = txtKeyLength
keyNumber1 = txtKeyNumber1
keyNumber2 = txtKeyNumber2

If Not (keyLen <> "" And keyNumber1 <> "" And keyNumber2 <> "" And IsNumeric(keyLen) And IsNumeric(keyNumber1) And IsNumeric(keyNumber2)) Then
    '#' snap user's fingers
    MsgBox "One of the key data is missing or is not numeric"
    Exit Sub
Else
    '#' is range correct?? length must be between 5 and 25 and the two numbers
    '#' between 5 and 500
    If keyLen < 5 Or keyLen > 500 Then
        MsgBox "The key length must be a number between 5 and 25"
        Exit Sub
    ElseIf (keyNumber1 < 5 Or keyNumber1 > 500) Or (keyNumber2 < 5 Or keyNumber2 > 500) Then
        MsgBox "The two numbers must be a number between 5 and 500"
        Exit Sub
    End If
    
    '#' everything ok, let's disable and make visible what needs to be
    lstGeneratedKeys.Clear
    txtKeyCopy = ""
    
    Shape2.Visible = True
    lblNumberOfGeneratedKeys.Visible = True
    lblNumberOfGeneratedKeys.Caption = ""
    btnStopGenerate.Visible = True
    
    cmbSoftwareList.Enabled = False
    btnGenerateKeys.Enabled = False
    txtKeyLength.Enabled = False
    txtKeyNumber1.Enabled = False
    txtKeyNumber2.Enabled = False
        
    '#' and let's start generating
    GenerateKeys keyLen, keyNumber1, keyNumber2
    
End If
End Sub

Private Sub btnStopGenerate_Click()
varStop = True
End Sub

Private Sub cmbSoftwareList_Click()
Dim keyData, keyLen, keyNumber1, keyNumber2

If cmbSoftwareList.ListIndex <> -1 Then
    '#' get the key data from the itemData property, process it and
    '#' update the input boxes with it
    keyData = cmbSoftwareList.ItemData(cmbSoftwareList.ListIndex)
    
    '#' the Key data is a number that can be decomposed as this:
    '#'   the number of millions is the Key length (ex: 23287123 is Key length of 23)
    '#'   the number of thousands is the value of Number#1 (in the same example, Number#1 = 287)
    '#'   what remains is the value of Number#2 (in the same example, Number#2 = 123)
    keyLen = (keyData - (keyData Mod 1000000)) / 1000000
    keyNumber2 = keyData Mod 1000
    keyNumber1 = (keyData - (keyLen * 1000000) - keyNumber2) / 1000
    
    txtKeyLength = keyLen
    txtKeyNumber1 = keyNumber1
    txtKeyNumber2 = keyNumber2

End If
End Sub

Private Sub Form_Load()

'#' We'll first see if we can find a file named "keyGenerator.dat" in the
'#' application directory . If so, read it and load the pre-defined Key data
'#' into the combobox.

Dim datLine, keySoftwareName, keyData, atSign


'#' The timer will be used to calculate the time needed to generate the first
'#' key. If this time is greater than 10 seconds, it's almost sure that the
'#' two numbers selected cannot be used together. If this happens, try changing
'#' one of those two or both numbers and/or the key length
tmTimer.Interval = 1000
tmTimer.Enabled = False

appDirectory = App.Path & "\"
If Dir(appDirectory & "keyGenerator.dat", vbNormal) <> "" Then
    '#' clear the combobox
    cmbSoftwareList.Clear
    '#' Yes, it's there, open it
    Open appDirectory & "keyGenerator.dat" For Input As #1
        '#' and read each line
        Do Until EOF(1)
            Line Input #1, datLine
            '#' now check that the line is not empty and that the data separator
            '#' (the @ sign) is there.
            If datLine <> "" Then
                atSign = InStr(datLine, "@")
                '#' also, if line begins with a ', ignore it cause it's a comment
                If atSign <> 0 And Left(datLine, 1) <> "'" Then
                    '#' the Software name is on the left of the @sign,
                    '#' the Key data on the right
                    keySoftwareName = Left(datLine, atSign - 1)
                    keyData = Right(datLine, Len(datLine) - atSign)
                    '#' let's verify keyData is a number
                    If IsNumeric(keyData) = True Then
                        '#' ok, add to the combobox
                        cmbSoftwareList.AddItem keySoftwareName
                        '#' the fact that the whole Key data is contained into one
                        '#' numeric value is helpfull: we can store it in the
                        '#' itemdata property of the combobox items for later
                        '#' retrieval.
                        cmbSoftwareList.ItemData(cmbSoftwareList.NewIndex) = keyData
                        
                    End If
                End If
            End If
        Loop
    Close #1
End If

End Sub

Public Sub GenerateKeys(vKeyLen, vKN1, vKN2)
Dim generated, counter, newKey, varDE, varMsg

'#' initialize the random generator
Randomize Timer

'#' let's start the timer
tmTimer.Enabled = True
timeElapsed = 0

varStop = False
generated = 0
varMsg = ""

Do
    '#' initialize the new key
    newKey = String(vKeyLen, " ")
    '#' and fill it with as many random characters (taken from the 'keySigns' constant)
    '#' as its length requires
    For counter = 1 To vKeyLen
        Mid(newKey, counter, 1) = Mid(keySigns, Int(Rnd * Len(keySigns)) + 1, 1)
    Next counter
    
    Do
        '#' change one character anywhere in the newly generated key
        '#' this will serve, if the two numbers formula does not match, to modify
        '#' the key until it actually matches
        Mid(newKey, Int(Rnd * vKeyLen) + 1, 1) = Mid(keySigns, Int(Rnd * Len(keySigns)) + 1, 1)
    
        '#' now, verify if the generated key matches the two numbers formula
        If checkFormula(newKey, vKN1, vKN2) = True Then
            '#' add it to the list
            lstGeneratedKeys.AddItem newKey
            generated = generated + 1
            lblNumberOfGeneratedKeys = generated & " keys generated"
            '#' this key is accepted. exit this loop and let's start generating
            '#' the next key
            Exit Do
        End If
        '#' frees up the CPU, so that other programs can work as well
        varDE = DoEvents()
        
        '#' Check the time elapsed since the start if no key has been generated yet.
        '#' If this time is greater than 10 seconds, it's almost sure that the
        '#' two numbers selected cannot be used together. If this happens, try changing
        '#' one of those two or both numbers and/or the key length
        If generated = 0 And timeElapsed > 10 Then
            varStop = True
            varMsg = "No key could be created within the first 10 seconds. Try changing one of the two numbers or both of them and/or the key length."
        End If
        
        If varStop = True Then Exit Do
    Loop
    varDE = DoEvents()
    If varStop = True Then Exit Do
Loop 'Until generated = 100

If varStop = True Then
    tmTimer.Enabled = False
    Shape2.Visible = False
    lblNumberOfGeneratedKeys.Visible = False
    lblNumberOfGeneratedKeys.Caption = ""
    btnStopGenerate.Visible = False

    cmbSoftwareList.Enabled = True
    btnGenerateKeys.Enabled = True
    txtKeyLength.Enabled = True
    txtKeyNumber1.Enabled = True
    txtKeyNumber2.Enabled = True

    If varMsg <> "" Then MsgBox varMsg
End If

End Sub

Public Function checkFormula(vKey, vKN1, vKN2)
Dim counter, keyTotal
keyTotal = 0

'#' let's get the sum of the Ascii values of all characters in the key
For counter = 1 To Len(vKey)
    keyTotal = keyTotal + Asc(Mid(vKey, counter, 1))
Next counter

'#' if this sum can be divided by the two numbers without rest (modal division),
'#' the key is good
'#' example:
'#'   the sum of all Ascii values of the characters of the key is 35. Number 1
'#'   is 7 and number 2 is 5.  35 divided by 7 equals 5 without rest.  35 divided
'#'   by 5 equals 7 without rest. Key is OK
'#'
'#'   the sum of all Ascii values of the characters of the key is 36. Number 1
'#'   is 6 and number 2 is 8.  36 divided by 6 equals 6 without rest.  36 divided
'#'   by 8 equals 4 but has a rest of 4. Key is not good.

If keyTotal Mod vKN1 = 0 And keyTotal Mod vKN2 = 0 Then
    checkFormula = True
Else
    checkFormula = False
End If
End Function

Private Sub lstGeneratedKeys_Click()
If lstGeneratedKeys.ListIndex <> -1 Then
    txtKeyCopy = lstGeneratedKeys.List(lstGeneratedKeys.ListIndex)
Else
    txtKeyCopy = ""
End If
End Sub

Private Sub tmTimer_Timer()
timeElapsed = timeElapsed + 1
End Sub
