VERSION 5.00
Begin VB.Form MagicSquare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magic Square"
   ClientHeight    =   6285
   ClientLeft      =   -2385
   ClientTop       =   180
   ClientWidth     =   9660
   Icon            =   "MagicSquare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   26.188
   ScaleMode       =   4  'Character
   ScaleWidth      =   80.5
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Magic Aquare"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "MagicSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This code was created by me Philip V. Naparan. I upload this
'code to PSC to gave a simple idea to those beginner in VB
'to create a simple analytical program that uses two dimensional
'Arrays
'I CREATED THIS APPLICATION IN JUST 13 MINUTES
'**************************************************************

'i use double so that the user can enter number without a limit

Option Explicit

Dim cols(), cNumber, counter, nArray, xCounter As Double
Dim hN As String
Dim cols1(0, 0) As String

Private Sub Command1_Click()
If Val(Text1.Text) Mod 2 = 0 Then
    MsgBox "Pls. Enter a valid even number", vbCritical, "Confirm"
    Text1.SetFocus
    Exit Sub
End If
Dim currentCols, currentRows As Double
Dim distCOLS, distROWS As Double

ReDim cols(Val(Text1.Text), Val(Text1.Text))
cNumber = (Val(Text1.Text) \ 2) + Val(Text1.Text) Mod 2
nArray = Val(Text1.Text)

distROWS = 1
distCOLS = cNumber
Me.Cls
For counter = 1 To (Val(Text1.Text) * Val(Text1.Text))
    cols(distROWS, distCOLS) = counter
    
    currentRows = distROWS
    currentCols = distCOLS
    
    distROWS = distROWS - 1
    If distROWS < 1 Then
        distROWS = nArray
    End If
    distCOLS = distCOLS + 1
    If distCOLS > nArray Then
        distCOLS = 1
    End If
    If cols(distROWS, distCOLS) <> 0 Then
        distROWS = currentRows + 1
        If distROWS > nArray Then distROWS = distROWS - 1
        distCOLS = currentCols
    End If
    
Next counter
Print
Print
Print
Print
For counter = 1 To nArray
    hN = ""
    For xCounter = 1 To nArray
        hN = hN & vbTab & cols(counter, xCounter)
    Next xCounter
    Print hN
Next counter
End Sub
Private Sub Command2_Click()
frmAbout.Show vbModal
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
