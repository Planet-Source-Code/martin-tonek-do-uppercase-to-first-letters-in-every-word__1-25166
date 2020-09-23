VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "(c) 2000 martin tonek"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   Icon            =   "fixletters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "test function"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "kalle ANKA is aLIVE"
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'this string is possible to put in any of the attributes of a control.
Text1.Text = FixLetters(Text1.Text)
End Sub


Public Function FixLetters(txt As String)
    '---------------------
    ' EasyLetterFixModul
    ' (c)2001 Martin Tonek
    ' www.mrslade.com
    '---------------------
    Dim numOFletters
    'error fix if the string is empty
    If Len(txt) > 0 Then
        'do all letters small
        txt = LCase(txt)
        'checking the first letter
        Mid$(txt, 1, 1) = UCase(Mid$(txt, 1, 1))
        'start the loop for txt-string
        For numOFletters = 1 To Len(txt)
            'check if the numOFletters is a space if then make the next letter uppercase
            If Mid$(txt, numOFletters, 1) = " " Then
                Mid$(txt, numOFletters + 1, 1) = UCase(Mid$(txt, numOFletters + 1, 1))
            End If
        Next numOFletters
    End If
    FixLetters = txt
End Function

Private Sub Command2_Click()
Me.Caption = Asc(Text2.Text)
End Sub
