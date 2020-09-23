VERSION 5.00
Begin VB.Form frmDigiNums 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Digital Numbers!"
   ClientHeight    =   1710
   ClientLeft      =   5415
   ClientTop       =   3765
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      Caption         =   "... and they will be displayed here as digital numbers"
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3975
      Begin VB.PictureBox picDisplay 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3705
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      Caption         =   "Type any numbers into this Textbox ..."
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtNums 
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   9
      Left            =   2280
      Picture         =   "Form1.frx":0000
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   8
      Left            =   2040
      Picture         =   "Form1.frx":04BA
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   7
      Left            =   1800
      Picture         =   "Form1.frx":0974
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   6
      Left            =   1560
      Picture         =   "Form1.frx":0E2E
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   5
      Left            =   1320
      Picture         =   "Form1.frx":12E8
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   4
      Left            =   1080
      Picture         =   "Form1.frx":17A2
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   3
      Left            =   840
      Picture         =   "Form1.frx":1C5C
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   2
      Left            =   600
      Picture         =   "Form1.frx":2116
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":25D0
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Digits 
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":2A8A
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmDigiNums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
'* Digital Numbers Example            *
'*------------------------------------*
'* Copyright 2000 by Druid Developing *
'**************************************

'VOTE FOR THIS IF YOU LIKE IT!

Dim cNumber As Integer 'The next number to display

Private NumCount As Integer 'To count the numbers in the Textbox,

Private LoopCount As Integer 'Needed to calculate the digital number´s
                             'position

Private Sub txtNums_Change()
    'Clear the Picturebox
    picDisplay.Cls
    'Count the numbers
    NumCount = Len(txtNums.Text)
    'Initialize the LoopCount
    LoopCount = 1
    'Loop until every number has been processed
    For nCount = 1 To NumCount
        'Get the next number
        cNumber = CInt(Mid(txtNums.Text, NumCount - nCount + 1, 1))
        'Get the picture from the picture of the number
        'and paint it into the PictureBox
        Call picDisplay.PaintPicture(Digits(cNumber).Picture, picDisplay.Width - Digits(cNumber).Width * LoopCount, 0)
        LoopCount = LoopCount + 1
    Next nCount
    'Refresh the PictureBox
    picDisplay.Refresh
End Sub
