VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom URL Protocol"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Launch 
      Caption         =   "Launch Sample Page"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Text            =   $"Form1.frx":0442
      Top             =   7320
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You need to compile this program in order to work correctly."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   810
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   2820
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Sub Form_Load()
Open "c:\test.html" For Output As 1     ' this is just to create the sample HTML page
Print #1, "" + Text4.Text + ""          ' this is just to create the sample HTML page
Close 1                                 ' this is just to create the sample HTML page
End Sub

Private Sub Launch_Click()
'execute the sample page
ShellExecute Me.hwnd, vbNullString, "c:\test.html", vbNullString, "C:\", SW_SHOWNORMAL
End Sub
