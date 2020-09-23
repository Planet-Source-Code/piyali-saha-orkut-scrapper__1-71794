VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Font Style"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2550
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Style"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Powered By Sou"
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Arial"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Private Sub Command1_Click()
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
   CommonDialog1.ShowFont
   s = CommonDialog1.FontName
   text1.Font.Name = CommonDialog1.FontName
   text1.Font.Size = CommonDialog1.FontSize
   text1.Font.Bold = CommonDialog1.FontBold
   text1.Font.Italic = CommonDialog1.FontItalic
   text1.Font.Underline = CommonDialog1.FontUnderline
   text1.FontStrikethru = CommonDialog1.FontStrikethru
   text1.ForeColor = CommonDialog1.Color
   text1.Caption = s
Form2.Text5.Text = s
Form2.Text6.Text = Form2.Text2.Text + Form2.Text5.Text + Form2.Text4.Text
Exit Sub
ErrHandler:
   Exit Sub
End Sub
Private Sub Command2_Click()
Form2.text1.Text = Form2.text1.Text + Form2.Text6.Text
Form1.Hide
End Sub
