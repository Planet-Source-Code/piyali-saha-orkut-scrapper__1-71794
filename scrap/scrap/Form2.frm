VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smart Scrapper"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   480
      TabIndex        =   34
      Text            =   "<font size=""9000"">"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Top             =   8400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      TabIndex        =   30
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6600
      TabIndex        =   29
      Text            =   """>"
      Top             =   8400
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6960
      TabIndex        =   28
      Text            =   "<font style=""text-decoration:blink"">"
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      TabIndex        =   27
      Text            =   "<Font Face="""
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear                 Text"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   25
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      MaxLength       =   1024
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   3120
      Width           =   9375
   End
   Begin VB.Image Image14 
      Height          =   435
      Left            =   9240
      Picture         =   "Form2.frx":08CA
      ToolTipText     =   "Force To End"
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Big"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   33
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Blink"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   32
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Text Style"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   26
      Top             =   120
      Width           =   1035
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   3120
      Picture         =   "Form2.frx":0D25
      Top             =   600
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   3120
      Picture         =   "Form2.frx":126B
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yellow"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Violet"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teal"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   1800
      Width           =   405
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Silver"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   19
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purple"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   17
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pink"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orange"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Olive"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Navy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maroon"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   12
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lime"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gray"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gold"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fuchsia"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aqua"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ITALIC"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "BOLD"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ITALIC"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "UNDERLINE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BOLD"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   510
   End
   Begin VB.Image Image13 
      Height          =   300
      Left            =   7800
      Picture         =   "Form2.frx":17B1
      Top             =   480
      Width           =   300
   End
   Begin VB.Image Image12 
      Height          =   270
      Left            =   6360
      Picture         =   "Form2.frx":1B66
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image11 
      Height          =   300
      Left            =   7800
      Picture         =   "Form2.frx":1EBB
      Top             =   120
      Width           =   345
   End
   Begin VB.Image Image10 
      Height          =   300
      Left            =   6840
      Picture         =   "Form2.frx":221D
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   7320
      Picture         =   "Form2.frx":25CA
      Top             =   120
      Width           =   315
   End
   Begin VB.Image Image8 
      Height          =   270
      Left            =   5880
      Picture         =   "Form2.frx":2982
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image7 
      Height          =   315
      Left            =   8280
      Picture         =   "Form2.frx":2CE1
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Image6 
      Height          =   315
      Left            =   8280
      Picture         =   "Form2.frx":3081
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   7320
      Picture         =   "Form2.frx":3433
      Top             =   480
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   0
      Picture         =   "Form2.frx":37C1
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   9375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   960
      Picture         =   "Form2.frx":4CBA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form3.Show
End Sub
Private Sub Command3_Click()
text1.Text = ""
End Sub
Private Sub Form_Load()
Me.Caption = "Smart Scrapper By Sou"
On Error Resume Next
Me.Show
              Me.Refresh
              With nid
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon 'use form's icon in system tray
            .szTip = "Smart Scrapper Powered By Sou" & vbNullChar 'tooltip text
        End With
    Shell_NotifyIcon NIM_ADD, nid 'add to tray
Label4.Visible = False
Label2.Visible = False
Label5.Visible = False
Image3.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
For intCtr = (Forms.Count - 1) To 0 Step -1
Unload Forms(intCtr)
Next intCtr
End
End Sub
Private Sub Image1_Click()
text1.Text = text1.Text + "<STRIKE>"
Image1.Visible = False
Image3.Visible = True
End Sub
Private Sub Image10_Click()
text1.Text = text1.Text + "[;)]"
End Sub
Private Sub Image11_Click()
text1.Text = text1.Text + "[:o]"
End Sub
Private Sub Image12_Click()
text1.Text = text1.Text + "[:(]"
End Sub
Private Sub Image13_Click()
text1.Text = text1.Text + "[:D]"
End Sub
Private Sub Image14_Click()
For intCtr = (Forms.Count - 1) To 0 Step -1
Unload Forms(intCtr)
Next intCtr

free = Space(10485760)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub
Private Sub Image2_Click()
text1.Text = text1.Text + "[u]"
Image2.Visible = False
Label2.Visible = True
End Sub
Private Sub Image3_Click()
text1.Text = text1.Text + "</STRIKE>"
Image3.Visible = False
Image1.Visible = True
End Sub
Private Sub Image5_Click()
text1.Text = text1.Text + "[:x]"
End Sub
Private Sub Image6_Click()
text1.Text = text1.Text + "[:p]"
End Sub
Private Sub Image7_Click()
text1.Text = text1.Text + "[/)]"
End Sub
Private Sub Image8_Click()
text1.Text = text1.Text + "[8)]"
End Sub
Private Sub Image9_Click()
text1.Text = text1.Text + "[:)]"
End Sub
Private Sub Label1_Click()
text1.Text = text1.Text + "[b]"
Label4.Visible = True
Label1.Visible = False
End Sub
Private Sub Label10_Click()
text1.Text = text1.Text + "[Gray]"
End Sub
Private Sub Label11_Click()
text1.Text = text1.Text + "[green]"
End Sub
Private Sub Label12_Click()
text1.Text = text1.Text + "[lime]"
End Sub
Private Sub Label13_Click()
text1.Text = text1.Text + "[maroon]"
End Sub
Private Sub Label14_Click()
text1.Text = text1.Text + "[navy]"
End Sub
Private Sub Label15_Click()
text1.Text = text1.Text + "[olive]"
End Sub
Private Sub Label16_Click()
text1.Text = text1.Text + "[orange]"
End Sub
Private Sub Label17_Click()
text1.Text = text1.Text + "[pink]"
End Sub
Private Sub Label18_Click()
text1.Text = text1.Text + "[purple]"
End Sub
Private Sub Label19_Click()
text1.Text = text1.Text + "[red]"
End Sub
Private Sub Label2_Click()
text1.Text = text1.Text + "[/u]"
Image2.Visible = True
Label2.Visible = False
End Sub
Private Sub Label20_Click()
text1.Text = text1.Text + "[silver]"
End Sub
Private Sub Label21_Click()
text1.Text = text1.Text + "[teal]"
End Sub
Private Sub Label22_Click()
text1.Text = text1.Text + "[violet]"
End Sub
Private Sub Label23_Click()
text1.Text = text1.Text + "[yellow]"
End Sub
Private Sub Label24_Click()
text1.Text = Text3.Text + text1.Text
Label24.Visible = False
End Sub
Private Sub Label25_Click()
text1.Text = text1.Text + Text7.Text
End Sub
Private Sub Label26_Click()
Form1.Show
End Sub
Private Sub Label3_Click()
text1.Text = text1.Text + "[i]"
Label5.Visible = True
Label3.Visible = False
End Sub
Private Sub Label4_Click()
text1.Text = text1.Text + "[/b]"
Label1.Visible = True
Label4.Visible = False
End Sub
Private Sub Label5_Click()
text1.Text = text1.Text + "[/i]"
Label3.Visible = True
Label5.Visible = False
End Sub
Private Sub Label6_Click()
text1.Text = text1.Text + "[aqua]"
End Sub
Private Sub Label7_Click()
text1.Text = text1.Text + "[blue]"
End Sub
Private Sub Label8_Click()
text1.Text = text1.Text + "[fuchsia]"
End Sub
Private Sub Label9_Click()
text1.Text = text1.Text + "[gold]"
End Sub
