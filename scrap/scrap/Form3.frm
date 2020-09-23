VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Me"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   10905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   4080
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "http://www.orkut.com/Profile.aspx?uid=4091101399868562695"
      Top             =   2640
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contact Me "
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Text            =   "http://www.orkut.com/Profile.aspx?uid=17490378593619702546"
         Top             =   3240
         Width           =   6495
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "lovesourav@yahoo.com"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "sourav17ghosh@rediffmail.com"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "sourav18ghosh@gmail.com"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "sourav17ghosh@gmail.com"
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Sourav Ghosh"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Image Image3 
      Height          =   3495
      Left            =   6840
      Picture         =   "Form3.frx":08CA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4020
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   3750
      Left            =   7080
      Picture         =   "Form3.frx":4FFD
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3930
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4245
      Left            =   7080
      Picture         =   "Form3.frx":A20E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3720
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
End Sub
Private Sub Timer1_Timer()
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Timer2.Enabled = True
Timer1.Enabled = False
Timer3.Enabled = False
End Sub
Private Sub Timer2_Timer()
Image2.Visible = True
Image1.Visible = False
Image3.Visible = False
Timer3.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
End Sub
Private Sub Timer3_Timer()
Image3.Visible = True
Image1.Visible = False
Image2.Visible = False
Timer1.Enabled = True
Timer1.Enabled = False
Timer3.Enabled = False
End Sub
