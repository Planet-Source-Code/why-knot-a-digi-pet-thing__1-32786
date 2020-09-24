VERSION 5.00
Begin VB.Form frmDigiPet 
   Caption         =   "Digital Pet"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "Start New Pet"
      Height          =   2775
      Left            =   -240
      TabIndex        =   10
      Top             =   -240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Play 
      Interval        =   20000
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Clean 
      Interval        =   8000
      Left            =   960
      Top             =   120
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "Clean"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer Food 
      Interval        =   12500
      Left            =   1200
      Top             =   120
   End
   Begin VB.CommandButton cmdFood 
      Caption         =   "Feed"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1080
      ScaleHeight     =   1335
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblExercise 
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblHealth 
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblHunger 
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Exercise:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Health:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Hunger:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "frmDigiPet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Hunger As Integer
Dim Health As Integer
Dim Exercise As Integer
Dim Death As Variant

Private Sub Clean_Timer()
If Health > 0 Then
Health = Health - 1
Else
Death = True
End If
End Sub

Private Sub cmdClean_Click()
If Health < 20 And Health > 0 Then
Health = Health + 1
Else
Health = Health
End If
End Sub

Private Sub cmdFood_Click()
If Hunger < 20 And Hunger > 0 Then
Hunger = Hunger + 1
Else
Hunger = Hunger
End If
End Sub

Private Sub cmdNew_Click()
cmdFood.Enabled = True
cmdClean.Enabled = True
cmdPlay.Enabled = True
lblExercise.Visible = True
lblHealth.Visible = True
lblHunger.Visible = True
cmdNew.Visible = True
Picture1.Picture = LoadPicture(App.Path + "\Happy.BMP")
Hunger = 20
Health = 20
Exercise = 20
lblHunger = Hunger
lblExercise = Exercise
lblHealth = Health
cmdNew.Visible = False
Death = False
End Sub

Private Sub cmdPlay_Click()
If Exercise < 20 And Exercise > 0 Then
Exercise = Exercise + 1
Else
Exercise = Exercise
End If
End Sub

Private Sub Food_Timer()
If Hunger > 0 Then
Hunger = Hunger - 1
Else
Death = True
End If
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture(App.Path + "\Happy.BMP")
Hunger = 20
Health = 20
Exercise = 20
End Sub

Private Sub Play_Timer()
If Exercise > 0 Then
Exercise = Exercise - 1
Else
Death = True
End If
End Sub

Private Sub Timer1_Timer()
If Death = True Then
cmdFood.Enabled = False
cmdClean.Enabled = False
cmdPlay.Enabled = False
lblExercise.Visible = False
lblHealth.Visible = False
lblHunger.Visible = False
cmdNew.Visible = True
End If
If Health > 10 Or Hunger > 10 Or Exercise > 10 Then
Picture1.Picture = LoadPicture(App.Path + "\Happy.BMP")
End If
If Health < 11 And Health > 4 Or Hunger < 11 And Hunger > 4 Or Exercise < 11 And Exercise > 4 Then
Picture1.Picture = LoadPicture(App.Path + "\Normal.BMP")
End If
If Health < 5 And Health > 0 Or Hunger < 5 And Hunger > 0 Or Exercise < 5 And Exercise > 0 Then
Picture1.Picture = LoadPicture(App.Path + "\Upset.BMP")
End If
If Health = 0 Or Hunger = 0 Or Exercise = 0 Then
Picture1.Picture = LoadPicture(App.Path + "\Dead.BMP")
End If
lblHunger = Hunger
lblExercise = Exercise
lblHealth = Health
End Sub
