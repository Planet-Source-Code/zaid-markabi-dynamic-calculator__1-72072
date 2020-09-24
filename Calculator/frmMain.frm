VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Calculator"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   1065
      TabIndex        =   35
      Top             =   3360
      Width           =   1095
      Begin VB.Label Prev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~= 0"
         Height          =   195
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.TextBox ValueText 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ctg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plank"
      Height          =   495
      Index           =   21
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sqr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton CmdButnEq 
      BackColor       =   &H00FFFFFF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton CmdButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   120
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Written by   Zaid Markabi , Arabic Syrian Programmer ."
      Height          =   195
      Left            =   480
      TabIndex        =   37
      Top             =   3720
      Width           =   3825
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdButn_Click(index As Integer)
If index < 10 Then
ValueText.Text = ValueText.Text + CmdButn(index).Caption
Else
ValueText.Text = ValueText.Text + " " + CmdButn(index).Caption + " "
End If
End Sub

Private Sub CmdButnEq_Click()
ValueText_KeyPress (13)
End Sub

Private Sub Command1_Click()
On Error GoTo 1
If Right(ValueText.Text, 1) = " " Then
ValueText.Text = Left(ValueText.Text, Len(ValueText.Text) - 2)
Else
ValueText.Text = Left(ValueText.Text, Len(ValueText.Text) - 1)
End If
Do While Not Right(ValueText.Text, 1) = " "
DoEvents
ValueText.Text = Left(ValueText.Text, Len(ValueText.Text) - 1)
Loop
1:
End Sub

Private Sub Command2_Click()
ValueText.Text = ""
End Sub

Private Sub Form_Load()
Image1.Picture = Me.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub ValueText_Change()
On Error Resume Next

If InStr(1, ValueText.Text, "| -") > 0 Then
ValueText.Text = ChangeA2B(ValueText.Text, "| -", " |-", 1)
ValueText.SelStart = Len(ValueText.Text)
End If

If InStr(1, ValueText.Text, " .") > 0 Then
ValueText.Text = ChangeA2B(ValueText.Text, " .", ".", 1)
ValueText.SelStart = Len(ValueText.Text)
End If
If InStr(1, ValueText.Text, ". ") > 0 Then
ValueText.Text = ChangeA2B(ValueText.Text, ". ", ".", 1)
ValueText.SelStart = Len(ValueText.Text)
End If

If InStr(1, ValueText.Text, " E ") > 0 Then
ValueText.Text = ChangeA2B(ValueText.Text, " E ", "E", 1)
ValueText.SelStart = Len(ValueText.Text)
End If
If InStr(1, ValueText.Text, " e ") > 0 Then
ValueText.Text = ChangeA2B(ValueText.Text, " e ", "e", 1)
ValueText.SelStart = Len(ValueText.Text)
End If

Prev.Caption = "~= " + Get_Vaule(ValueText.Text)
If Prev.Caption = "~= " Then Prev.Caption = "~= 0"
End Sub

Private Sub ValueText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Not Trim(ValueText.Text) = "" Then
 ValueText.Text = Get_Vaule(ValueText.Text)
 ValueText.SelStart = Len(ValueText.Text)
End If
End Sub
