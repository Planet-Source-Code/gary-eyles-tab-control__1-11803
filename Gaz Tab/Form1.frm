VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Custom tabs"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Delete All"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Own text"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete tab"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add tab"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin CustomTabs.gTab gTab1 
      Height          =   3855
      Left            =   240
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6800
      Begin VB.Frame Frame1 
         Caption         =   "Information"
         Height          =   3255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton Command4 
            Caption         =   "Bottom"
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   9
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Right"
            Height          =   375
            Index           =   3
            Left            =   2160
            TabIndex        =   11
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Style"
            Height          =   375
            Left            =   1200
            TabIndex        =   14
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Highlight selected tab"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Rotatet text (left / right views only)"
            Height          =   495
            Left            =   360
            TabIndex        =   12
            Top             =   2280
            Width           =   2055
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Left"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Top"
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Tab index ="
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1200
            TabIndex        =   6
            Top             =   600
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1080
            TabIndex        =   5
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Tab text ="
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CustomTabs = "Network|New|Add/Remove|Boot|Repair|Paranoia|Mouse|General|Explorer|Desktop|My Computer"
Private Const CustomTabsTips = "Change Network settings|New settings|Add/Remove programs|Change different Boot settings|Repair various settings|Some useful options|Change Mouse settings|Various settings|Change explorer settings|Different desktop settings|My Computer settings"

Private Sub Check1_Click()
gTab1.tRotateText CBool(Check1.Value)
gTab1.RefreshTabs
End Sub

Private Sub Check2_Click()
gTab1.ButtonHighlight CBool(Check2.Value)
End Sub

Private Sub Command1_Click()
gTab1.DeleteTab gTab1.TabIndex
DoEvents
gTab1.RefreshTabs
Frame1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command2_Click()
gTab1.InsertTab "Tab " & Rnd
DoEvents
gTab1.RefreshTabs
Frame1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command3_Click()
Dim cTmp As Long
cTmp = gTab1.TabIndex
gTab1.InsertTab InputBox("Enter text", "Enter text"), InputBox("Enter text", "Enter text"), gTab1.TabIndex + 1

DoEvents
gTab1.RefreshTabs
Frame1.Visible = False
Frame1.Visible = True

gTab1.TabIndex cTmp
End Sub

Private Sub Command4_Click(Index As Integer)
If Index = 0 Then
    gTab1.ChangeStyle tTop, gTab1.GetStyleButton
ElseIf Index = 1 Then
    gTab1.ChangeStyle tBottom, gTab1.GetStyleButton
ElseIf Index = 2 Then
    gTab1.ChangeStyle tLeft, gTab1.GetStyleButton
ElseIf Index = 3 Then
    gTab1.ChangeStyle tRight, gTab1.GetStyleButton
End If

DoEvents
gTab1.RefreshTabs
Frame1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command5_Click()
If gTab1.GetStyleButton Then
    gTab1.ChangeStyle gTab1.GetStyle, False
Else
    gTab1.ChangeStyle gTab1.GetStyle, True
End If
End Sub

Private Sub Command6_Click()
gTab1.DeleteAllTabs
End Sub

Private Sub Form_Load()
'Dim c As Integer
'For c = 1 To 10
'    gTab1.InsertTab "Tab " & c
'Next

'Setup up tabs like Tweak UI
Dim TmpStrings() As String
TmpStrings() = Split(CustomTabs, "|")
Dim TmpStringsTips() As String
TmpStringsTips() = Split(CustomTabsTips, "|")

Dim c As Integer
For c = UBound(TmpStrings) To LBound(TmpStrings) Step -1
        gTab1.InsertTab TmpStrings(c), TmpStringsTips(c)
Next

'Have Mouse selected to start with
gTab1.TabIndex 6
End Sub

Private Sub Form_Resize()
On Error Resume Next
Command1.Top = ScaleHeight - Command1.Height - 5
Command2.Top = ScaleHeight - Command2.Height - 5
Command3.Top = ScaleHeight - Command3.Height - 5
Command6.Top = ScaleHeight - Command6.Height - 5
gTab1.Move 0, 0, ScaleWidth, ScaleHeight - Command1.Height - 10
End Sub

Private Sub gTab1_Resize()
On Error Resume Next
Frame1.Left = gTab1.pLeft * 15 + 10 * 15
Frame1.Top = gTab1.pTop * 15 + (10 * 15)
Frame1.Width = gTab1.pRight * 15 - (20 * 15) - gTab1.pLeft * 15
Frame1.Height = (gTab1.pBottom * 15) - Frame1.Top - (10 * 15)
End Sub

Private Sub gTab1_TabChange(TabIndex As Long, TabText As String)
Label2.Caption = TabText
Label3.Caption = TabIndex
End Sub

