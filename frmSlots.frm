VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scoring"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1755
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  $500"
         Height          =   195
         Left            =   300
         TabIndex        =   8
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  $1000"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2 Slots Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All 3 Slots Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4620
      Top             =   3120
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Click a window to stop it from spinning."
      Height          =   435
      Left            =   2340
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Press ""GO"" to spin."
      Height          =   255
      Left            =   2340
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Spin:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Any of the above are doubled if they are Contest Winners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   3
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Image slot1 
      Height          =   1110
      Index           =   2
      Left            =   5260
      Picture         =   "frmSlots.frx":0000
      Top             =   990
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image slot1 
      Height          =   1110
      Index           =   1
      Left            =   3960
      Picture         =   "frmSlots.frx":0B50
      Top             =   990
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image slot1 
      Height          =   1110
      Index           =   0
      Left            =   2670
      Picture         =   "frmSlots.frx":19AF
      Top             =   990
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   6600
      Picture         =   "frmSlots.frx":297E
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00.00"
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Money:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   4020
      Left            =   2160
      Picture         =   "frmSlots.frx":330A
      Top             =   -60
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyMoney As Single

Dim SlotTags(0 To 2) As Slots
Dim Scroll(0 To 2) As Boolean
Dim Last(0 To 20) As Slots
'MyValue = Int((6 * Rnd) + 1)
Public Enum Slots
    BX
    CW
    EG
    WZ
    VB
End Enum

Private Function GetTag(TagNum As Slots) As String
    Select Case TagNum
        Case 0
            GetTag = "BX"
        Case 1
            GetTag = "CW"
        Case 2
            GetTag = "EG"
        Case 3
            GetTag = "WZ"
        Case 4
            GetTag = "VB"
    End Select
End Function

Private Sub Form_Load()
    MyMoney = 2000
    
    Image1.Picture = LoadPicture(App.Path & "\Machine.gif")
    
    Scroll(0) = False
    Scroll(1) = False
    Scroll(2) = False
    
    slot1(0).Visible = False
    slot1(1).Visible = False
    slot1(2).Visible = False
    
    GenerateStartingSpots
    Me.Visible = True
End Sub

Private Sub GenerateStartingSpots()
    For x = 0 To 2
        Randomize
        SlotTags(x) = Int((5 * Rnd))
        slot1(x).Picture = LoadPicture(App.Path & "\" & GetTag(SlotTags(x)) & "1.gif")
        slot1(x).Visible = True
    Next x
End Sub

Private Sub SpinWheels()
    
    If MyMoney < 100 Then GoTo CantSpin
    
    MyMoney = MyMoney - 100
    Scroll(0) = True
    Scroll(1) = True
    Scroll(2) = True
    
    Last(0) = 5
    Last(1) = 5
    Last(2) = 5
    
    slot1(0).Tag = ""
    slot1(1).Tag = ""
    slot1(2).Tag = ""
    
    Do Until Not (Scroll(0)) And Not (Scroll(1)) And Not (Scroll(2))
        For x = 0 To 2
            If Scroll(x) = True Then
10:             Randomize
                SlotTags(x) = Int((5 * Rnd))
                If SlotTags(x) = Last(x) Then GoTo 10
                Last(x) = SlotTags(x)
                slot1(x).Picture = LoadPicture(App.Path & "\" & GetTag(SlotTags(x)) & ".gif")
            Else
                If slot1(x).Tag <> "S" Then
                    slot1(x).Picture = LoadPicture(App.Path & "\" & GetTag(SlotTags(x)) & "1.gif")
                    slot1(x).Tag = "S"
                End If
            End If
            slot1(x).Visible = True
            DoEvents
        Next x
    Loop
    
    For x = 0 To 2
        slot1(x).Picture = LoadPicture(App.Path & "\" & GetTag(SlotTags(x)) & "1.gif")
    Next x
    CheckForAWin
    Exit Sub
    
CantSpin:
    MsgBox "You don't have enough money to spin!"
End Sub

Private Sub Image2_Click()
    If Scroll(0) Or Scroll(1) Or Scroll(2) Then Exit Sub
    SpinWheels
End Sub

Private Sub slot1_Click(Index As Integer)
    Scroll(Index) = False
End Sub

Private Sub Timer1_Timer()
    lblMoney = Format(MyMoney, "0.00")
End Sub


Private Sub CheckForAWin()
    Label8 = "Last Spin:  -100"
    Label8.ForeColor = vbRed
    
    If (SlotTags(0) = SlotTags(1)) And (SlotTags(2) = SlotTags(1)) Then
        Label8.ForeColor = vbBlack
        Select Case SlotTags(0)
            Case 1
                MsgBox "You are a Contest Winner! You Earn $3000!!"
                MyMoney = MyMoney + 3000
                Label8 = "Last Spin:  +3000"
            Case Else
                MsgBox "JackPot! Here's $1500."
                MyMoney = MyMoney + 1500
                Label8 = "Last Spin:  +1500"
        End Select
        
    ElseIf (SlotTags(0) = SlotTags(1)) Then
        Label8.ForeColor = vbBlack
        Select Case SlotTags(0)
            Case 1
                MyMoney = MyMoney + 1500
                Label8 = "Last Spin:  +1500"
            Case Else
                MyMoney = MyMoney + 750
                Label8 = "Last Spin:  +750"
        End Select
    ElseIf (SlotTags(2) = SlotTags(1)) Then
        Label8.ForeColor = vbBlack
        Select Case SlotTags(1)
            Case 1
                MyMoney = MyMoney + 1500
                Label8 = "Last Spin:  +1500"
            Case Else
                MyMoney = MyMoney + 750
                Label8 = "Last Spin:  +750"
        End Select
    ElseIf (SlotTags(0) = SlotTags(2)) Then
        Label8.ForeColor = vbBlack
        Select Case SlotTags(0)
            Case 1
                MyMoney = MyMoney + 1500
                Label8 = "Last Spin:  +1500"
            Case Else
                MyMoney = MyMoney + 750
                Label8 = "Last Spin:  +750"
        End Select
    End If
    
    
End Sub
