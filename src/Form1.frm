VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2048"
   ClientHeight    =   8670
   ClientLeft      =   4290
   ClientTop       =   1575
   ClientWidth     =   12675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12675
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   16
      Left            =   13440
      Picture         =   "Form1.frx":0ECA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   15
      Left            =   12960
      Picture         =   "Form1.frx":5C95
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   42
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   14
      Left            =   12480
      Picture         =   "Form1.frx":A6B8
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   13
      Left            =   12000
      Picture         =   "Form1.frx":F19D
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   12
      Left            =   13440
      Picture         =   "Form1.frx":140A8
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   11
      Left            =   12960
      Picture         =   "Form1.frx":18DAD
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   10
      Left            =   12480
      Picture         =   "Form1.frx":1DF01
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   37
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   9
      Left            =   12000
      Picture         =   "Form1.frx":228E1
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   36
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   8
      Left            =   13440
      Picture         =   "Form1.frx":27770
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   35
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   7
      Left            =   12960
      Picture         =   "Form1.frx":2CFFB
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   34
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   6
      Left            =   12480
      Picture         =   "Form1.frx":3235E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   33
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   5
      Left            =   12000
      Picture         =   "Form1.frx":37875
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   32
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   4
      Left            =   13440
      Picture         =   "Form1.frx":3D125
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   3
      Left            =   12960
      Picture         =   "Form1.frx":4207D
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   2
      Left            =   12480
      Picture         =   "Form1.frx":465DA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   1
      Left            =   12000
      Picture         =   "Form1.frx":4A2FD
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   28
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picbank 
      Height          =   495
      Index           =   0
      Left            =   12000
      Picture         =   "Form1.frx":4E0C0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8625
      Left            =   0
      Picture         =   "Form1.frx":513AB
      ScaleHeight     =   8625
      ScaleMode       =   0  'User
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   0
      Width           =   11700
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   15
         Left            =   6600
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   27
         Top             =   6480
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   14
         Left            =   4560
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   26
         Top             =   6480
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   13
         Left            =   2520
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   25
         Top             =   6480
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   12
         Left            =   480
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   24
         Top             =   6480
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   11
         Left            =   6600
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   23
         Top             =   4440
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   10
         Left            =   4560
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   22
         Top             =   4440
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   9
         Left            =   2520
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   21
         Top             =   4440
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   8
         Left            =   480
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   20
         Top             =   4440
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   7
         Left            =   6600
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   19
         Top             =   2400
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   6
         Left            =   4560
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   18
         Top             =   2400
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   5
         Left            =   2520
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   17
         Top             =   2400
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   4
         Left            =   480
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   3
         Left            =   6600
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   2
         Left            =   4560
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   1
         Left            =   2520
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox block 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   480
         ScaleHeight     =   1785
         ScaleWidth      =   1785
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton right 
         Caption         =   "¡ú"
         Height          =   615
         Left            =   10560
         TabIndex        =   10
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton left 
         Caption         =   "¡û"
         Height          =   615
         Left            =   9120
         TabIndex        =   9
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton down 
         Caption         =   "¡ý"
         Height          =   615
         Left            =   9840
         TabIndex        =   8
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton up 
         Caption         =   "¡ü"
         Height          =   615
         Left            =   9840
         TabIndex        =   7
         Top             =   6600
         Width           =   615
      End
      Begin VB.CommandButton help 
         Caption         =   "Help"
         Height          =   375
         Left            =   9600
         TabIndex        =   6
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton undo 
         Caption         =   "Undo(U)"
         Height          =   375
         Left            =   10560
         TabIndex        =   5
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton start 
         Caption         =   "Start"
         Height          =   375
         Left            =   8760
         TabIndex        =   4
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by Gaoxiang Liu"
         Height          =   255
         Left            =   9120
         TabIndex        =   50
         Top             =   8160
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit:  ESC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   49
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Undo: U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   48
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Right:  D,L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   47
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Left:  A,J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   46
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Down:  S,K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   45
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Up:  I,W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   44
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label_score 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   9000
         TabIndex        =   3
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Best:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   9000
         TabIndex        =   2
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2048"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   9000
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim map(3, 3) As Integer        'current map
'the element in map can be 0 to 16, they represent block's number
'0:null,1:2^1,2:2^2...
Dim pre_map(3, 3) As Integer    'previous map (used for undo)

Dim IfFirst As Boolean          'true:the first to play
Dim GameStatus As Integer       '0:normal, 1:win(still can be played), -1: lose


Private Sub block_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub down_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub Form_Load()
    Unload Form2
    Width = Picture1.ScaleWidth + 120
    Height = Picture1.ScaleHeight + 500
    Picture1.Top = 0
    Picture1.left = 0
    Call Init
End Sub

Public Sub Check()
    Dim i, j, t As Integer
    Dim b1 As Boolean                       'decide whether all the blocks are full
    Dim b2 As Boolean                       'decide whethr user have lost when all the blocks are full
    b1 = False
    b2 = False
    Call ShowBestScore
    If GameStatus = 0 Then
        For i = 0 To 3                      'check win
            For j = 0 To 3
                If map(i, j) = 11 Then      '2048 block exists, win
                    GameStatus = 1
                    MsgBox "You win!", , "Congratulations!"
                End If
            Next j
        Next i
        For i = 0 To 3                      'check lose
            For j = 0 To 3
                If map(i, j) = 0 Then b1 = True
            Next j
        Next i
        If b1 = False Then                  'all the blocks are full
            For i = 0 To 3                  'find whether it can be moved
                t = 0
                For j = 0 To 3
                    If map(i, j) = t Then b2 = True Else t = map(i, j)
                Next j
            Next i
            For j = 0 To 3
                t = 0
                For i = 0 To 3
                    If map(i, j) = t Then b2 = True Else t = map(i, j)
                Next i
            Next j
            If b2 = False Then
                If GameStatus <> 1 Then     'haven't won
                    MsgBox "You lose!", , "Sorry"
                Else
                    MsgBox "Good Luck next time!", , "Sorry"
                End If
                MsgBox "Press start to replay.", , "Info"
                GameStatus = -1
                undo.Enabled = False
            End If
        End If
    End If
End Sub

Public Sub ShowBestScore()
    Dim i, j, max As Integer
    max = 0
    For i = 0 To 3
        For j = 0 To 3
            If map(i, j) > max Then max = map(i, j)
        Next j
    Next i
    Label_score.Caption = 2 ^ max
End Sub

Public Sub Init()
    Dim i, j As Integer
    IfFirst = True
    GameStatus = 0
    undo.Enabled = False
    For i = 0 To 3
        For j = 0 To 3
            map(i, j) = 0
        Next j
    Next i
    Randomize
    i = Int(4 * Rnd)                 'int((b-a+1)*rnd+a)
    j = Int(4 * Rnd)
    map(i, j) = Int(2 * Rnd + 1)    'set the block to 2 or 4
    Do
        i = Int(4 * Rnd)
        j = Int(4 * Rnd)
    Loop Until map(i, j) = 0
    map(i, j) = Int(2 * Rnd + 1)
    Call PrintMap
    Call ShowBestScore
End Sub

Public Sub PrintMap()
    Dim n As Integer
    For n = 0 To 15
        block(n).Picture = picbank(map(n \ 4, n Mod 4)).Picture
    Next n
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Form2
End Sub

Private Sub help_Click()
    Load Form2
    Form2.Show
End Sub

Private Sub help_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub left_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub right_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub start_Click()
    Call Init
End Sub

Private Sub start_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub undo_Click()
    Dim i, j As Integer
    For i = 0 To 3
        For j = 0 To 3
            map(i, j) = pre_map(i, j)
        Next j
    Next i
    Call PrintMap
    undo.Enabled = False
End Sub

Private Sub undo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub

Private Sub up_Click()
    IfFirst = False
    Dim i, j, k, t As Integer
    Dim canbemoved As Boolean               'whether the blocks can be moved
    canbemoved = False
    Dim temp_map(3, 3) As Integer           'temporatory map
    Dim IfMerge(3) As Integer               'to mark whether the block has been merged
                                            'if so, this block won't be merged again
    If GameStatus <> -1 Then
        For i = 0 To 3                      'record the previous operation
            For j = 0 To 3
                temp_map(i, j) = map(i, j)
            Next j
        Next i
        For j = 0 To 3
            For i = 0 To 3
                IfMerge(i) = 0              'reset
            Next i
            For i = 1 To 3
                If map(i, j) <> 0 Then
                    If map(i - 1, j) = 0 Then
                        t = i
                        k = i
                        Do
                            k = k - 1
                            map(k, j) = map(t, j)
                            map(t, j) = 0
                            t = t - 1
                            canbemoved = True
                            If k = 0 Then
                                Exit Do
                            Else
                                If map(k - 1, j) <> 0 Then
                                    If map(k - 1, j) = map(k, j) And IfMerge(k - 1) = 0 Then
                                        map(k - 1, j) = map(k - 1, j) + 1
                                        map(k, j) = 0
                                        IfMerge(k - 1) = 1
                                    End If
                                    Exit Do
                                End If
                            End If
                        Loop
                    Else
                        If map(i - 1, j) = map(i, j) And IfMerge(i - 1) = 0 Then
                            map(i - 1, j) = map(i - 1, j) + 1
                            map(i, j) = 0
                            IfMerge(i - 1) = 1
                            canbemoved = True
                        End If
                    End If
                End If
            Next i
        Next j
        If canbemoved = True Then
            Do
                i = Int(4 * Rnd)
                j = Int(4 * Rnd)
            Loop Until map(i, j) = 0
            undo.Enabled = True
            map(i, j) = Int(2 * Rnd + 1)
            For i = 0 To 3
                For j = 0 To 3
                    pre_map(i, j) = temp_map(i, j)  'if the blocks were moved, pre_map will take the previous map
                                                    'otherwise pre_map will be remained
                Next j
            Next i
        End If
    End If
    Call PrintMap
    Call Check
End Sub

Private Sub down_Click()
    IfFirst = False
    Dim i, j, k, t As Integer
    Dim canbemoved As Boolean               'whether the blocks can be moved
    canbemoved = False
    Dim temp_map(3, 3) As Integer           'temporatory map
    Dim IfMerge(3) As Integer               'to mark whether the block has been merged
                                            'if so, this block won't be merged again
    If GameStatus <> -1 Then
        For i = 0 To 3                      'record the previous operation
            For j = 0 To 3
                temp_map(i, j) = map(i, j)
            Next j
        Next i
        For j = 0 To 3
            For i = 0 To 3
                IfMerge(i) = 0              'reset
            Next i
            For i = 2 To 0 Step -1
                If map(i, j) <> 0 Then
                    If map(i + 1, j) = 0 Then
                        t = i
                        k = i
                        Do
                            k = k + 1
                            map(k, j) = map(t, j)
                            map(t, j) = 0
                            t = t + 1
                            canbemoved = True
                            If k = 3 Then
                                Exit Do
                            Else
                                If map(k + 1, j) <> 0 Then
                                    If map(k + 1, j) = map(k, j) And IfMerge(k + 1) = 0 Then
                                        map(k + 1, j) = map(k + 1, j) + 1
                                        map(k, j) = 0
                                        IfMerge(k + 1) = 1
                                    End If
                                    Exit Do
                                End If
                            End If
                        Loop
                    Else
                        If map(i + 1, j) = map(i, j) And IfMerge(i + 1) = 0 Then
                            map(i + 1, j) = map(i + 1, j) + 1
                            map(i, j) = 0
                            IfMerge(i + 1) = 1
                            canbemoved = True
                        End If
                    End If
                End If
            Next i
        Next j
        If canbemoved = True Then
            Do
                i = Int(4 * Rnd)
                j = Int(4 * Rnd)
            Loop Until map(i, j) = 0
            undo.Enabled = True
            map(i, j) = Int(2 * Rnd + 1)
            For i = 0 To 3
                For j = 0 To 3
                    pre_map(i, j) = temp_map(i, j)  'if the blocks were moved, pre_map will take the previous map
                                                    'otherwise pre_map will be remained
                Next j
            Next i
        End If
    End If
    Call PrintMap
    Call Check
End Sub

Private Sub left_Click()
    IfFirst = False
    Dim i, j, k, t As Integer
    Dim canbemoved As Boolean                'whether the blocks can be moved
    canbemoved = False
    Dim temp_map(3, 3) As Integer           'temporatory map
    Dim IfMerge(3) As Integer               'to mark whether the block has been merged
                                            'if so, this block won't be merged again
    If GameStatus <> -1 Then
        For i = 0 To 3                      'record the previous operation
            For j = 0 To 3
                temp_map(i, j) = map(i, j)
            Next j
        Next i
        For i = 0 To 3
            For j = 0 To 3
                IfMerge(j) = 0              'reset
            Next j
            For j = 1 To 3
                If map(i, j) <> 0 Then
                    If map(i, j - 1) = 0 Then
                        t = j
                        k = j
                        Do
                            k = k - 1
                            map(i, k) = map(i, t)
                            map(i, t) = 0
                            t = t - 1
                            canbemoved = True
                            If k = 0 Then
                                Exit Do
                            Else
                                If map(i, k - 1) <> 0 Then
                                    If map(i, k - 1) = map(i, k) And IfMerge(k - 1) = 0 Then
                                        map(i, k - 1) = map(i, k - 1) + 1
                                        map(i, k) = 0
                                        IfMerge(k - 1) = 1
                                    End If
                                    Exit Do
                                End If
                            End If
                        Loop
                    Else
                        If map(i, j - 1) = map(i, j) And IfMerge(j - 1) = 0 Then
                            map(i, j - 1) = map(i, j - 1) + 1
                            map(i, j) = 0
                            IfMerge(j - 1) = 1
                            canbemoved = True
                        End If
                    End If
                End If
            Next j
        Next i
        If canbemoved = True Then
            Do
                i = Int(4 * Rnd)
                j = Int(4 * Rnd)
            Loop Until map(i, j) = 0
            undo.Enabled = True
            map(i, j) = Int(2 * Rnd + 1)
            For i = 0 To 3
                For j = 0 To 3
                    pre_map(i, j) = temp_map(i, j)  'if the blocks were moved, pre_map will take the previous map
                                                    'otherwise pre_map will be remained
                Next j
            Next i
        End If
    End If
    Call PrintMap
    Call Check
End Sub

Private Sub right_Click()
IfFirst = False
    Dim i, j, k, t As Integer
    Dim canbemoved As Boolean               'whether the blocks can be moved
    canbemoved = False
    Dim temp_map(3, 3) As Integer           'temporatory map
    Dim IfMerge(3) As Integer               'to mark whether the block has been merged
                                            'if so, this block won't be merged again
    If GameStatus <> -1 Then
        For i = 0 To 3                      'record the previous operation
            For j = 0 To 3
                temp_map(i, j) = map(i, j)
            Next j
        Next i
        For i = 0 To 3
            For j = 0 To 3
                IfMerge(j) = 0              'reset
            Next j
            For j = 2 To 0 Step -1
                If map(i, j) <> 0 Then
                    If map(i, j + 1) = 0 Then
                        t = j
                        k = j
                        Do
                            k = k + 1
                            map(i, k) = map(i, t)
                            map(i, t) = 0
                            t = t + 1
                            canbemoved = True
                            If k = 3 Then
                                Exit Do
                            Else
                                If map(i, k + 1) <> 0 Then
                                    If map(i, k + 1) = map(i, k) And IfMerge(k + 1) = 0 Then
                                        map(i, k + 1) = map(i, k + 1) + 1
                                        map(i, k) = 0
                                        IfMerge(k + 1) = 1
                                    End If
                                    Exit Do
                                End If
                            End If
                        Loop
                    Else
                        If map(i, j + 1) = map(i, j) And IfMerge(j + 1) = 0 Then
                            map(i, j + 1) = map(i, j + 1) + 1
                            map(i, j) = 0
                            IfMerge(j + 1) = 1
                            canbemoved = True
                        End If
                    End If
                End If
            Next j
        Next i
        If canbemoved = True Then
            Do
                i = Int(4 * Rnd)
                j = Int(4 * Rnd)
            Loop Until map(i, j) = 0
            undo.Enabled = True
            map(i, j) = Int(2 * Rnd + 1)
            For i = 0 To 3
                For j = 0 To 3
                    pre_map(i, j) = temp_map(i, j)  'if the blocks were moved, pre_map will take the previous map
                                                    'otherwise pre_map will be remained
                Next j
            Next i
        End If
    End If
    Call PrintMap
    Call Check
End Sub

Private Sub up_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Or KeyAscii = 65 Or KeyAscii = 74 Or KeyAscii = 106 Then   'left key(aj)
        Call left_Click
    End If
    If KeyAscii = 87 Or KeyAscii = 119 Or KeyAscii = 73 Or KeyAscii = 105 Then  'up key(wi)
        Call up_Click
    End If
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 76 Or KeyAscii = 108 Then  'right key(dl)
        Call right_Click
    End If
    If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 75 Or KeyAscii = 107 Then  'down key(sk)
        Call down_Click
    End If
    If KeyAscii = 27 Then                                                       'ESC
        Unload Me
    End If
    If KeyAscii = 85 Or KeyAscii = 117 Then                                     'undo(uU)
        If undo.Enabled = True Then Call undo_Click
    End If
End Sub
