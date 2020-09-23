VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Maths Challenge 2002"
   ClientHeight    =   4185
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5985
   Icon            =   "Maths Challenge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCheck 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox TxtAnswer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   840
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      Begin VB.Label lblOperation 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Operation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LblLevel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Score Chart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   3960
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
      Begin VB.Label LblIncorrect 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1080
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblCorrect 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label LblWinLose 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image ImgSign 
      Height          =   495
      Left            =   240
      Top             =   2040
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   720
      X2              =   2280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label LblSecondNum 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label LblFirstNum 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClearitm 
         Caption         =   "&Clear Score"
      End
      Begin VB.Menu mnuExititm 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOperation 
      Caption         =   "&Operation"
      Begin VB.Menu mnuPlusItm 
         Caption         =   "&Addition"
      End
      Begin VB.Menu mnuMinusItm 
         Caption         =   "&Subtraction"
      End
      Begin VB.Menu mnuDivideItm 
         Caption         =   "&Division"
      End
      Begin VB.Menu mnuMultiplicationItm 
         Caption         =   "Multiplication"
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "&Level"
      Begin VB.Menu mnu1Itm 
         Caption         =   "&1"
      End
      Begin VB.Menu mnu2Itm 
         Caption         =   "&2"
      End
      Begin VB.Menu mnu3Itm 
         Caption         =   "&3"
      End
      Begin VB.Menu mnu4Itm 
         Caption         =   "&4"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------'
'This was done for the complete Newbie to VB. It demonstrates                                     '
'the use of functions and the Module. If you liked this example please vote                       '
'for me on Planet Source.                                                                         '
'If you find a bug in this or would just like to say something email me at Cool9546@hotmail.com   '
'-------------------------------------------------------------------------------------------------'
Option Explicit
Private Sub CmdCheck_Click()
    'Load the Functions
    TxtAnswer.SetFocus
    Calculate
    Random
    CheckAnswer
    TxtAnswer.SetFocus 'Set Focus on the text box
End Sub
Private Sub CheckAnswer()
    Times = Times + 1
    
    If Val(Result) = Val(TxtAnswer.Text) Then 'If answer is correct
        Wins = Wins + 1
        LblWinLose.ForeColor = &HC000& 'Change text colour to Green
        LblWinLose.Caption = "Correct"
        LblCorrect.Caption = Wins
        Result = sndPlaySound(App.Path & "/Applause.wav", 1)
    Else 'If Answer is incorrect
        Loses = Loses + 1
        LblWinLose.ForeColor = &HFF& 'Change text colour to Red
        LblWinLose.Caption = "Incorrect"
        LblIncorrect.Caption = Loses
    End If
    Label1.Caption = "Percent Correct: " & PercentWins(Wins, Times)
    TxtAnswer.Text = ""
End Sub
Private Sub Form_Load()
On Error GoTo ImageErr
    Label1.Caption = "Percent Correct:"
    ImgSign.Picture = LoadPicture(App.Path & "\misc18.ico")
    Randomize 'Initialize random number generator
    LblFirstNum.Caption = Int((Rnd * 15) + 1)
    LblSecondNum.Caption = Int((Rnd * 15) + 1)
    LblLevel.Caption = "Level: " & "1"
    lblOperation.Caption = "Addition"
ImageErr: 'If Picture files are not in the directory
    If Err.Number = 53 Then
        MsgBox ("Picture files not found"), , "Error"
        End
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MsgBox ("If you thought that was good please vote for me."), , "Vote"
End Sub

Private Sub mnu1Itm_Click()
    LblLevel.Caption = "Level: " & "1"
End Sub

Private Sub mnu2Itm_Click()
     LblLevel.Caption = "Level: " & "2"
End Sub

Private Sub mnu3Itm_Click()
    LblLevel.Caption = "Level: " & "3"
End Sub

Private Sub mnu4Itm_Click()
     LblLevel.Caption = "Level: " & "4"
End Sub

Private Sub mnuClearitm_Click()
    LblCorrect.Caption = "" 'Delete everything in the score box
    LblIncorrect.Caption = "" 'and percent box
    Label1.Caption = ""
    Wins = 0
    Loses = 0
    Times = 0
End Sub

Private Sub mnuDivideItm_Click()
On Error GoTo PicError
     lblOperation.Caption = "Division"
     ImgSign.Picture = LoadPicture(App.Path & "\misc21.ico")
     Random
PicError:
     If Err.Number = 53 Then
        MsgBox ("Picture files are missing put then into the directory"), , _
        "Error"
        End
     End If
End Sub

Private Sub mnuExititm_Click()
    Unload Me 'Quite
End Sub

Private Sub mnuMinusItm_Click()
On Error GoTo PicError
    lblOperation.Caption = "Subtraction"
    ImgSign.Picture = LoadPicture(App.Path & "\misc19.ico")
    Random
PicError:
     If Err.Number = 53 Then
        MsgBox ("Picture files are missing put then into the directory"), , _
        "Error"
        End
    End If
End Sub

Private Sub mnuMultiplicationItm_Click()
On Error GoTo PicError
     lblOperation.Caption = "Multiplication"
     ImgSign.Picture = LoadPicture(App.Path & "\misc20.ico")
     Random
PicError:
     If Err.Number = 53 Then
        MsgBox ("Picture files are missing put then into the directory"), , _
        "Error"
        End
     End If
End Sub

Private Sub mnuPlusItm_Click()
On Error GoTo PicError
    lblOperation.Caption = "Addition"
    ImgSign.Picture = LoadPicture(App.Path & "\misc18.ico")
    Random
PicError:
     If Err.Number = 53 Then
        MsgBox ("Picture files are missing put then into the directory"), , _
        "Error"
        End
     End If
End Sub
Private Sub Calculate() 'Calculates the correct answer to be compared to later
    Dim First   As Integer
    Dim Second  As Integer
    
    First = Val(LblFirstNum.Caption)
    Second = Val(LblSecondNum.Caption)
    
    If (lblOperation.Caption = "Addition") Then
       Result = First + Second
    End If
    If (lblOperation.Caption = "Subtraction") Then
       Result = First - Second
    End If
    If (lblOperation.Caption = "Multiplication") Then
       Result = First * Second
    End If
    If (lblOperation.Caption = "Division") And LblLevel.Caption _
        = "Level: " & "1" Then
       Result = Format(First / Second, "0")
    ElseIf (lblOperation.Caption = "Division") Then
       Result = Format(First / Second, "0.0")
    End If
End Sub
Private Sub Random() 'Function which picks the random numbers
    'This is the heart of the program which pics the random numbers
    If LblLevel.Caption = "Level: " & "1" And _
        (lblOperation.Caption = "Subtraction") Then
        LblFirstNum.Caption = Int((Rnd * 20) + 15) 'Pick a random number from 1 to 15
        LblSecondNum.Caption = Int((Rnd * 15) + 1)
    ElseIf LblLevel.Caption = "Level: " & "1" Then
        LblFirstNum.Caption = Int((Rnd * 15) + 1) 'Pick a random number from 1 to 15
        LblSecondNum.Caption = Int((Rnd * 15) + 1)
    End If
    If LblLevel.Caption = "Level: " & "2" Then
        LblFirstNum.Caption = Int((Rnd * 25) + 1)
        LblSecondNum.Caption = Int((Rnd * 25) + 1)
    End If
    If LblLevel.Caption = "Level: " & "3" Then
        LblFirstNum.Caption = Int((Rnd * 60) + 1)
        LblSecondNum.Caption = Int((Rnd * 60) + 1)
    End If
    If LblLevel.Caption = "Level: " & "4" Then
        LblFirstNum.Caption = Int((Rnd * 100) + 1)
        LblSecondNum.Caption = Int((Rnd * 100) + 1)
    End If
End Sub
Private Sub TxtAnswer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me 'Exit if the Escape key was pressed
    End If
    If KeyAscii = 13 Then
       Call CmdCheck_Click
    End If
End Sub
