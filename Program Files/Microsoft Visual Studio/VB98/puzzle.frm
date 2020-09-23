VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "That old Puzzle game"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "puzzle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&About me"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdShuffle 
      Caption         =   "&Shuffle"
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Turns :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   15
      Left            =   3960
      TabIndex        =   15
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   14
      Left            =   2760
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   13
      Left            =   1560
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   12
      Left            =   360
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   11
      Left            =   3960
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   10
      Left            =   2760
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   1560
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   3960
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   2760
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function swap_caption(index1 As Integer, index2 As Integer) 'swap captions
Dim str As String
str = Label(index1).Caption
Label(index1).Caption = Label(index2).Caption
Label(index2).Caption = str
End Function
Private Sub cmdShuffle_Click()
'When this button is clicked, two number are produced (d1 and d2) randomly. Then using d1
'd2 as index, swap the caption of label(d1) with label(d2), so two random labels swap their values.
Dim d1, d2, val, i As Integer 'val-useless, but without it VB wont compile this program
For i = 0 To 255    'do the swapping process 255 times, you may increase it if you like
    Randomize
    d1 = Int(Rnd(2) * 100) 'generate first random number
    Randomize
    d2 = Int(Rnd(2) * 100) 'generate second random number
    If d1 > 15 Then 'check if random number is within range
    d1 = 15         'if not then cut it down to size
    End If
    If d2 > 15 Then 'check if random number is within range
    d2 = 0
    End If
    val = swap_caption(Int(d1), Int(d2)) 'I am not returning any value, still val is required, whats the logic
Next i
Label2.Caption = "0"    'Make the turns equal to zero
End Sub

Private Sub Command1_Click() 'About me
MsgBox "Programmer:- Niloy Mondal. Email:- niloygk@yahoo.com"
End Sub

Private Sub Label_Click(Index As Integer)
'When you click, label(index) checks which of the four surrounding boxes is empty and then swaps value with it
Dim i, winflag As Integer
'Checks if index does not become negative and swapping doesnt happen across the both ends
If Index > 0 And Index <> 4 And Index <> 8 And Index <> 12 Then
    If Label(Index - 1).Caption = " " Then
        Label(Index - 1).Caption = Label(Index).Caption
        Label(Index).Caption = " "
        Label2.Caption = Label2.Caption + 1
    End If
End If
If Index > 3 Then
    If Label(Index - 4).Caption = " " Then
        Label(Index - 4).Caption = Label(Index).Caption
        Label(Index).Caption = " "
        Label2.Caption = Label2.Caption + 1
    End If
End If
'Checks if index does not go out of range and swapping doesnt happen across the both ends
If Index < 15 And Index <> 3 And Index <> 7 And Index <> 11 Then
    If Label(Index + 1).Caption = " " Then
        Label(Index + 1).Caption = Label(Index).Caption
        Label(Index).Caption = " "
        Label2.Caption = Label2.Caption + 1
    End If
End If
If Index < 12 Then
    If Label(Index + 4).Caption = " " Then
        Label(Index + 4).Caption = Label(Index).Caption
        Label(Index).Caption = " "
        Label2.Caption = Label2.Caption + 1
    End If
End If
For i = 0 To 14     'check if user has solved the puzzle
    If Label(i).Caption = (i + 1) Then
        winflag = 1
    Else
        winflag = 0
        Exit For
    End If
Next i
If winflag = 1 Then 'show the welldone message if user has solved the puzzle
    MsgBox "Welldone, You have solved the puzzle"
End If
End Sub
