VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Additional material on the book: Markov Chains from theory to implementation and experimentation"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1046
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame startsta 
      Caption         =   "Number of steps:"
      Height          =   1335
      Index           =   2
      Left            =   7800
      TabIndex        =   46
      Top             =   4200
      Width           =   2895
      Begin VB.TextBox KST 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   47
         Text            =   "50"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Number of steps (k) ="
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   48
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame startsta 
      Caption         =   "Plot line for:"
      Height          =   1335
      Index           =   1
      Left            =   10920
      TabIndex        =   37
      Top             =   4200
      Width           =   4335
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   45
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   44
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   43
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PlotL 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   40
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   39
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   38
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame startsta 
      Caption         =   "Initial state vector:"
      Height          =   1935
      Index           =   0
      Left            =   4320
      TabIndex        =   28
      Top             =   4200
      Width           =   3255
      Begin VB.OptionButton StartFrom 
         Caption         =   "Option1"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   53
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton StartFrom 
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   52
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton StartFrom 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   51
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton StartFrom 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   50
         Top             =   1080
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox VComponent 
         Height          =   375
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox VComponent 
         Height          =   375
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox VComponent 
         Height          =   375
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox VComponent 
         Height          =   375
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Start this system from state:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   31
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   30
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transition matrix:"
      Height          =   2775
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   3735
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Text            =   "0.33333"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   18
         Text            =   "0.33333"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   17
         Text            =   "0.33333"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   3
         Left            =   2760
         TabIndex        =   16
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   15
         Text            =   "0.5"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   14
         Text            =   "0.5"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   6
         Left            =   2040
         TabIndex        =   13
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   7
         Left            =   2760
         TabIndex        =   12
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   11
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   10
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   10
         Left            =   2040
         TabIndex        =   9
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   8
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   12
         Left            =   600
         TabIndex        =   7
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   13
         Left            =   1320
         TabIndex        =   6
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   14
         Left            =   2040
         TabIndex        =   5
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox MCell 
         Height          =   375
         Index           =   15
         Left            =   2760
         TabIndex        =   4
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   21
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton Go 
      Caption         =   "Do it !"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   6360
      Width           =   3255
   End
   Begin VB.PictureBox graf_val 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   480
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   983
      TabIndex        =   1
      Top             =   360
      Width           =   14775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5880
      Width           =   7455
   End
   Begin VB.Label Label5 
      Caption         =   "Values calculated for each step:"
      Height          =   255
      Index           =   3
      Left            =   7800
      TabIndex        =   55
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Prediction:"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   49
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:      Markov Chains: From Theory To Implementation And Experimentation                #
'# Author:    Dr. Paul A. Gagniuc                                                             #
'# Data:      01/09/2016                                                                      #
'# Category:  open source software                                                            #
'#                                                                                            #
'# Description:                                                                               #
'# Probabilistic framework for prediction used for the main figures in CHAPTER 6,7,8.         #
'##############################################################################################

Dim M(1 To 4, 1 To 4) As String

Private Sub Form_Load()
Dim v(0 To 3, 0 To 1) As Variant

Form1.DrawWidth = 3

Call draw_scale(Val(KST(0).Text))

M(1, 1) = Val(MCell(0).Text)
M(1, 2) = Val(MCell(1).Text)
M(1, 3) = Val(MCell(2).Text)
M(1, 4) = Val(MCell(3).Text)

M(2, 1) = Val(MCell(4).Text)
M(2, 2) = Val(MCell(5).Text)
M(2, 3) = Val(MCell(6).Text)
M(2, 4) = Val(MCell(7).Text)

M(3, 1) = Val(MCell(8).Text)
M(3, 2) = Val(MCell(9).Text)
M(3, 3) = Val(MCell(10).Text)
M(3, 4) = Val(MCell(11).Text)

M(4, 1) = Val(MCell(12).Text)
M(4, 2) = Val(MCell(13).Text)
M(4, 3) = Val(MCell(14).Text)
M(4, 4) = Val(MCell(15).Text)


chain = Val(KST(0).Text)

v(0, 0) = Val(VComponent(0).Text)
v(1, 0) = Val(VComponent(1).Text)
v(2, 0) = Val(VComponent(2).Text)
v(3, 0) = Val(VComponent(3).Text)

v(0, 1) = 0
v(1, 1) = 0
v(2, 1) = 0
v(3, 1) = 0


    'oldxA = xA
    oldyA = (graf_val.ScaleHeight / 100) * (100 * v(0, 0))

    'oldxT = xT
    oldyT = (graf_val.ScaleHeight / 100) * (100 * v(1, 0))
    
    'oldxC = xC
    oldyC = (graf_val.ScaleHeight / 100) * (100 * v(2, 0))
    
    'oldxG = xG
    oldyG = (graf_val.ScaleHeight / 100) * (100 * v(3, 0))


For k = 1 To chain
    
    For i = 0 To 3
        For j = 0 To 3
            v(i, 1) = v(i, 1) + (v(j, 0) * M(j + 1, i + 1))
        Next j
    Next i

    For i = 0 To 3
        v(i, 0) = v(i, 1)
        v(i, 1) = 0
    Next i
    
    ww = 3
    
    A = Round(v(0, 0), ww)
    T = Round(v(1, 0), ww)
    C = Round(v(2, 0), ww)
    G = Round(v(3, 0), ww)
    
    Text1.Text = Text1.Text & "Step (" & k & ")=[" & A & " | " & T & " | " & C & " | " & G & "] = " & (A + T + C + G) & vbCrLf

    revers = graf_val.ScaleHeight

    xA = (graf_val.ScaleHeight / 100) * (100 * A)
    yA = (graf_val.ScaleHeight / 100) * (100 * A)
    
    xT = (graf_val.ScaleHeight / 100) * (100 * T)
    yT = (graf_val.ScaleHeight / 100) * (100 * T)
    
    xC = (graf_val.ScaleHeight / 100) * (100 * C)
    yC = (graf_val.ScaleHeight / 100) * (100 * C)
    
    xG = (graf_val.ScaleHeight / 100) * (100 * G)
    yG = (graf_val.ScaleHeight / 100) * (100 * G)
    
    graf_val.DrawWidth = 4
    
    'If i > 1 Then
        If PlotL(0).Value = 1 Then graf_val.Line (oldn, revers - oldyA)-((graf_val.ScaleWidth / chain) * k, revers - yA), &H40C0&
        If PlotL(1).Value = 1 Then graf_val.Line (oldn, revers - oldyT)-((graf_val.ScaleWidth / chain) * k, revers - yT), &H808000
        If PlotL(2).Value = 1 Then graf_val.Line (oldn, revers - oldxC)-((graf_val.ScaleWidth / chain) * k, revers - xC), &H404040
        If PlotL(3).Value = 1 Then graf_val.Line (oldn, revers - oldxG)-((graf_val.ScaleWidth / chain) * k, revers - xG), &HC0&
    'End If

    oldn = (graf_val.ScaleWidth / chain) * k

    oldxA = xA
    oldyA = yA

    oldxT = xT
    oldyT = yT
    
    oldxC = xC
    oldyC = yC
    
    oldxG = xG
    oldyG = yG

Next k


End Sub


Function Day(ByRef v() As Variant)

For i = 0 To UBound(v)

    If v(i) > old Then
        x = v(i)
        h = i
    End If
    
    old = x

Next i

    If h = 0 Then n = "A"
    If h = 1 Then n = "T"
    If h = 2 Then n = "G"
    If h = 3 Then n = "C"
        
Day = n

End Function





Function draw_scale(ByVal k_stat As Integer)
Dim zx, qx, zy, qy As Variant
Dim sp As Variant
Dim i As Integer

Form1.Cls

'X axis on graf_val OBJ
'-------------------------------------
sp = graf_val.ScaleWidth / k_stat
For i = 0 To k_stat

    zx = graf_val.Left + (sp * i)
    qx = zx
    zy = graf_val.Top + graf_val.ScaleHeight
    qy = graf_val.Top + graf_val.ScaleHeight + 6

    If k_stat < 10 Then
        Form1.CurrentX = zx - 6
        Form1.CurrentY = qy
        Form1.Print "S" & i
    End If

    Form1.Line (zx, zy)-(qx, qy), &H808080

Next i
'-------------------------------------

'Y axis on graf_val OBJ
'-------------------------------------
    zx = graf_val.Left - 6
    qx = graf_val.Left
    zy = graf_val.Top
    qy = zy
    Form1.Line (zx, zy)-(qx, qy), &H808080
    Form1.CurrentX = zx - 7
    Form1.CurrentY = qy - 6
    Form1.Print "1"

    zx = graf_val.Left - 6
    qx = graf_val.Left
    zy = graf_val.Top + graf_val.ScaleHeight
    qy = zy
    Form1.Line (zx, zy)-(qx, qy), &H808080
    Form1.CurrentX = zx - 7
    Form1.CurrentY = qy - 6
    Form1.Print "0"
'-------------------------------------

End Function

Private Sub Go_Click()
Dim Trow(1 To 4) As String

Trow(1) = Val(MCell(0).Text) + Val(MCell(1).Text) + Val(MCell(2).Text) + Val(MCell(3).Text)
Trow(2) = Val(MCell(4).Text) + Val(MCell(5).Text) + Val(MCell(6).Text) + Val(MCell(7).Text)
Trow(3) = Val(MCell(8).Text) + Val(MCell(9).Text) + Val(MCell(10).Text) + Val(MCell(11).Text)
Trow(4) = Val(MCell(12).Text) + Val(MCell(13).Text) + Val(MCell(14).Text) + Val(MCell(15).Text)

For i = 1 To UBound(Trow)
    If Val(Trow(i)) > 0.98 And Val(Trow(i)) <= 1 Then
        ElseIf Val(Trow(i)) = 0 Then
        Else
            MsgBox "The values from row " & i & " of the transition matrix do not" & vbCrLf & "sum up to 1 (or close: ex. 0.99). Check the values from row " & i
            Exit Sub
    End If
Next i

graf_val.Cls
Form_Load
End Sub

Private Sub StartFrom_Click(Index As Integer)
    If Index = 0 Then VComponent(0).Text = "1" Else VComponent(0).Text = "0"
    If Index = 1 Then VComponent(1).Text = "1" Else VComponent(1).Text = "0"
    If Index = 2 Then VComponent(2).Text = "1" Else VComponent(2).Text = "0"
    If Index = 3 Then VComponent(3).Text = "1" Else VComponent(3).Text = "0"
End Sub
