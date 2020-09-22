VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4425
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "step"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.HScrollBar blue 
      Height          =   255
      Left            =   720
      Max             =   255
      TabIndex        =   10
      Top             =   2760
      Width           =   2730
   End
   Begin VB.HScrollBar green 
      Height          =   255
      Left            =   720
      Max             =   255
      TabIndex        =   9
      Top             =   2400
      Value           =   255
      Width           =   2730
   End
   Begin VB.HScrollBar red 
      Height          =   255
      Left            =   720
      Max             =   255
      TabIndex        =   8
      Top             =   2040
      Width           =   2730
   End
   Begin VB.PictureBox Picture5 
      Height          =   315
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   4020
      TabIndex        =   4
      Top             =   1560
      Width           =   4080
   End
   Begin VB.PictureBox Picture4 
      Height          =   300
      Left            =   180
      ScaleHeight     =   240
      ScaleWidth      =   4020
      TabIndex        =   3
      Top             =   1185
      Width           =   4080
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Left            =   180
      ScaleHeight     =   225
      ScaleWidth      =   4020
      TabIndex        =   2
      Top             =   825
      Width           =   4080
   End
   Begin VB.PictureBox Picture2 
      Height          =   300
      Left            =   180
      ScaleHeight     =   240
      ScaleWidth      =   4020
      TabIndex        =   1
      Top             =   450
      Width           =   4080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3720
      Top             =   2280
   End
   Begin VB.PictureBox Picture1 
      Height          =   270
      Left            =   180
      ScaleHeight     =   210
      ScaleWidth      =   4020
      TabIndex        =   0
      Top             =   135
      Width           =   4080
   End
   Begin VB.Label Label3 
      Caption         =   "blue"
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   2760
      Width           =   330
   End
   Begin VB.Label Label2 
      Caption         =   "green"
      Height          =   270
      Left            =   180
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "red:"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dm
Dim inc As Single
Dim av As Single


Public Sub UpdateProgress(pBar As PictureBox, Percent As Single, Optional ByVal ShowPercent = False, Optional ByVal c = 0, Optional ByVal r = 255, Optional ByVal g = 255, Optional ByVal b = 255, Optional ByVal ste = 0.1)
    Dim sData As String


    If Not pBar.AutoRedraw Then '//no memory?
        pBar.AutoRedraw = -1 'nope, we make some
    End If
    pBar.Cls '//clear memory
    pBar.ScaleWidth = 100 'set scale's width



    If ShowPercent = True Then
        sData$ = Format(Percent, "###0") + "%" '//format percent
        pBar.CurrentX = 50 - pBar.TextWidth(sData$) 'set the currentX for the progress
        pBar.CurrentY = pBar.ScaleHeight - pBar.TextHeight(sData$) 'set the currentY for the progress
        Debug.Print sData$ 'show percent in debug window
    End If
    For valc = 0 To Percent Step ste
     If c = 0 Then pBar.Line (valc, 0)-(valc, pBar.ScaleHeight), RGB(r, g, b)
2    If c = 1 Then pBar.Line (valc, 0)-(valc, pBar.ScaleHeight), RGB(valc * 2.55, g, b)
    If c = 2 Then pBar.Line (valc, 0)-(valc, pBar.ScaleHeight), RGB(r, valc * 2.55, b)
    If c = 3 Then pBar.Line (valc, 0)-(valc, pBar.ScaleHeight), RGB(r, g, valc * 2.55)
   
    Next valc
    pBar.Refresh
End Sub





Private Sub Command1_Click()
av = av + 0.1
Debug.Print av

End Sub

Private Sub Timer1_Timer()
inc = inc + 1
If inc > 100 Then inc = 0

Call UpdateProgress(Picture1, inc, False, 0, 0, 0, 0, 0.2 + av)
Call UpdateProgress(Picture2, inc, False, 1, 0, 0, 0, 0.2 + av)
Call UpdateProgress(Picture3, inc, False, 2, 0, 0, 0, 0.2 + av)
Call UpdateProgress(Picture4, inc, False, 3, 0, 0, 0, 0.2 + av)
Call UpdateProgress(Picture5, inc, False, 0, red.Value, green.Value, blue.Value, 0.2 + av)

End Sub
