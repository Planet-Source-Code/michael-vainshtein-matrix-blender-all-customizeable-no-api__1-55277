VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "ABOUT MEEEEEEEEEEEEEEEE"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAbout.frx":0000
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Timer tmrMoveMatrix 
      Interval        =   100
      Left            =   2640
      Top             =   2280
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I
Dim X(10)
Dim K(10)
Dim y0(10)

Private Sub Form_Load()
    Dim B
    For B = 0 To UBound(X)
        Randomize
        X(B) = Rnd * Width
        y0(B) = Rnd * 17
    Next
End Sub

Private Sub tmrMoveMatrix_Timer()
    Static B
    Cls
    
    ForeColor = &H8000&
    For B = 0 To UBound(X)
        K(B) = K(B) + 1
        For I = 0 To 10
            FontBold = False
            CurrentY = X(B)
            CurrentX = (K(B) - y0(B) + I) * TextHeight("A") - 10 * TextHeight("A")
            Randomize
            Print Chr(Rnd * 42 + 48)
        Next
        If (K(B) - y0(B) + 10) * TextHeight("A") > Height + 10 * TextHeight("A") * 2 Then
            K(B) = 0
            Randomize
            X(B) = Rnd * Me.Width
            y0(B) = Rnd * 17
        End If
    Next
    CurrentX = 800
    CurrentY = 10 * TextHeight("A")
    FontBold = True
    ForeColor = RGB(255, 255, 0)
    Print "M a d e   b y   M i c h a e l   V a i n s h t e i n"
    CurrentY = 11 * TextHeight("A")
    Print "I've put some work into this thing least thing you can do is go to"
    CurrentY = 12 * TextHeight("A")
    CurrentX = 2200
    Print "PSC and vote!"
    

End Sub
