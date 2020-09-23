VERSION 5.00
Begin VB.Form frmnThPrime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "nTh Prime"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWanted 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   495
      Left            =   1733
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtnThPrime 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtN 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Prime ="
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Found n ="
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Wanted n ="
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmnThPrime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bQuit As Boolean
Dim bRun As Boolean

Private Sub cmdGo_Click()

    Dim Num As Long
    Dim LastNum As Long
    Dim FileNoW As Integer
    Dim sNum As String
    Dim Count As Long
    Dim NumWanted As Long

    FileNoW = FreeFile
    Open "PrimesW.txt" For Input Access Read Lock Write As FileNoW
    Input #FileNoW, sNum
    LastNum = sNum
    Count = 1
    NumWanted = txtWanted
    bRun = True
    bQuit = False
    Do While (Not EOF(FileNoW) And Not bQuit)
        If Count = NumWanted Then Exit Do
        Input #FileNoW, sNum
        Num = sNum
        If Num < LastNum Then Err.Raise 123456789
        LastNum = Num
        Count = Count + 1
        If Count Mod 10000 = 0 Then
            txtN = Count
            txtnThPrime = Num
            DoEvents
        End If
    Loop
    txtN = Count
    txtnThPrime = Num
    bRun = False
    Close #FileNoW
    If bQuit Then Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If bRun Then Cancel = 1
    bQuit = True

End Sub
