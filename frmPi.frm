VERSION 5.00
Begin VB.Form frmPi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pi Calculator"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtNDigits 
      Height          =   285
      Left            =   1665
      TabIndex        =   1
      Text            =   "20"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblNDigits 
      Caption         =   "Number Of Digits"
      Height          =   255
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmPi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()

    Dim Pi() As Integer
    Dim UBPi As Long
    Dim Four(0) As Integer
    Dim Five(0) As Integer
    Dim TwoThreeNine(1) As Integer
    Dim Delta() As Integer
    Dim UBDelta As Long
    Dim Pow() As Integer
    Dim UBPow As Long
    Dim tmpAns() As Integer
    Dim UBtmpAns As Long
    Dim tmpArr() As Integer
    Dim UBtmpArr As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim LastTime As Long
    Dim FileNo As Integer
    Dim bZero As Boolean

    LastTime = timeGetTime
    StringToArray txtNDigits, Delta, UBDelta
    Four(0) = 4
    AddArray2 Delta, UBDelta, Four, 0
    Five(0) = 10
    ArrayPower Five, 0, Delta, UBDelta, Q, UBQ
    Five(0) = 239
    ArrayDivide Q, UBQ, Five, 0, tmpArr, UBtmpArr, R, UBR
    ArrayMultiply Four, 0, Q, UBQ, Delta, UBDelta
    Five(0) = 5
    ArrayDivide Delta, UBDelta, Five, 0, tmpAns, UBtmpAns, R, UBR

    ZeroArray Pi, UBPi
    UnityArray Pow, UBPow
    Four(0) = 2
    Five(0) = 25
    TwoThreeNine(0) = 7121
    TwoThreeNine(1) = 5
    Do While True
        SubArray tmpAns, UBtmpAns, tmpArr, UBtmpArr, Delta, UBDelta
        If IsZero(Delta, UBDelta) Then Exit Do
        If (Pow(0) And 3) = 1 Then
            AddArray2 Pi, UBPi, Delta, UBDelta
        Else
            SubArray2 Pi, UBPi, Delta, UBDelta
        End If
        ArrayMultiply tmpAns, UBtmpAns, Pow, UBPow, Q, UBQ
        ArrayMultiply tmpArr, UBtmpArr, Pow, UBPow, R, UBR
        AddArray2 Pow, UBPow, Four, 0
        ArrayDivide Q, UBQ, Pow, UBPow, Delta, UBDelta, tmpAns, UBtmpAns
        ArrayDivide Delta, UBDelta, Five, 0, tmpAns, UBtmpAns, Q, UBQ
        ArrayDivide R, UBR, Pow, UBPow, Delta, UBDelta, tmpArr, UBtmpArr
        ArrayDivide Delta, UBDelta, TwoThreeNine, 1, tmpArr, UBtmpArr, R, UBR
    Loop
    Four(0) = 4
    ArrayMultiply Pi, UBPi, Four, 0, tmpAns, UBtmpAns
    TwoThreeNine(0) = 0
    TwoThreeNine(1) = 1
    ArrayDivide tmpAns, UBtmpAns, TwoThreeNine, 1, Pi, UBPi, R, UBR
    FileNo = FreeFile
    Open "Pi.txt" For Output Access Write Lock Write As FileNo
    Print #FileNo, ArrayToString(Pi, UBPi)
    Close #FileNo
    LastTime = timeGetTime - LastTime
    MsgBox "Time taken = " & LastTime & " ms"

End Sub
