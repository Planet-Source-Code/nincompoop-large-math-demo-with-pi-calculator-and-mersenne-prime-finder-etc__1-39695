VERSION 5.00
Begin VB.Form frmMersenne 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mersenne Prime Finder"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLLTest 
      Caption         =   "&LLTest"
      Default         =   -1  'True
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtResult 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
   Begin VB.TextBox txtNum 
      Height          =   7095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   10215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
      Height          =   375
      Left            =   6255
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   7515
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblExp 
      Caption         =   "Exponent to be Tested"
      Height          =   255
      Left            =   5595
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMersenne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NumOfDiv As Long = 14

Private Sub cmdCheck_Click()

    Dim Num() As Integer
    Dim UBNum As Long
    Dim Root() As Integer
    Dim UBRoot As Long
    Dim Div() As Integer
    Dim UBDiv As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim Delta() As Integer
    Dim UBDelta As Long
    Dim DivArr(NumOfDiv) As Boolean
    Dim One(0) As Integer
    Dim bMersenne As Boolean
    Dim LastTime As Long
    Dim i As Long

    LastTime = timeGetTime()
    StringToArray txtExp, Num, UBNum
    If Not IsArrayPrime(Num, UBNum) Then
        txtResult = "The Number is not Prime!" + vbCrLf + "Beacuse The exponent is not Prime!!"
        Exit Sub
    End If

    One(0) = 1
    PowerOf2 txtExp, Num, UBNum
    SubArray2 Num, UBNum, One, 0
    txtNum = ArrayToString(Num, UBNum)
    DoEvents
    Sqrt Num, UBNum, Root, UBRoot

    StringToArray txtExp, Delta, UBDelta
    AddArray2 Delta, UBDelta, Delta, UBDelta

    ReDim Div(0)
    Div(0) = 1

    One(0) = 3
    For i = 0 To NumOfDiv
        AddArray2 Div, UBDiv, Delta, UBDelta
        If Not ArrayDivide(Div, UBDiv, One, 0, Q, UBQ, R, UBR) Then
            If Div(0) Mod 5 Then
                DivArr(i) = True
            End If
        End If
    Next i

    UnityArray Div, UBDiv
    bMersenne = True
    Do While 1
        For i = 0 To NumOfDiv
            AddArray2 Div, UBDiv, Delta, UBDelta
            If (ArrayCmp(Div, UBDiv, Root, UBRoot) > 0) Then Exit Do
            If DivArr(i) Then
                If ((Div(0) Mod 8) = 1) Or ((Div(0) Mod 8) = 7) Then
                    If ArrayDivide(Num, UBNum, Div, UBDiv, Q, UBQ, R, UBR) Then
                        bMersenne = False
                        Exit Do
                    End If
                End If
            End If
        Next i
    Loop

    If bMersenne Then
        txtResult = "Congratulations!!" + vbCrLf + "The number is a Mersenne Prime"
    Else
        txtResult = "Sorry!!" + vbCrLf + "The number is divisible by :-" + vbCrLf + ArrayToString(Div, UBDiv)
    End If
    MsgBox "Time Taken = " + Str$(timeGetTime - LastTime) + "ms"

End Sub

Private Sub cmdLLTest_Click()

    Dim Num() As Integer
    Dim UBNum As Long
    Dim LastTime As Long
    Dim One(0) As Integer

    LastTime = timeGetTime()
    If IsMersennePrimeExp(txtExp) Then
        One(0) = 1
        PowerOf2 txtExp, Num, UBNum
        SubArray2 Num, UBNum, One, 0
        txtNum = ArrayToString(Num, UBNum)
        txtResult = "Congratulations!!" + vbCrLf + "The number is a Mersenne Prime."
    Else
        txtResult = "Sorry!!" + vbCrLf + "The number is not Prime."
    End If
    MsgBox "Time Taken = " + Str$(timeGetTime - LastTime) + "ms"

End Sub
