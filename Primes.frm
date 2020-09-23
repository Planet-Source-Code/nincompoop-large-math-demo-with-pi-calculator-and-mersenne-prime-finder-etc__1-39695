VERSION 5.00
Begin VB.Form frmPrimes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Primes Checker"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
      Height          =   315
      Left            =   4538
      TabIndex        =   2
      Top             =   833
      Width           =   1215
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   285
      TabIndex        =   0
      Top             =   458
      Width           =   9720
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number to be checked :"
      Height          =   255
      Left            =   3878
      TabIndex        =   1
      Top             =   113
      Width           =   2535
   End
End
Attribute VB_Name = "frmPrimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()

    Dim Root() As Integer
    Dim UBRoot As Long
    Dim Num() As Integer
    Dim UBNum As Long
    Dim Div() As Integer
    Dim UBDiv As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim Two(0) As Integer
    Dim FileNo As Integer
    Dim LastTime As Long
    Dim curNum As String
    Dim IsPrime As Boolean

    LastTime = timeGetTime
    On Error GoTo FileErr
    StringToArray txtNum, Num, UBNum
    FileNo = FreeFile
    Open "Primes.txt" For Input Access Read Lock Write As FileNo
    Sqrt Num, UBNum, Root, UBRoot
    IsPrime = True
    Do While Not EOF(FileNo)
        Input #FileNo, curNum
        StringToArray curNum, Div, UBDiv
        If ArrayCmp(Div, UBDiv, Root, UBRoot) = 1 Then Exit Do
        If ArrayDivide(Num, UBNum, Div, UBDiv, Q, UBQ, R, UBR) Then
            IsPrime = False
            Exit Do
        End If
    Loop
    Close #FileNo
    If ArrayCmp(Div, UBDiv, Root, UBRoot) <> 1 And IsPrime Then
        Two(0) = 2
        Do While 1
            AddArray2 Div, UBDiv, Two, 0
            If ArrayCmp(Div, UBDiv, Root, UBRoot) = 1 Then Exit Do
            If ArrayDivide(Num, UBNum, Div, UBDiv, Q, UBQ, R, UBR) Then
                IsPrime = False
                Exit Do
            End If
        Loop
    End If
    LastTime = timeGetTime - LastTime
    If IsPrime Then
        MsgBox "The number is Prime!"
    Else
        MsgBox "The number is Not Prime!" + "Divisible by : " + ArrayToString(Div, UBDiv)
    End If
    MsgBox "Time taken = " + Format$(LastTime, "0 ms")
    Exit Sub

FileErr:
    Open "Primes.txt" For Output Access Write Lock Write As FileNo
    Print #FileNo, "2"
    Print #FileNo, "3"
    Close #FileNo
    Resume

End Sub
