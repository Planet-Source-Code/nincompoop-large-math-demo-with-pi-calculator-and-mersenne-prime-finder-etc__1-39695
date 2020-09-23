VERSION 5.00
Begin VB.Form frmGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prime Numbers Generator"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGen 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   4433
      TabIndex        =   2
      Top             =   1110
      Width           =   1455
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   173
      TabIndex        =   1
      Top             =   615
      Width           =   9975
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number upto which Primes should be generated :"
      Height          =   255
      Left            =   2993
      TabIndex        =   0
      Top             =   150
      Width           =   4335
   End
End
Attribute VB_Name = "frmGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGen_Click()

    Dim Num() As Integer
    Dim UBNum As Long
    Dim Root() As Integer
    Dim UBRoot As Long
    Dim Div() As Integer
    Dim UBDiv As Long
    Dim Prime() As Integer
    Dim UBPrime As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim Two(0) As Integer
    Dim FileNo As Integer
    Dim curNum As String
    Dim IsPrime As Boolean
    Dim LastTime As Long

    LastTime = timeGetTime
    On Error GoTo FileErr

    Two(0) = 2
    StringToArray txtNum, Num, UBNum
    FileNo = FreeFile
    Open "Primes.txt" For Input Access Read Lock Write As FileNo
    Do While Not EOF(FileNo)
        Input #FileNo, curNum
    Loop
    StringToArray curNum, Prime, UBPrime
    Do While 1
        AddArray2 Prime, UBPrime, Two, 0
        If ArrayCmp(Num, UBNum, Prime, UBPrime) = -1 Then Exit Do
        Seek #FileNo, 1
        Sqrt Prime, UBPrime, Root, UBRoot
        IsPrime = True
        Do While 1
            Input #FileNo, curNum
            StringToArray curNum, Div, UBDiv
            If ArrayCmp(Div, UBDiv, Root, UBRoot) = 1 Then Exit Do
            If ArrayDivide(Prime, UBPrime, Div, UBDiv, Q, UBQ, R, UBR) Then
                IsPrime = False
                Exit Do
            End If
        Loop
        If IsPrime Then
            Close #FileNo
            Open "Primes.txt" For Append Access Write Lock Write As FileNo
            Print #FileNo, ArrayToString(Prime, UBPrime)
            Close #FileNo
            Open "Primes.txt" For Input Access Read Lock Write As FileNo
        End If
    Loop
    Close #FileNo
    MsgBox "Time Taken : " + Format$(timeGetTime - LastTime, "0 ms")
    Exit Sub

FileErr:
    Open "Primes.txt" For Output Access Write Lock Write As FileNo
    Print #FileNo, "2"
    Print #FileNo, "3"
    Close #FileNo
    Resume

End Sub
