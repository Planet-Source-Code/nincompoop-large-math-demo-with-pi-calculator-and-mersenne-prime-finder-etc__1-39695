VERSION 5.00
Begin VB.Form frmGenSmall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Small Prime Numbers Generator"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGen 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   2063
      TabIndex        =   2
      Top             =   1110
      Width           =   1455
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1823
      MaxLength       =   10
      TabIndex        =   1
      Top             =   615
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number upto which Primes should be generated : (Limit = 2^31-1)"
      Height          =   255
      Left            =   173
      TabIndex        =   0
      Top             =   150
      Width           =   5535
   End
End
Attribute VB_Name = "frmGenSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub cmdGen_Click()

    Dim Num As Long
    Dim Root As Long
    Dim Div As Long
    Dim Prime As Long
    Dim FileNoR As Integer
    Dim FileNoW As Integer
    Dim curNum As String
    Dim IsPrime As Boolean
    Dim LastTime As Long

    LastTime = timeGetTime

    Num = txtNum
    FileNoR = FreeFile
    Open "PrimesR.txt" For Input Access Read Lock Write As FileNoR
    FileNoW = FreeFile
    Open "PrimesW.txt" For Input Access Read Lock Write As FileNoW
    Do While Not EOF(FileNoW)
        Input #FileNoW, curNum
    Loop
    Close #FileNoW
    FileNoW = FreeFile
    Open "PrimesW.txt" For Append Access Write Lock Write As FileNoW
    Prime = curNum
    Do While 1
        Prime = Prime + 2
        If Prime > Num Then Exit Do
        Seek #FileNoR, 1
        Root = Fix(Sqr(Prime))
        IsPrime = True
        Do While 1
            Input #FileNoR, curNum
            Div = curNum
            If Div > Root Then Exit Do
            If Prime Mod Div = 0 Then
                IsPrime = False
                Exit Do
            End If
        Loop
        If IsPrime Then
            Print #FileNoW, Format$(Prime, "0")
        End If
    Loop
    Close #FileNoR
    Close #FileNoW
    MsgBox "Time Taken : " + Format$(timeGetTime - LastTime, "0 ms")

End Sub
