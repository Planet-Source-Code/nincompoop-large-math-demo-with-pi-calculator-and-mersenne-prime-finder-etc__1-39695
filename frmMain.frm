VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Large Math Demo"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrimes 
      Caption         =   "Primes Checker"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton cmdPi 
      Caption         =   "Pi Calculator"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdnTh 
      Caption         =   "nTh Prime"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdMersenne 
      Caption         =   "Mersenne Prime Finder"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton cmdGenSmall 
      Caption         =   "Small Prime Numbers Generator"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Prime Numbers Generator"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   $"frmMain.frx":0000
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label6 
      Caption         =   $"frmMain.frx":00B1
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   5640
      Width           =   6015
   End
   Begin VB.Label Label5 
      Caption         =   $"frmMain.frx":0163
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   4920
      Width           =   6015
   End
   Begin VB.Label Label4 
      Caption         =   "A small utility which allows you to quickly scan PrimesW.txt for nth Prime."
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   4320
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   $"frmMain.frx":01FC
      Height          =   1455
      Left            =   3240
      TabIndex        =   8
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0452
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":0545
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGen_Click()

    frmGen.Show vbModal

End Sub

Private Sub cmdGenSmall_Click()

    frmGenSmall.Show vbModal

End Sub

Private Sub cmdMersenne_Click()

    frmMersenne.Show vbModal

End Sub

Private Sub cmdnTh_Click()

    frmnThPrime.Show vbModal

End Sub

Private Sub cmdPi_Click()

    frmPi.Show vbModal

End Sub

Private Sub cmdPrimes_Click()

    frmPrimes.Show vbModal

End Sub

Private Sub Form_Load()

    MsgBox "If too large values are given, the Application may appear to hang. No mechanism has been put to gracefully handle such conditions in the demo." + vbCrLf + "Try with small values first and then increase them depending on the time taken for previous operation." + vbCrLf + "Sorry for poor interface also, didn't put much efforts to improve it. If you will see Large Math.bas, you will know where did all the effort went." + vbCrLf + "For best results, run in compiled mode. But then you can't break the application out of long calculations."

End Sub
