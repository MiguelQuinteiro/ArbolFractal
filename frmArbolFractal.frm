VERSION 5.00
Begin VB.Form frmArbolFractal 
   BackColor       =   &H00000000&
   Caption         =   "Árbol Fractal"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " Parámetros "
      Height          =   8175
      Left            =   10200
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdIncremento 
         Caption         =   "Incremento"
         Height          =   495
         Left            =   1080
         TabIndex        =   16
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox txtIncremento 
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Text            =   "1"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtEsparcir 
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Text            =   "50.5"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtProfundidad 
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Text            =   "10"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtAngulo 
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Text            =   "270"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtTamaño 
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Text            =   "1100"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtY1 
         Height          =   495
         Left            =   1680
         TabIndex        =   3
         Text            =   "8000"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtX1 
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Text            =   "5000"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdArbolFractal 
         Caption         =   "Árbol Fractal"
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Incremento"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Esparcir"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Profundidad"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Ángulo"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tamaño"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Y1"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "X1"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmArbolFractal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miX As Double
Dim miY As Double
Dim miTamaño As Double
Dim miAngulo As Double
Dim miEsparcir As Integer
Dim miEscala As Double
Dim miProfundidad As Integer

Private Sub cmdArbolFractal_Click()
  Cls
  miX = Val(txtX1.Text)
  miY = Val(txtY1.Text)
  miTamaño = Val(txtTamaño.Text)
  miAngulo = Val(txtAngulo.Text) * 3.1415 / 180  ' 86.401
  miProfundidad = Val(txtProfundidad.Text)  ' 9
  miEsparcir = Val(txtEsparcir.Text)  ' 50.5
  miEscala = 0.9
  Call Rama(miX, miY, miTamaño, miAngulo, miProfundidad)
End Sub

Public Sub Rama(X1, Y1, Tamaño, Angulo, Profundidad)
  Dim x2 As Double
  Dim y2 As Double
  x2 = X1 + Tamaño * Cos(Angulo) * 1#
  y2 = Y1 + Tamaño * Sin(Angulo) * 1#
  Line (X1, Y1)-(x2, y2), vbWhite
  If Profundidad > 0 Then
    Call Rama(x2, y2, Tamaño * miEscala, (Angulo - miEsparcir) * 1, Profundidad - 1)
    Call Rama(x2, y2, Tamaño * miEscala, (Angulo + miEsparcir) * 1, Profundidad - 1)
  End If
End Sub

Private Sub cmdIncremento_Click()
  txtEsparcir.Text = miEsparcir + Val(txtIncremento.Text)
  Call cmdArbolFractal_Click
End Sub
