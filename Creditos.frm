VERSION 5.00
Begin VB.Form Creditos 
   BackColor       =   &H000080FF&
   Caption         =   "Super Cuadros '98"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   13035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   12600
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Left            =   12600
      Top             =   120
   End
   Begin VB.Label labEpilepsia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Epilepsia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   885
   End
   Begin VB.Label labMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Volver al menú"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1350
      Left            =   3120
      TabIndex        =   4
      Top             =   6720
      Width           =   6840
   End
   Begin VB.Label labAA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAHHHH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   -9960
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   10185
   End
   Begin VB.Label labGracias 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gracias por jugar!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2025
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   12495
   End
   Begin VB.Label labTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Super Cuadros'98"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label labJuani 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hecho por Juani"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2025
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   11355
   End
End
Attribute VB_Name = "Creditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim A As Integer

Private Sub Form_Load()
    Timer1.Enabled = False
    Timer1.Interval = 10
    Timer2.Enabled = False
    Timer2.Interval = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labMenu.ForeColor = vbWhite
    labEpilepsia.ForeColor = &H1180FF
End Sub

Private Sub labEpilepsia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labEpilepsia.ForeColor = vbWhite
End Sub

Private Sub labEpilepsia_Click()
    If Timer1.Enabled = False Then
        Timer1.Enabled = True
        Timer2.Enabled = True
        labAA.Visible = True
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
        labAA.Visible = False
    End If
    
    Me.BackColor = &H1180FF
    labTitulo.ForeColor = vbWhite
    labJuani.ForeColor = vbWhite
    labGracias.ForeColor = vbWhite
    labMenu.ForeColor = vbWhite
    labTitulo.FontSize = 36
    labJuani.FontSize = 72
    labGracias.FontSize = 72
    labMenu.FontSize = 48
    
    A = -10000
End Sub

Private Sub labMenu_Click()
    Menu.Show
    Unload Me
End Sub

Private Sub labMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labMenu.ForeColor = vbBlack
End Sub

Private Sub Timer1_Timer()
    Me.BackColor = RGB((Rnd * 256), (Rnd * 256), (Rnd * 256))
    labTitulo.ForeColor = RGB((Rnd * 256), (Rnd * 256), (Rnd * 256))
    labJuani.ForeColor = RGB((Rnd * 256), (Rnd * 256), (Rnd * 256))
    labGracias.ForeColor = RGB((Rnd * 256), (Rnd * 256), (Rnd * 256))
    labMenu.ForeColor = RGB((Rnd * 256), (Rnd * 256), (Rnd * 256))
    labTitulo.FontSize = Int((Rnd * 48) + 12)
    labJuani.FontSize = Int((Rnd * 72) + 12)
    labGracias.FontSize = Int((Rnd * 72) + 12)
    labMenu.FontSize = Int((Rnd * 48) + 12)
End Sub

Private Sub Timer2_Timer()
    labAA.Left = A
    If A = 13000 Then
        A = -10000
    End If
    A = A + 10
End Sub
