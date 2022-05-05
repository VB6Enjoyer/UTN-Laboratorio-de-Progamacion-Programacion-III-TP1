VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Super Cuadros '98"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label labCreditos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Créditos"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   3960
      Width           =   5760
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
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label labSalir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   6240
      Width           =   3285
   End
   Begin VB.Label labJugar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jugar"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   3960
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labJugar.ForeColor = vbWhite
    labSalir.ForeColor = vbWhite
    labCreditos.ForeColor = vbWhite
End Sub

Private Sub labCreditos_Click()
    Creditos.Show
    Hide
End Sub

Private Sub labCreditos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCreditos.ForeColor = &HFFC0C0
End Sub

Private Sub labJugar_Click()
    Niveles.Show
    Hide
End Sub

Private Sub labJugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labJugar.ForeColor = &HFFC0C0
End Sub

Private Sub labSalir_Click()
    End
End Sub

Private Sub labSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labSalir.ForeColor = &HFFC0C0
End Sub
