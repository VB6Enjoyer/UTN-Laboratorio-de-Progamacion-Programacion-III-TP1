VERSION 5.00
Begin VB.Form Niveles 
   BackColor       =   &H00404040&
   Caption         =   "Super Cuadros '98"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label labMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menú"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   6840
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar un nivel"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   11115
   End
   Begin VB.Label labNivel3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel 3"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   5040
      Width           =   3330
   End
   Begin VB.Label labNivel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel 2"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   3120
      Width           =   3330
   End
   Begin VB.Label labNivel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel 1"
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
      Left            =   3840
      TabIndex        =   0
      Top             =   1440
      Width           =   3330
   End
End
Attribute VB_Name = "Niveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub labMenu_Click()
    Menu.Show
    Unload Me
End Sub

Private Sub labMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labMenu.ForeColor = &HFFC0C0
End Sub

Private Sub labNivel1_Click()
    Dim i As Integer
    
    Randomize
    i = Int((3 * Rnd) + 1)
    'i = 1 'Debug
    
    Select Case i
        Case 1
            Nivel1A.Show
        Case 2
            Nivel1B.Show
        Case 3
            Nivel1C.Show
        Case Else
            MsgBox ("Ha ocurrido un error")
    End Select
    Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labNivel1.ForeColor = vbWhite
    labNivel2.ForeColor = vbWhite
    labNivel3.ForeColor = vbWhite
    labMenu.ForeColor = vbWhite
End Sub

Private Sub labNivel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labNivel1.ForeColor = &HFFC0C0
End Sub

Private Sub labNivel2_Click()
    Dim i As Integer
    
    Randomize
    i = Int((3 * Rnd) + 1)
    'i = 3 'Debug
    
    Select Case i
        Case 1
            Nivel2A.Show
        Case 2
            Nivel2B.Show
        Case 3
            Nivel2C.Show
        Case Else
            MsgBox ("Ha ocurrido un error")
    End Select
    Me.Hide
End Sub

Private Sub labNivel2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labNivel2.ForeColor = &HFFC0C0
End Sub

Private Sub labNivel3_Click()
    Nivel3.Show
    Me.Hide
End Sub

Private Sub labNivel3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labNivel3.ForeColor = &HFFC0C0
End Sub
