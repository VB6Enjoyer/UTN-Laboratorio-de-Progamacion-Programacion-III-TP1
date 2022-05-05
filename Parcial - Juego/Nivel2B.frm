VERSION 5.00
Begin VB.Form Nivel2B 
   BackColor       =   &H00404040&
   Caption         =   "Super Cuadros '98"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   6840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H000000FF&
         Height          =   735
         Index           =   0
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   71
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FF00&
         Height          =   735
         Index           =   1
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   70
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H000000FF&
         Height          =   735
         Index           =   2
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   69
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FF00&
         Height          =   735
         Index           =   3
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   68
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H000000FF&
         Height          =   735
         Index           =   4
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   67
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FF00&
         Height          =   735
         Index           =   5
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   66
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   6
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   65
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   7
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   64
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF0000&
         Height          =   735
         Index           =   8
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   63
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF00FF&
         Height          =   735
         Index           =   9
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   62
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF00FF&
         Height          =   735
         Index           =   10
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   61
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF0000&
         Height          =   735
         Index           =   11
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   60
         Top             =   3720
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         Caption         =   "Colores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   5295
         Left            =   7680
         TabIndex        =   54
         Top             =   240
         Width           =   1215
         Begin VB.PictureBox pctColor 
            BackColor       =   &H000000FF&
            Height          =   735
            Index           =   0
            Left            =   240
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   59
            Top             =   480
            Width           =   735
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H0000FF00&
            Height          =   735
            Index           =   1
            Left            =   240
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   58
            Top             =   1440
            Width           =   735
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H0000FFFF&
            Height          =   735
            Index           =   2
            Left            =   240
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   57
            Top             =   2400
            Width           =   735
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00FF0000&
            Height          =   735
            Index           =   3
            Left            =   240
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   56
            Top             =   3360
            Width           =   735
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00FF00FF&
            Height          =   735
            Index           =   4
            Left            =   240
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   55
            Top             =   4320
            Width           =   735
         End
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H000000FF&
         Height          =   735
         Index           =   12
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   53
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FF00&
         Height          =   735
         Index           =   13
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   52
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF0000&
         Height          =   735
         Index           =   14
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   51
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FF00&
         Height          =   735
         Index           =   15
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   50
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H000000FF&
         Height          =   735
         Index           =   16
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   49
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF0000&
         Height          =   735
         Index           =   17
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   48
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   18
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   47
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   19
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   46
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF0000&
         Height          =   735
         Index           =   20
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   45
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF00FF&
         Height          =   735
         Index           =   21
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   44
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF00FF&
         Height          =   735
         Index           =   22
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   43
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H00FF0000&
         Height          =   735
         Index           =   23
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   42
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   24
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   41
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   25
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   40
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   26
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   39
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   27
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   38
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   28
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   37
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   29
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   36
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   30
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   35
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   31
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   32
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   33
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   33
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   32
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   34
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   31
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   35
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   30
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   36
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   29
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   37
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   38
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   39
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   26
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   40
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   25
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   41
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   24
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H000000FF&
         Height          =   735
         Index           =   42
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   23
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   43
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   22
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H000000FF&
         Height          =   735
         Index           =   44
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   21
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   45
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   20
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   46
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   19
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   47
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   18
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   48
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   17
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   49
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   50
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   51
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   14
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   52
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   13
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   53
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   12
         Top             =   3000
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   54
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   11
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   55
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   10
         Top             =   4440
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FF00&
         Height          =   735
         Index           =   56
         Left            =   1680
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   9
         Top             =   5160
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   57
         Left            =   2400
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   8
         Top             =   5160
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FF00&
         Height          =   735
         Index           =   58
         Left            =   3120
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   7
         Top             =   5160
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         BackColor       =   &H0000FFFF&
         Height          =   735
         Index           =   59
         Left            =   3840
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   6
         Top             =   5160
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   60
         Left            =   4560
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   5
         Top             =   5160
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   61
         Left            =   5280
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   4
         Top             =   5160
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   62
         Left            =   6000
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   3
         Top             =   5160
         Width           =   735
      End
      Begin VB.PictureBox pctCuadro 
         Height          =   735
         Index           =   63
         Left            =   6720
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   2
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label labControles 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Controles"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7560
         TabIndex        =   78
         Top             =   5520
         Width           =   1590
      End
   End
   Begin VB.CommandButton btnNiveles 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Niveles"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volver al menú de niveles"
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label labTiempo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7680
      TabIndex        =   77
      Top             =   6360
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   6000
      TabIndex        =   76
      Top             =   6240
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Intentos:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   3240
      TabIndex        =   75
      Top             =   6240
      Width           =   1920
   End
   Begin VB.Label labIntentos 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   480
      Left            =   5160
      TabIndex        =   74
      Top             =   6360
      Width           =   225
   End
   Begin VB.Label labAciertos 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   480
      Left            =   2040
      TabIndex        =   73
      Top             =   6360
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Aciertos:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   120
      TabIndex        =   72
      Top             =   6240
      Width           =   1890
   End
End
Attribute VB_Name = "Nivel2B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Aciertos As Integer
Dim Intentos As Integer
Dim Color As Integer
Dim Tiempo As Integer

Private Sub btnNiveles_Click()
    Niveles.Show
    Unload Me
    Set Nivel2B = Nothing
End Sub

Private Sub Form_Load()
    Tiempo = 0
    Timer1.Enabled = False
    Timer1.Interval = 1000
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labControles.ForeColor = vbWhite
End Sub

Private Sub labControles_Click()
    MsgBox ("Mirá, los controles del juego son muy intuitivos, pero voy a asumir que presionaste este botón por error y decirtelo de todas maneras: tenés que agarrar un color y arrastrarlo a la casilla que espeje dicho color en el lado derecho. Listo, andá a jugar ahora")
    labControles.ForeColor = vbWhite
End Sub

Private Sub labControles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labControles.ForeColor = vbRed
End Sub

Private Sub pctColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pctColor(Index).Drag vbBeginDrag
    Color = Index
End Sub

Private Sub pctCuadro_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    pctCuadro(Index).Drag vbEndDrag
    
    If Timer1.Enabled = False Then Timer1.Enabled = True
    
    Select Case Index
       'ROJO
       '------------------------------------------------------------
        Case 30, 34, 47, 49, 51, 53, 55
            If Color = 0 Then
                If pctCuadro(Index).BackColor <> vbRed Then
                    pctCuadro(Index).BackColor = vbRed
                    Aciertos = Aciertos + 1
                End If
            End If
            Intentos = Intentos + 1
       'VERDE
       '------------------------------------------------------------
        Case 31, 33, 50, 52, 54, 61, 63
            If Color = 1 Then
                If pctCuadro(Index).BackColor <> vbGreen Then
                    pctCuadro(Index).BackColor = vbGreen
                    Aciertos = Aciertos + 1
                End If
            End If
            Intentos = Intentos + 1
       'AMARILLO
       '------------------------------------------------------------
        Case 24, 25, 36, 37, 46, 48, 60, 62
            If Color = 2 Then
                If pctCuadro(Index).BackColor <> vbYellow Then
                    pctCuadro(Index).BackColor = vbYellow
                    Aciertos = Aciertos + 1
                End If
            End If
            Intentos = Intentos + 1
       'AZUL
       '------------------------------------------------------------
        Case 26, 29, 32, 35, 38, 41
            If Color = 3 Then
                If pctCuadro(Index).BackColor <> vbBlue Then
                    pctCuadro(Index).BackColor = vbBlue
                    Aciertos = Aciertos + 1
                End If
            End If
            Intentos = Intentos + 1
       'ROSADO
       '------------------------------------------------------------
        Case 27, 28, 39, 40
            If Color = 4 Then
                If pctCuadro(Index).BackColor <> vbMagenta Then
                    pctCuadro(Index).BackColor = vbMagenta
                    Aciertos = Aciertos + 1
                End If
            End If
            Intentos = Intentos + 1
       'ORIGINALES
       '------------------------------------------------------------
        Case 0 To 17, 42, 43, 44, 56, 57, 58
            Intentos = Intentos + 1
    End Select
    
    labAciertos.Caption = Aciertos
    labIntentos.Caption = Intentos
    
    Dim Resp As Integer
    
    If Aciertos >= 32 Then
        Resp = MsgBox("Felicitaciones, completaste el Nivel 2. Presiona OK para ir al siguiente nivel", vbOKOnly, "Nivel Completado")
        'Carga el siguiente nivel
        Nivel3.Show
        Unload Me
        Set Nivel2B = Nothing
    End If
End Sub

Private Sub Timer1_Timer()
    Tiempo = Tiempo + 1
    labTiempo.Caption = Tiempo
End Sub
