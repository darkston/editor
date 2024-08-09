VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Editor Ex2 Retzu"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Tipoex 
      Height          =   375
      Left            =   9240
      TabIndex        =   75
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "GunBound Dat Editor"
      Height          =   5055
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10575
      Begin VB.Frame Frame9 
         BackColor       =   &H80000000&
         Caption         =   "  "
         Height          =   1575
         Left            =   7680
         TabIndex        =   69
         Top             =   600
         Width           =   2055
         Begin VB.OptionButton Ex2collar 
            BackColor       =   &H80000000&
            Caption         =   "Option1"
            Height          =   255
            Left            =   1080
            TabIndex        =   74
            Top             =   1200
            Width           =   255
         End
         Begin VB.OptionButton Ex2anillo 
            BackColor       =   &H80000000&
            Caption         =   "Option1"
            Height          =   195
            Left            =   360
            TabIndex        =   73
            Top             =   1200
            Width           =   255
         End
         Begin VB.OptionButton Ex2fondo 
            BackColor       =   &H80000000&
            Caption         =   "Option1"
            Height          =   255
            Left            =   1200
            TabIndex        =   72
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton Ex2adorno 
            BackColor       =   &H80000000&
            Caption         =   "Option1"
            Height          =   255
            Left            =   660
            TabIndex        =   71
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton Ex2Especial 
            BackColor       =   &H80000000&
            Height          =   255
            Left            =   160
            TabIndex        =   70
            Top             =   480
            Width           =   255
         End
         Begin VB.Image Image5 
            Height          =   255
            Left            =   1080
            Picture         =   "Form1.frx":08CA
            Top             =   840
            Width           =   300
         End
         Begin VB.Image Image4 
            Height          =   255
            Left            =   360
            Picture         =   "Form1.frx":0E5C
            Top             =   840
            Width           =   300
         End
         Begin VB.Image Image3 
            Height          =   255
            Left            =   1200
            Picture         =   "Form1.frx":13EE
            Top             =   240
            Width           =   300
         End
         Begin VB.Image Image2 
            Height          =   255
            Left            =   600
            Picture         =   "Form1.frx":1980
            Top             =   240
            Width           =   300
         End
         Begin VB.Image Image1 
            Height          =   255
            Left            =   120
            Picture         =   "Form1.frx":1F12
            Top             =   240
            Width           =   300
         End
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         Caption         =   "< < Anterior"
         Height          =   255
         Left            =   7680
         TabIndex        =   64
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         Caption         =   "Proximo > >"
         Height          =   255
         Left            =   8880
         TabIndex        =   63
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Appearance      =   0  'Flat
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   8880
         TabIndex        =   62
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardar 
         Appearance      =   0  'Flat
         Caption         =   "Guardar"
         Height          =   255
         Left            =   7680
         TabIndex        =   61
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Procurar"
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   7800
         TabIndex        =   56
         Top             =   3240
         Width           =   2655
         Begin VB.CommandButton cmdBuscar 
            Appearance      =   0  'Flat
            Caption         =   "Buscar"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Form1.frx":24A4
            Left            =   240
            List            =   "Form1.frx":24AE
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtBuscar 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   58
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtBuscar2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   57
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CheckBox chkAvActGold 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ativar Gold"
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   2040
         TabIndex        =   55
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkAvActCash 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ativar Cash"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   54
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Preços"
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   600
         TabIndex        =   39
         Top             =   2160
         Width           =   6975
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Eterno"
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   4680
            TabIndex        =   50
            Top             =   360
            Width           =   2175
            Begin VB.CheckBox chkAvActIlimit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Ativar Venda Eterna"
               ForeColor       =   &H00FF00FF&
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtAvPriceIlimitG 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               TabIndex        =   52
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtAvPriceIlimitC 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               TabIndex        =   51
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Cash :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   68
               Top             =   1080
               Width           =   555
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Gold :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   67
               Top             =   720
               Width           =   525
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Mensal"
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   2400
            TabIndex        =   46
            Top             =   360
            Width           =   2175
            Begin VB.TextBox txtAvPriceMonthG 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               TabIndex        =   49
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtAvPriceMonthC 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               TabIndex        =   48
               Top             =   1080
               Width           =   855
            End
            Begin VB.CheckBox chkAvActMonth 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Ativar Venda Mensal"
               ForeColor       =   &H0000C000&
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Cash :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   66
               Top             =   1080
               Width           =   555
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Gold :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   65
               Top             =   720
               Width           =   525
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Semanal"
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   2175
            Begin VB.CheckBox chkAvActWeek 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Ativar Venda Semanal"
               ForeColor       =   &H0080C0FF&
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtAvPriceWeekG 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               TabIndex        =   42
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtAvPriceWeekC 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   720
               TabIndex        =   41
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Gold :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   45
               Top             =   720
               Width           =   525
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Cash :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   44
               Top             =   1080
               Width           =   555
            End
         End
      End
      Begin VB.TextBox txtAvDescription 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   1080
         TabIndex        =   37
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtAvName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   35
         Top             =   840
         Width           =   3255
      End
      Begin VB.CheckBox chkAvNew 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Avatar Novo"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   31
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkAvVisible 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Avatar Visivel"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtAvVisible 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6480
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAvNew 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6480
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAvNoImg 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   3600
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtAvNoMenu 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtContador 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   1080
         TabIndex        =   8
         Top             =   4200
         Visible         =   0   'False
         Width           =   5895
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   5280
            Picture         =   "Form1.frx":24C0
            ScaleHeight     =   390
            ScaleWidth      =   390
            TabIndex        =   24
            Top             =   360
            Width           =   390
         End
         Begin VB.TextBox txtAvBunge 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5160
            TabIndex        =   23
            Top             =   840
            Width           =   615
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3840
            Picture         =   "Form1.frx":25C3
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   22
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtAvDefense 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3720
            TabIndex        =   21
            Top             =   840
            Width           =   615
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   3120
            Picture         =   "Form1.frx":271D
            ScaleHeight     =   390
            ScaleWidth      =   330
            TabIndex        =   20
            Top             =   360
            Width           =   330
         End
         Begin VB.TextBox txtAvIDelay 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3000
            TabIndex        =   19
            Top             =   840
            Width           =   615
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   4560
            Picture         =   "Form1.frx":280C
            ScaleHeight     =   405
            ScaleWidth      =   390
            TabIndex        =   18
            Top             =   360
            Width           =   390
         End
         Begin VB.TextBox txtAvPopularity 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4440
            TabIndex        =   17
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtAvShield 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2280
            TabIndex        =   16
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtAvHealt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1560
            TabIndex        =   15
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtAvAttack 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   840
            TabIndex        =   14
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtAvSDelay 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000001&
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   615
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   960
            Picture         =   "Form1.frx":292B
            ScaleHeight     =   375
            ScaleWidth      =   330
            TabIndex        =   12
            Top             =   360
            Width           =   330
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   1680
            Picture         =   "Form1.frx":2A55
            ScaleHeight     =   390
            ScaleWidth      =   405
            TabIndex        =   11
            Top             =   360
            Width           =   405
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   2400
            Picture         =   "Form1.frx":2B8B
            ScaleHeight     =   390
            ScaleWidth      =   390
            TabIndex        =   10
            Top             =   360
            Width           =   390
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            Picture         =   "Form1.frx":2C9D
            ScaleHeight     =   375
            ScaleWidth      =   300
            TabIndex        =   9
            Top             =   360
            Width           =   300
         End
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Descrição :"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nome do Avatar:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Número da IMG :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ordem :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox txtAvActCash 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAvActGold 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAvActIlimit 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAvActMonth 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAvActWeek 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFileTitle 
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog dlgAbrir 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LogonUser Lib "advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32" (ByVal hToken As Long) As Long
' CÓDIGO ESCRITO POR LUIZ FELIPE E HADA
' ISHADOWN@ LIVE.COM !

Private Type StructAvatars
'    AvContador As String * 4
    AvNoMenu As String * 4
    AvNoMenu0 As String * 1
    
    AVEx2  As String * 1
    AVEx20  As String * 1 ' defecto 4
    AVEx21  As String * 1
    
    AvNoImg As String * 4
    
    AvNew As String * 1
    AvNew0 As String * 7
    
    AvName As String * 20
    
    AvVisible0 As String * 3
    AvVisible As String * 1
    
    AvUnknow As String * 1
    
    AvActWeek As String * 1
    AvActWeek0 As String * 2
    AvPriceWeekG As String * 4
    AvPriceWeekC As String * 4
    
    AvActN1 As String * 1
    AvActN10 As String * 3
    AvPriceN1G As String * 4
    AvPriceN1C As String * 4
    
    AvActMonth As String * 1
    AvActMonth0 As String * 3
    AvPriceMonthG As String * 4
    AvPriceMonthC As String * 4
    
    AvActN2 As String * 1
    AvActN20 As String * 3
    AvPriceN2G As String * 4
    AvPriceN2C As String * 4
    
    AvActIlimit As String * 1
    AvActIlimit0 As String * 3
    AvPriceIlimitG As String * 4
    AvPriceIlimitC As String * 4
    
    AvActGold As String * 1
    AvActCash As String * 1
    AvGoldCash0 As String * 2
    
    Avvacio As String * 12
    AvFF As String * 1
    AvFF1 As String * 1
    AvFF2 As String * 1
    AvFF3 As String * 1
    AvFF4 As String * 1
    AvFF5 As String * 1
    AvFF6 As String * 1
    AvFF7 As String * 1
    AvFF8 As String * 1
    AvFF9 As String * 1
    AvFF10 As String * 1
    AvFF11 As String * 1
    AvFF12 As String * 1
    AvFF13 As String * 1
    AvFF14 As String * 1
    AvFF15 As String * 1
    AvFF16 As String * 1
    AvFF17 As String * 1
    AvFF18 As String * 1
    AvFF19 As String * 1
    AvFF20 As String * 1
    AvFF21 As String * 1
    AvFF22 As String * 1
    AvFF23 As String * 1
    AvFF24 As String * 1
    AvFF25 As String * 1
    AvFF26 As String * 1
    AvFF27 As String * 1
    AvFF28 As String * 1
    AvFF29 As String * 1
    AvFF30 As String * 1
    AvFF31 As String * 1
    AvFF32 As String * 1
    AvFF33 As String * 1
    AvFF34 As String * 1
    AvFF35 As String * 1
    AvFF36 As String * 1
    AvFF37 As String * 1
    AvFF38 As String * 1
    AvFF39 As String * 1
    AvFF40 As String * 1
    AvFF41 As String * 1
    AvFF42 As String * 1
    AvFF43 As String * 1
    AvFF44 As String * 1
    AvFF45 As String * 1
    AvFF46 As String * 1
    AvFF47 As String * 1
    AvFF48 As String * 1
    AvFF49 As String * 1
    AvFF50 As String * 1
    AvFF51 As String * 1

    AvDescription As String * 64
        
    AvSepara As String * 448
    
End Type

Dim FileFree As Integer
Dim FileTemp As Integer
Dim RegActual As Long
Dim RegUltimo As Long
Dim RegActualTemp As Long
Dim Pos As Integer, p As Integer
Dim Datos As StructAvatars
Dim DatosTemp As StructAvatars

Private Sub cmdAnterior_Click()
If RegActual = 1 Then
    MsgBox " Primer registro ", vbInformation
Else
    'Diminuir a variável que mantém a posição do registro atual
    RegActual = RegActual - 1
    'Mostramos os dados nas caixas de texto
    VisualizarDatos
End If
End Sub


Private Sub cmdGuardar_Click()
    GuardarDatos
End Sub
Private Sub GuardarDatos()

'Atribuir estrutura de dados com o conteúdo do textBox
With Datos

    .AvNoMenu = txtToSave(txtAvNoMenu.Text)
    
    .AvNoImg = txtToSave(txtAvNoImg.Text)
    .AVEx2 = txtToSave(Tipoex.Text)
    .AVEx20 = HexToString("01")
    .AVEx21 = NullByte("0", 1)
    
    
    If chkAvNew.Value = 1 Then
        .AvNew = HexToString("00")
    Else
        .AvNew = HexToString("0A")
    End If
    
    .AvNew0 = NullByte("0", 7)
    
    .AvName = txtAvName + NullByte("0", Len(.AvName) - Len(txtAvName))
    
    .AvVisible0 = NullByte("0", 3)
    
    If chkAvVisible.Value = 1 Then
        .AvVisible = HexToString("01")
    Else
        .AvVisible = HexToString("00")
    End If
    
    .AvUnknow = NullByte("0", 1)
    
    If chkAvActWeek.Value = 1 Then
        .AvActWeek = HexToString("01")
    Else
        .AvActWeek = HexToString("00")
    End If
    
    .AvActWeek0 = NullByte("0", 3)
        
    .AvPriceWeekG = txtToSave(txtAvPriceWeekG.Text)
    .AvPriceWeekC = txtToSave(txtAvPriceWeekC.Text)
    
    .AvActN1 = NullByte("0", 1)
    .AvActN10 = NullByte("0", 3)
    .AvPriceN1G = NullByte("0", 4)
    .AvPriceN1C = NullByte("0", 4)
    
    If chkAvActMonth.Value = 1 Then
        .AvActMonth = HexToString("01")
    Else
        .AvActMonth = HexToString("00")
    End If

    .AvActMonth0 = NullByte("0", 3)
    .AvPriceMonthG = txtToSave(txtAvPriceMonthG.Text)
    .AvPriceMonthC = txtToSave(txtAvPriceMonthC.Text)
    
    .AvActN2 = NullByte("0", 1)
    .AvActN20 = NullByte("0", 3)
    .AvPriceN2G = NullByte("0", 4)
    .AvPriceN2C = NullByte("0", 4)
    
    If chkAvActIlimit.Value = 1 Then
        .AvActIlimit = HexToString("01")
    Else
        .AvActIlimit = HexToString("00")
    End If
    
    .AvActIlimit0 = NullByte("0", 3)
    .AvPriceIlimitG = txtToSave(txtAvPriceIlimitG.Text)
    .AvPriceIlimitC = txtToSave(txtAvPriceIlimitC.Text)
    
    If chkAvActGold.Value = 1 Then
        .AvActGold = HexToString("01")
    Else
        .AvActGold = HexToString("00")
    End If
    
    If chkAvActCash.Value = 1 Then
        .AvActCash = HexToString("01")
    Else
        .AvActCash = HexToString("00")
    End If
    
    .AvGoldCash0 = NullByte("0", 2)
    
    .Avvacio = NullByte("0", 12)
    .AvFF = HexToString("FF")
    .AvFF1 = HexToString("FF")
    .AvFF2 = HexToString("FF")
    .AvFF3 = HexToString("FF")
    .AvFF4 = HexToString("FF")
    .AvFF5 = HexToString("FF")
    .AvFF6 = HexToString("FF")
    .AvFF7 = HexToString("FF")
    .AvFF8 = HexToString("FF")
    .AvFF9 = HexToString("FF")
    .AvFF10 = HexToString("FF")
    .AvFF11 = HexToString("FF")
    .AvFF12 = HexToString("FF")
    .AvFF13 = HexToString("FF")
    .AvFF14 = HexToString("FF")
    .AvFF15 = HexToString("FF")
    .AvFF16 = HexToString("FF")
    .AvFF17 = HexToString("FF")
    .AvFF18 = HexToString("FF")
    .AvFF19 = HexToString("FF")
    .AvFF20 = HexToString("FF")
    .AvFF21 = HexToString("FF")
    .AvFF22 = HexToString("FF")
    .AvFF23 = HexToString("FF")
    .AvFF24 = HexToString("FF")
    .AvFF25 = HexToString("FF")
    .AvFF26 = HexToString("FF")
    .AvFF27 = HexToString("FF")
    .AvFF28 = HexToString("FF")
    .AvFF29 = HexToString("FF")
    .AvFF30 = HexToString("FF")
    .AvFF31 = HexToString("FF")
    .AvFF32 = HexToString("FF")
    .AvFF33 = HexToString("FF")
    .AvFF34 = HexToString("FF")
    .AvFF35 = HexToString("FF")
    .AvFF36 = HexToString("FF")
    .AvFF37 = HexToString("FF")
    .AvFF38 = HexToString("FF")
    .AvFF39 = HexToString("FF")
    .AvFF40 = HexToString("FF")
    .AvFF41 = HexToString("FF")
    .AvFF42 = HexToString("FF")
    .AvFF43 = HexToString("FF")
    .AvFF44 = HexToString("FF")
    .AvFF45 = HexToString("FF")
    .AvFF46 = HexToString("FF")
    .AvFF47 = HexToString("FF")
    .AvFF48 = HexToString("FF")
    .AvFF49 = HexToString("FF")
    .AvFF50 = HexToString("FF")
    .AvFF51 = HexToString("FF")
    
    .AvDescription = txtAvDescription.Text + NullByte("0", Len(.AvDescription) - Len(txtAvDescription))
        
    .AvSepara = NullByte("0", 448)
 
End With

'Escreve dados em um arquivo e posição
Put #FileFree, (RegActual - 1) * Len(Datos) + 5, Datos
End Sub

Private Sub Label8_Click()

End Sub

Private Sub Ex2adorno_Click()
Tipoex.Text = 1
End Sub

Private Sub Ex2anillo_Click()
Tipoex.Text = 3
End Sub

Private Sub Ex2collar_Click()
Tipoex.Text = 4
End Sub

Private Sub Ex2Especial_Click()
Tipoex.Text = 0
End Sub

Private Sub Ex2fondo_Click()
Tipoex.Text = 2
End Sub

Private Sub mnuOpen_Click()

    Cargar
    
End Sub


Private Sub mnuExit_Click()
    End
End Sub
Private Sub Cargar()
    dlgAbrir.DialogTitle = "Abrir"
    dlgAbrir.Filter = "Avatars Dat (*.dat)|*.dat"
    
    dlgAbrir.ShowOpen
    
    txtFileName.Text = dlgAbrir.FileName
    txtFileTitle.Text = dlgAbrir.FileTitle
    Form1.Caption = "Editor Ex2 Retzu - File " + txtFileTitle.Text
    
    If Not txtFileName.Text = "" And Not txtFileTitle.Text = "" Then
    
        FileFree = FreeFile
        Open dlgAbrir.FileName For Binary As FileFree Len = Len(Datos)
        RegActual = 1
        ' Armazenar a posição do último registro
        RegUltimo = LOF(FileFree) / Len(Datos)
        txtContador = RegUltimo
        If RegUltimo = 0 Then
            RegUltimo = 1
        End If
        
        VisualizarDatos
    End If
    
End Sub

Private Sub VisualizarDatos()
    Get #FileFree, (RegActual - 1) * Len(Datos) + 5, Datos
    
'    With Datos0
        'txtContador = Val("&H" + StringToHex(Mid$(.AvContador, 4, 1)) + StringToHex(Mid$(.AvContador, 3, 1)) + StringToHex(Mid$(.AvContador, 2, 1)) + StringToHex(Mid$(.AvContador, 1, 1)))
        'Datos = .AvPart2
    'End With
    
    With Datos
        txtAvNoMenu = Val("&H" + StringToHex(Mid$(.AvNoMenu, 4, 1)) + StringToHex(Mid$(.AvNoMenu, 3, 1)) + StringToHex(Mid$(.AvNoMenu, 2, 1)) + StringToHex(Mid$(.AvNoMenu, 1, 1)) + "&")
        Tipoex = StringToHex(.AVEx2)
        txtAvNoImg = Val("&H" + StringToHex(Mid$(.AvNoImg, 4, 1)) + StringToHex(Mid$(.AvNoImg, 3, 1)) + StringToHex(Mid$(.AvNoImg, 2, 1)) + StringToHex(Mid$(.AvNoImg, 1, 1)) + "&")
        txtAvNew = StringToHex(.AvNew)
        txtAvName = Trim(.AvName)
    
        txtAvVisible = StringToHex(.AvVisible)
        
        txtAvActWeek = Val("&H" + StringToHex(Mid$(.AvActWeek, 1, 1)))
        txtAvPriceWeekG = Val("&H" + StringToHex(Mid$(.AvPriceWeekG, 4, 1)) + StringToHex(Mid$(.AvPriceWeekG, 3, 1)) + StringToHex(Mid$(.AvPriceWeekG, 2, 1)) + StringToHex(Mid$(.AvPriceWeekG, 1, 1)) + "&")
        txtAvPriceWeekC = Val("&H" + StringToHex(Mid$(.AvPriceWeekC, 4, 1)) + StringToHex(Mid$(.AvPriceWeekC, 3, 1)) + StringToHex(Mid$(.AvPriceWeekC, 2, 1)) + StringToHex(Mid$(.AvPriceWeekC, 1, 1)) + "&")
        
        txtAvActMonth = Val("&H" + StringToHex(Mid$(.AvActMonth, 1, 1)))
        txtAvPriceMonthG = Val("&H" + StringToHex(Mid$(.AvPriceMonthG, 4, 1)) + StringToHex(Mid$(.AvPriceMonthG, 3, 1)) + StringToHex(Mid$(.AvPriceMonthG, 2, 1)) + StringToHex(Mid$(.AvPriceMonthG, 1, 1)) + "&")
        txtAvPriceMonthC = Val("&H" + StringToHex(Mid$(.AvPriceMonthC, 4, 1)) + StringToHex(Mid$(.AvPriceMonthC, 3, 1)) + StringToHex(Mid$(.AvPriceMonthC, 2, 1)) + StringToHex(Mid$(.AvPriceMonthC, 1, 1)) + "&")
        
        txtAvActIlimit = Val("&H" + StringToHex(Mid$(.AvActIlimit, 1, 1)))
        txtAvPriceIlimitG = Val("&H" + StringToHex(Mid$(.AvPriceIlimitG, 4, 1)) + StringToHex(Mid$(.AvPriceIlimitG, 3, 1)) + StringToHex(Mid$(.AvPriceIlimitG, 2, 1)) + StringToHex(Mid$(.AvPriceIlimitG, 1, 1)) + "&")
        txtAvPriceIlimitC = Val("&H" + StringToHex(Mid$(.AvPriceIlimitC, 4, 1)) + StringToHex(Mid$(.AvPriceIlimitC, 3, 1)) + StringToHex(Mid$(.AvPriceIlimitC, 2, 1)) + StringToHex(Mid$(.AvPriceIlimitC, 1, 1)) + "&")
        
        txtAvActGold = Val("&H" + StringToHex(Mid$(.AvActGold, 1, 1)))
        txtAvActCash = Val("&H" + StringToHex(Mid$(.AvActCash, 1, 1)))
        

        
        txtAvDescription = Trim(.AvDescription)
        
    End With
    
    'MsgBox (txtAvNew)
    If Val("&H" + Tipoex.Text) = 0 Then
        Ex2Especial.Value = 1
    End If
    If Val("&H" + Tipoex.Text) = 1 Then
        Ex2adorno.Value = 1
    End If
    If Val("&H" + Tipoex.Text) = 2 Then
        Ex2fondo.Value = 1
    End If
    If Val("&H" + Tipoex.Text) = 3 Then
        Ex2anillo.Value = 1
    End If
    If Val("&H" + Tipoex.Text) = 4 Then
        Ex2collar.Value = 1
    End If
    If Val("&H" + txtAvNew.Text) = 0 Then
        chkAvNew.Value = 1
    Else
        chkAvNew.Value = 0
    End If
    
    If Val("&H" + txtAvVisible) = 1 Then
        chkAvVisible.Value = 1
    Else
        chkAvVisible.Value = 0
    End If
    
    If Val("&H" + txtAvActWeek) = 1 Then
        chkAvActWeek.Value = 1
    Else
        chkAvActWeek.Value = 0
    End If
    
    If Val("&H" + txtAvActMonth) = 1 Then
        chkAvActMonth.Value = 1
    Else
        chkAvActMonth.Value = 0
    End If
    
    If Val("&H" + txtAvActIlimit) = 1 Then
        chkAvActIlimit.Value = 1
    Else
        chkAvActIlimit.Value = 0
    End If
    
    If Val("&H" + txtAvActGold) = 1 Then
        chkAvActGold.Value = 1
    Else
        chkAvActGold.Value = 0
    End If
    
    If Val("&H" + txtAvActCash) = 1 Then
        chkAvActCash.Value = 1
    Else
        chkAvActCash.Value = 0
    End If
        
    Combo1 = Combo1.List(0)
    mnuSave.Enabled = True
    
End Sub

Private Sub cmdSiguiente_click()

If RegActual = RegUltimo Then
    MsgBox " Ultimo registro ", vbInformation
Else
'Aumenta a posição
RegActual = RegActual + 1
'Coloque os dados na caixa de texto próximo registro
VisualizarDatos
End If

End Sub

Private Sub cmdBuscar_click()

Dim Encontrado As Boolean, PosReg As Long, tmp As StructAvatars

If txtBuscar = "" Then txtAvName.SetFocus: Exit Sub

Encontrado = False

'Vamos do começo ao fim em busca do registro para encontrar

For PosReg = 1 To RegUltimo

'Nós lemos o registro
Get #FileFree, (PosReg - 1) * Len(tmp) + 5, tmp

'Se os dados é o mesmo ciclo que
txtBuscar2 = BuscarPor(tmp)

If UCase(txtBuscar) = UCase(txtBuscar2) Then
    Encontrado = True
    Exit For
End If

Next

If Encontrado Then
    
    RegActual = PosReg
    'Coloque os dados do texto
    VisualizarDatos

Else
    MsgBox "Nome: " & txtBuscar & " Nenhum registro encontrado"
End If

End Sub

Private Function BuscarPor(t As StructAvatars)

Select Case Combo1.ListIndex

Case 0: BuscarPor = Trim(t.AvName)
Case 1: BuscarPor = Val("&H" + StringToHex(Mid$(t.AvNoMenu, 4, 1)) + StringToHex(Mid$(t.AvNoMenu, 3, 1)) + StringToHex(Mid$(t.AvNoMenu, 2, 1)) + StringToHex(Mid$(t.AvNoMenu, 1, 1)) + "&")

End Select

End Function

Private Sub CmdNuevo_click()

'Limpeza estrutura de dados para adicionar um novo registro
With Datos
    .AvNoMenu = ""
    .AVEx2 = ""
    .AVEx20 = ""
    .AVEx21 = ""
    .AvNoImg = ""
    
    .AvNew = ""
    .AvNew0 = ""
    
    .AvName = ""
    
    .AvVisible0 = ""
    .AvVisible = ""
    
    .AvUnknow = ""
    
    .AvActWeek = ""
    .AvActWeek0 = ""
    .AvPriceWeekG = ""
    .AvPriceWeekC = ""
    
    .AvActN1 = ""
    .AvActN10 = ""
    .AvPriceN1G = ""
    .AvPriceN1C = ""
    
    .AvActMonth = ""
    .AvActMonth0 = ""
    .AvPriceMonthG = ""
    .AvPriceMonthC = ""
    
    .AvActN2 = ""
    .AvActN20 = ""
    .AvPriceN2G = ""
    .AvPriceN2C = ""
    
    .AvActIlimit = ""
    .AvActIlimit0 = ""
    .AvPriceIlimitG = ""
    .AvPriceIlimitC = ""
    
    .AvActGold = ""
    .AvActCash = ""
    .AvGoldCash0 = ""
    
    .Avvacio = ""
    .AvFF = ""
    .AvDescription = ""
        
    .AvSepara = ""
 
End With

' Grava dados no novo registro até que você pressione o botão _
Salvar que registra os dados reais
'MsgBox (RegUltimo & "----" & Len(Datos))
Put #FileFree, (RegUltimo) * Len(Datos) + 5, Datos

RegActual = RegUltimo

VisualizarDatos
End Sub


Private Function HexToString(ByVal HexToStr As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(HexToStr) Step 3
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
        strReturn = strReturn & strTemp
    Next I
    HexToString = strReturn
End Function

Private Function StringToHex(ByVal StrToHex As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp
    Next I
    StringToHex = strReturn
End Function

Private Function NullByte(ByVal StrToNull As String, ByVal Contador As Integer) As String
Dim strReturn As String
Dim I As Long

For I = 1 To Contador
    strReturn = strReturn + HexToString("0")
Next I

    NullByte = strReturn

End Function

Private Function NullHex(ByVal StrToHex As String, ByVal Contador As Integer) As String
Dim strReturn As String
Dim I As Long
strReturn = StrToHex
For I = 1 To Contador
    strReturn = "0" + strReturn
Next I

    NullHex = strReturn

End Function

Private Function StatsView(ByVal StrToStat As String) As String
    If StrToStat > 50 Then
        StatsView = StrToStat - 255
    Else
        StatsView = StrToStat
    End If
End Function

Private Function StatsSave(ByVal StrToStat As String) As String
    If StrToStat < 0 Then
        StatsSave = StrToStat + 255
    Else
        StatsSave = StrToStat
    End If
End Function

Private Function txtToSave(ByVal StrToHex As String) As String
Dim t As StructAvatars

    StrToHex = NullHex(Hex(StrToHex), Len(t.AvNoImg) * 2 - Len(Hex(StrToHex)))
    StrToHex = HexToString(Mid$(StrToHex, 7, 2)) + HexToString(Mid$(StrToHex, 5, 2)) + HexToString(Mid$(StrToHex, 3, 2)) + HexToString(Mid$(StrToHex, 1, 2))
    txtToSave = StrToHex + NullByte("0", Len(t.AvNoImg) - Len(StrToHex))
    
End Function


