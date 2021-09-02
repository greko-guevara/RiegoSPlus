VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmriegosurcos 
   Caption         =   "Riego por Surcos "
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmriegosurcos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   1973
      TabIndex        =   25
      Top             =   3240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parámetros Generales del Riego"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Espaciamiento entre Surcos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Longitud de Surco"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Caudales y Tiempos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame11 
      Height          =   1095
      Left            =   2813
      TabIndex        =   16
      Top             =   6480
      Width           =   6255
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2280
         Picture         =   "frmriegosurcos.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bS 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   4320
         Picture         =   "frmriegosurcos.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   240
         Picture         =   "frmriegosurcos.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para el diseño"
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   10095
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtS 
         Height          =   285
         Left            =   8400
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtcultivo 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtWC 
         Height          =   285
         Left            =   6960
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLN 
         Height          =   285
         Left            =   5040
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtETR 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtIb 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "Textura"
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label57 
         Caption         =   "mm/día"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label56 
         Caption         =   "mm"
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label55 
         Caption         =   "m"
         Height          =   255
         Left            =   8280
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label50 
         Caption         =   "%"
         Height          =   255
         Left            =   9720
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "cm/hr"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Pendiente "
         Height          =   255
         Left            =   7200
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Espaciamiento entre las plantas "
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Lámina Neta"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Evapotranspiración Real"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Infiltración base "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cultivo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7815
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   18521
            MinWidth        =   18521
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "15/08/2007"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FpARAMETROS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   1613
      TabIndex        =   26
      Top             =   3720
      Width           =   8655
      Begin VB.TextBox txtLN1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtLB1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   31
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtFR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6480
         TabIndex        =   30
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtLB 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   720
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtEFAP 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton BEFAP 
         Caption         =   "Evaluar"
         Height          =   615
         Left            =   6360
         Picture         =   "frmriegosurcos.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "mm"
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "Lámina Neta corregida"
         Height          =   255
         Left            =   3240
         TabIndex        =   42
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "mm"
         Height          =   255
         Left            =   5160
         TabIndex        =   41
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "Lámina Bruta corregida"
         Height          =   255
         Left            =   3360
         TabIndex        =   40
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "Frecuencia de Riego"
         Height          =   255
         Left            =   6000
         TabIndex        =   39
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "día"
         Height          =   255
         Left            =   8040
         TabIndex        =   38
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Lámina Bruta"
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   -240
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "mm"
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Eficiencia de Aplicación"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "%"
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label58 
         Caption         =   "Lámina Bruta"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame fcaudal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   1613
      TabIndex        =   61
      Top             =   3720
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton BQM 
         Caption         =   "Caudales y Tiempos"
         Height          =   615
         Left            =   5520
         Picture         =   "frmriegosurcos.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Frame Frame10 
         Caption         =   "Caudal y Tiempo de Infiltración"
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   4440
         TabIndex        =   75
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtTI 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1920
            TabIndex        =   78
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtQI 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1920
            TabIndex        =   77
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtIP 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1920
            TabIndex        =   76
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label54 
            Caption         =   "Tiempo de Infiltración"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label53 
            Caption         =   "horas"
            Height          =   255
            Left            =   3240
            TabIndex        =   83
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label txtIpdfgdfgdfg 
            Caption         =   "Infiltración promedio"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label49 
            Caption         =   "Caudal de Infiltración"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label48 
            Caption         =   "lts/min"
            Height          =   255
            Left            =   3240
            TabIndex        =   80
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label47 
            Caption         =   "cm/hr"
            Height          =   255
            Left            =   3240
            TabIndex        =   79
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Caudal y tiempo de avance"
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   240
         TabIndex        =   62
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtTA 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1920
            TabIndex        =   66
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtQMP 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1920
            TabIndex        =   65
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtQMC 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1920
            TabIndex        =   64
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtQMG 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1920
            TabIndex        =   63
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label52 
            Caption         =   "Tiempo de Avance"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label51 
            Caption         =   "horas"
            Height          =   255
            Left            =   3240
            TabIndex        =   73
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label44 
            Caption         =   "lts/min"
            Height          =   255
            Left            =   3240
            TabIndex        =   72
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label43 
            Caption         =   "Caudal de Avance"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label42 
            Caption         =   "lts/min"
            Height          =   255
            Left            =   3240
            TabIndex        =   70
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "lts/min"
            Height          =   255
            Left            =   3240
            TabIndex        =   69
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label40 
            Caption         =   "Qmax según Gardner"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label39 
            Caption         =   "Qmax según Criddle"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fLongitud 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   1613
      TabIndex        =   55
      Top             =   3720
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox txtLM 
         BackColor       =   &H80000016&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   720
         TabIndex        =   58
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Blongitudmaxima 
         Caption         =   "Longitud Max."
         Height          =   615
         Left            =   480
         MaskColor       =   &H00008000&
         Picture         =   "frmriegosurcos.frx":315C
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtL 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6480
         TabIndex        =   56
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label31 
         Caption         =   "Longitud a que se Construiran los Surcos"
         Height          =   255
         Left            =   4920
         TabIndex        =   90
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label30 
         Caption         =   $"frmriegosurcos.frx":38C6
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   3120
         TabIndex        =   89
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label25 
         Caption         =   "m"
         Height          =   255
         Left            =   8040
         TabIndex        =   88
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label32 
         Caption         =   "m"
         Height          =   255
         Left            =   2160
         TabIndex        =   60
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "Longitud  Recomenda"
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   480
         Width           =   2415
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   3720
         X2              =   7080
         Y1              =   2160
         Y2              =   2160
      End
   End
   Begin VB.Frame Fespaciamiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   1613
      TabIndex        =   44
      Top             =   3720
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox txtW 
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4080
         TabIndex        =   49
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtW1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6480
         TabIndex        =   47
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton BCW 
         Caption         =   "Calcular W"
         Height          =   615
         Left            =   6360
         Picture         =   "frmriegosurcos.frx":3954
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3600
         TabIndex        =   46
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Left            =   720
         TabIndex        =   45
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "m"
         Height          =   255
         Left            =   8040
         TabIndex        =   87
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "m"
         Height          =   255
         Left            =   2160
         TabIndex        =   86
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "W"
         Height          =   255
         Left            =   4680
         TabIndex        =   54
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "Al comparar el espaciamiento entre plantas y el valor de W según Hozapfel se decide un W en metros de : "
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   840
         TabIndex        =   53
         Top             =   1680
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Line Line4 
         X1              =   720
         X2              =   5280
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label13 
         Caption         =   "W recomendado Hozapfel "
         Height          =   375
         Left            =   5760
         TabIndex        =   52
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "KS"
         Height          =   255
         Left            =   3240
         TabIndex        =   51
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Profundidad de Raices "
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   1680
      Picture         =   "frmriegosurcos.frx":40BE
      Top             =   2760
      Width           =   8505
   End
   Begin VB.Label Label10 
      Caption         =   "Diseño de Riego por Surcos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   3855
   End
   Begin VB.Menu mpsc 
      Caption         =   "Parámetros Suelo- Clima"
      Begin VB.Menu mps 
         Caption         =   "Parámetros Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mch 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meva 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu motrscal 
      Caption         =   "Otros Cálculos en Surcos"
      Begin VB.Menu mprueba 
         Caption         =   "Prueba de Avance"
      End
      Begin VB.Menu msurinf 
         Caption         =   "Surcos Infiltrómetros"
      End
   End
   Begin VB.Menu masimat 
      Caption         =   "Asistente Matemático"
      Begin VB.Menu mconcer 
         Caption         =   "Convertidor de Unidades"
      End
      Begin VB.Menu h 
         Caption         =   "Hidráulica de Canales"
      End
      Begin VB.Menu mregpotsim 
         Caption         =   "Regresión Potencial Simple"
      End
   End
   Begin VB.Menu mmpr 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmriegosurcos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub mch_Click()
frmconductividad.Show
End Sub

Private Sub mconcer_Click()
frmconvertidor.Show
End Sub

Private Sub meva_Click()
frmETO.Show
End Sub

Private Sub mmpr_Click()
frmGeneral.Show
Unload Me
End Sub

Private Sub mprueba_Click()
frmpruebaavance.Show

End Sub

Private Sub mps_Click()
frmgenerales.Show
End Sub

Private Sub mregpotsim_Click()
frmpruebaavance.Show
With frmpruebaavance
    .Caption = "Regresión Potencial Simple"
    .lbltitulo.Caption = "Regresión Potencial Simple"
    .grdDatos.TextMatrix(0, 0) = "X"
    .grdDatos.TextMatrix(0, 1) = "Y"
    .Label1 = "Y="
    .Label5 = "X"
    
End With

End Sub

Private Sub msurinf_Click()
frmsurcosinfiltrometros.Show
End Sub

Private Sub mt_Click()
frmtextura.Show
End Sub

Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    
    FpARAMETROS.Visible = True
    fLongitud.Visible = False
    Fespaciamiento.Visible = False
    fcaudal.Visible = False
  
    StatusBar1.Panels(1).text = "Digite los datos generales del proyecto y oprima el Botón de Evaluar para determinar los parámetros generales del Riego"
    Case 2
    FpARAMETROS.Visible = False
    fLongitud.Visible = False
    Fespaciamiento.Visible = True
    fcaudal.Visible = False
   
    StatusBar1.Panels(1).text = "Realice el cálculo de espaciamiento entre surcos recomendado y compárelo con el espaciamiento entre plantas"
    
    Case 3
    FpARAMETROS.Visible = False
    fLongitud.Visible = True
    Fespaciamiento.Visible = False
    fcaudal.Visible = False
    
    StatusBar1.Panels(1).text = "Oprima el Botón de Longitud Max. Recomendado y al valor recomenddo compárelo con las cualidades de la finca "
    Case 4
    FpARAMETROS.Visible = False
    fLongitud.Visible = False
    Fespaciamiento.Visible = False
    fcaudal.Visible = True
    
    StatusBar1.Panels(1).text = "Oprima el Botón de Caudales y Tiempos para determinar estas variantes"
    Case 5
    FpARAMETROS.Visible = False
    fLongitud.Visible = False
    Fespaciamiento.Visible = False
    fcaudal.Visible = False
    freporte.Visible = True
    
End Select
End Sub

Private Sub BCW_Click()
On Error GoTo mensaje
    wc = Val(txtWC.text)
    d = Val(txtD.text)
    ks = Combo2.text
    If wc = 0 Then
    MsgBox "Ingrese el valor de espacimiento entre plantas", 64, "Riego por Surcos"
    txtWC.SetFocus
    Exit Sub
    End If
    If d = 0 Then
    MsgBox "Ingrese el valor de profundidad de raices", 64, "Riego por Surcos"
    txtD.SetFocus
    Exit Sub
    End If
    If ks = 0 Then
    MsgBox "seleccione la constante de Ks", 64, "Riego por Surcos"
    Combo2.SetFocus
    Exit Sub
    End If
    
    Select Case ks
         Case "Arcilloso 2.5"
            ks = 2.5
         Case "Medio 1.5"
            ks = 1.5
         Case "Arenoso 0.5"
            ks = 0.5
    End Select
    w1 = d * ks
    txtW1.text = Format(w1, "##0.0#")
    Label14.Visible = True
    Label15.Visible = True
    txtW.Visible = True
    If w1 < 2 * wc Then
    txtW.text = wc
    Else
    txtW.text = 2 * wc
    End If
    txtW.SetFocus
Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Riego por Surcos"


End Sub

Private Sub BEFAP_Click()
 On Error GoTo mensaje
    s = Val(txtS.text)
    ln = Val(txtLn.text)
    etr = Val(txtetr.text)
    If ln = 0 Then
    MsgBox "Ingrese el valor de lámina neta", 64, "Riego por Surcos"
    txtLn.SetFocus
    Exit Sub
    End If
    If etr = 0 Then
    MsgBox "Ingrese el valor de evapotranspiración", 64, "Riego por Surcos"
    txtetr.SetFocus
    Exit Sub
    End If
    If s = 0 Then
    MsgBox "Ingrese el valor de la pendiente", 64, "Riego por Surcos"
    txtS.SetFocus
    Exit Sub
    End If
    
    Rem Eficiencia en la aplicacion
    If s <= 0.1 Then
        txtEFAP.text = 80
    Else
        If s <= 0.5 Then
            txtEFAP.text = 70
        Else
            If s <= 1 Then
                txtEFAP.text = 65
            Else
                If s <= 2 Then
                    txtEFAP.text = 55
                End If
            End If
        End If
    End If
        Rem lamina bruta
    efap = Val(txtEFAP.text)
    lb = ln / efap * 100
    
    Rem de Frecuencia de riego
    fr = ln / etr
    
    If fr <= 1 Then
        fr = 1
    End If
    
    Rem*********************
    txtLB.text = Format(lb, "#0.0#")
    txtFR.text = Format(Int(fr), "#0.0#")
    Rem laminas corregidas
    fr1 = Val(txtFR.text)
    ln1 = fr1 * etr
    lb1 = ln1 / efap * 100
    
    txtLN1.text = Format(ln1, "#0.0#")
    txtLB1.text = Format(lb1, "#0.0#")
Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Riego por Surcos"
End Sub



Private Sub Blongitudmaxima_Click()
On Error GoTo mensaje
    lb = Val(txtLB.text)
    s = Val(txtS.text)
    tex = Combo3.text
    If Combo3 = "" Then
    MsgBox "Seleccione la textura", 64, "Riego por Surcos"
    Combo3.SetFocus
    Exit Sub
    End If
    If s = 0 Then
    MsgBox "Ingrese la pendiente", 64, "Riego por Surcos"
    txtS.SetFocus
    Exit Sub
    End If
    If lb = 0 Then
    MsgBox "Ingrese el valor de la lámina bruta", 64, "Riego por Surcos"
    txtLB.SetFocus
    Exit Sub
    End If
    
    Select Case tex
        Case "muy fina"
            lm = 63.4204357 * s ^ -0.5516248 * (lb / 10) ^ 0.5208104
        Case "fina"
            lm = 56.6316209 * s ^ -0.5565067 * (lb / 10) ^ 0.5168311
        Case "media"
            lm = 69.033016 * s ^ -0.5582466 * (lb / 10) ^ 0.3605115
        Case "gruesa"
            lm = 38.6620717 * s ^ -0.5627258 * (lb / 10) ^ 0.5265393
        Case "muy gruesa"
            lm = 27.6075443 * s ^ -0.5618593 * (lb / 10) ^ 0.5551165
    End Select
    txtLM.text = Format(lm, "###0")
    Label25.Visible = True
    Label30.Visible = True
    Label31.Visible = True
    txtL.Visible = True
    Line1.Visible = True
    txtL.SetFocus
Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Riego por Surcos"

End Sub

Private Sub BQM_Click()
On Error GoTo mensaje
    s = Val(txtS.text)
    g = Combo3.text
    f = Combo2.text
    ib = Val(txtIb.text)
    w = Val(txtW.text)
    L = Val(txtL.text)
    lb1 = Val(txtLB1.text)
    If s = 0 Then
    MsgBox "Ingrese la pendiente", 64, "Riego por Surcos"
    txtS.SetFocus
    Exit Sub
    End If
    If ib = 0 Then
    MsgBox "Ingrese el valor de infiltración básica", 64, "Riego por Surcos"
    txtIb.SetFocus
    Exit Sub
    End If
    If L = 0 Then
    MsgBox "Ingrese el valor de longitud", 64, "Riego por Surcos"
    txtS.SetFocus
    Exit Sub
    End If
    If lb1 = 0 Then
    MsgBox "Ingrese el valor de lámina bruta corregida", 64, "Riego por Surcos"
    txtS.SetFocus
    Exit Sub
    End If
    If s = 0 Then
    MsgBox "Ingrese la pendiente", 64, "Riego por Surcos"
    txtS.SetFocus
    Exit Sub
    End If
    
    Select Case f
        Case "Arcilloso 2.5"
            f = 1.2
         Case "Medio 1.5"
            f = 1.33
         Case "Arenoso 0.5"
            f = 1.5
    End Select
    Select Case g
         Case "muy fina"
            c = 0.892
            a = 0.937
         Case "fina"
            c = 0.998
            a = 0.55
         Case "media"
            c = 0.613
            a = 0.733
         Case "gruesa"
            c = 0.644
            a = 0.704
         Case "muy gruesa"
            c = 0.665
            a = 0.548
    End Select
    Rem Caudales maximos
    qmc = 38 / s
    qmg = 60 * c / s ^ a
    qmp = (qmc + qmg) / 2
    
    txtQMC.text = Format(qmc, "##0.00#")
    txtQMG.text = Format(qmg, "##0.00#")
    txtQMP.text = Format(qmp, "##0.00#")
    Rem caudales de infiltracion
    ip = ib * f
    qi1 = ip * w * L * 0.167
    qi2 = qmp / 2
    qip = (qi1 + qi2) / 2
    txtIP.text = Format(ip, "##0.00#")
    txtQI.text = Format(qip, "##0.00#")
    
    Rem Timpos--------------------------------
    ti = lb1 / ip / 10
    ta = ti / 4
    txtTI.text = Format(ti, "##0.00#")
    txtTa.text = Format(ta, "##0.00#")
Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Riego por Surcos"
    
End Sub





Private Sub bS_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub Command1_Click()
txtcultivo.text = ""
txtWC.text = ""
txtIb.text = ""
txtetr.text = ""
txtLn.text = ""
txtS.text = ""

txtD.text = ""
txtW.text = ""
txtW1.text = ""
txtEFAP.text = ""
txtLB.text = ""
txtLB1.text = ""
txtLN1.text = ""
txtFR.text = ""

txtLM.text = ""


txtL.text = ""
txtQMC.text = ""
txtQMG.text = ""
txtQMP.text = ""
txtIP.text = ""
txtQI.text = ""
txtTI.text = ""
txtTa.text = ""

Label25.Visible = False
Label30.Visible = False
Label31.Visible = False
txtL.Visible = False
Line1.Visible = False
Label14.Visible = False
Label15.Visible = False
txtW.Visible = False
    

End Sub

Private Sub Command2_Click()
Print Form

End Sub

Private Sub Form_Load()
Rem ks
With Combo2
    .AddItem "Arcilloso 2.5"
    .AddItem "Medio 1.5"
    .AddItem "Arenoso 0.5"
End With
With Combo3
    .AddItem "muy fina"
    .AddItem "fina"
    .AddItem "media"
    .AddItem "gruesa"
    .AddItem "muy gruesa"
End With

 StatusBar1.Panels(1).text = "Digite los datos generales del proyecto y oprima el Botón de Evaluar para determinar los parámetros generales del Riego"
End Sub








