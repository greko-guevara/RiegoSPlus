VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmsurcosinfiltrometros 
   Caption         =   "Método de Surcos Infiltrómetros"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmsurcosinfiltrometros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame8 
      Height          =   3975
      Left            =   600
      TabIndex        =   70
      Top             =   3720
      Visible         =   0   'False
      Width           =   10215
      Begin GraphLib.Graph Graph1 
         Height          =   2775
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   96
         BorderStyle     =   1
         GraphType       =   6
         RandomData      =   1
         ColorData       =   0
         ExtraData       =   0
         ExtraData[]     =   0
         FontFamily      =   4
         FontSize        =   4
         FontSize[0]     =   200
         FontSize[1]     =   150
         FontSize[2]     =   100
         FontSize[3]     =   100
         FontStyle       =   4
         GraphData       =   0
         GraphData[]     =   0
         LabelText       =   0
         LegendText      =   0
         PatternData     =   0
         SymbolData      =   0
         XPosData        =   0
         XPosData[]      =   0
      End
      Begin GraphLib.Graph Graph2 
         Height          =   2775
         Left            =   3480
         TabIndex        =   72
         Top             =   600
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   96
         BorderStyle     =   1
         GraphType       =   6
         RandomData      =   1
         ColorData       =   0
         ExtraData       =   0
         ExtraData[]     =   0
         FontFamily      =   4
         FontSize        =   4
         FontSize[0]     =   200
         FontSize[1]     =   150
         FontSize[2]     =   100
         FontSize[3]     =   100
         FontStyle       =   4
         GraphData       =   0
         GraphData[]     =   0
         LabelText       =   0
         LegendText      =   0
         PatternData     =   0
         SymbolData      =   0
         XPosData        =   0
         XPosData[]      =   0
      End
      Begin GraphLib.Graph Graph3 
         Height          =   2775
         Left            =   6840
         TabIndex        =   73
         Top             =   600
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   96
         BorderStyle     =   1
         GraphType       =   6
         RandomData      =   1
         ColorData       =   0
         ExtraData       =   0
         ExtraData[]     =   0
         FontFamily      =   4
         FontSize        =   4
         FontSize[0]     =   200
         FontSize[1]     =   150
         FontSize[2]     =   100
         FontSize[3]     =   100
         FontStyle       =   4
         GraphData       =   0
         GraphData[]     =   0
         LabelText       =   0
         LegendText      =   0
         PatternData     =   0
         SymbolData      =   0
         XPosData        =   0
         XPosData[]      =   0
      End
      Begin VB.Label Label39 
         Caption         =   "Tiempo acumulado"
         Height          =   255
         Left            =   8520
         TabIndex        =   79
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Tiempo acumulado"
         Height          =   255
         Left            =   1800
         TabIndex        =   78
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label38 
         Caption         =   " Infiltración acumulada"
         Height          =   255
         Left            =   6960
         TabIndex        =   77
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label37 
         Caption         =   "Caudal infiltrado"
         Height          =   255
         Left            =   3600
         TabIndex        =   76
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "Caudal de salida "
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "Tiempo acumulado"
         Height          =   255
         Left            =   5160
         TabIndex        =   74
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Height          =   3135
      Left            =   9240
      TabIndex        =   17
      Top             =   360
      Width           =   1815
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   120
         Picture         =   "frmsurcosinfiltrometros.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   120
         Picture         =   "frmsurcosinfiltrometros.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   120
         Picture         =   "frmsurcosinfiltrometros.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Fijos"
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   600
      TabIndex        =   8
      Top             =   1080
      Width           =   8295
      Begin VB.OptionButton Option2 
         Caption         =   "lts/min"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "lts/seg"
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ecuación de avance"
         Height          =   615
         Left            =   4440
         TabIndex        =   21
         Top             =   1200
         Width           =   3255
         Begin VB.TextBox Text1 
            ForeColor       =   &H80000006&
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text2 
            ForeColor       =   &H80000006&
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label31 
            Caption         =   "B"
            Height          =   255
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label32 
            Caption         =   "T = A x L^B"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label33 
            Caption         =   "A"
            Height          =   255
            Left            =   1200
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox txtef 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtln 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Left            =   5880
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtW 
         Height          =   285
         Left            =   5880
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtQE 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblund 
         Caption         =   "l/min"
         Height          =   255
         Left            =   3600
         TabIndex        =   66
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label30 
         Caption         =   "Unidades de caudal"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Caudal de entrada"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "mm"
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "%"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Eficiencia"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Lámina neta"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "mts"
         Height          =   255
         Left            =   7440
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Longitud de Surco"
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "mts"
         Height          =   255
         Left            =   7440
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Ancho de Surco W"
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
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
            TextSave        =   "01/09/2006"
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
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   375
      Left            =   3720
      TabIndex        =   65
      Top             =   3360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Introduzca los datos de la Prueba"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resultados"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gráficos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdCrear 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.QsT)"
   End
   Begin MSComDlg.CommonDialog cdAccesar 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar el archivo a cargar"
      Filter          =   "(*.QsT)"
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   600
      TabIndex        =   30
      Top             =   3720
      Width           =   10215
      Begin VB.CommandButton bagregar 
         Caption         =   "Evaluar Datos"
         Height          =   615
         Left            =   360
         Picture         =   "frmsurcosinfiltrometros.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Es necesario introducir el valor del ancho y longitud del surco de prueba así como el caudal de entrada"
         Top             =   480
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdD 
         Height          =   3615
         Left            =   2040
         TabIndex        =   31
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   16761024
      End
      Begin VB.Label Label9 
         Caption         =   "Oprima  (*) para eliminar líneas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   855
         Left            =   240
         TabIndex        =   69
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Oprima Enter para insertar  líneas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1215
         Left            =   240
         TabIndex        =   68
         Top             =   1560
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3975
      Left            =   600
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Frame Frame3 
         Caption         =   "Ecuaciones potenciales"
         ForeColor       =   &H00800000&
         Height          =   3135
         Left            =   960
         TabIndex        =   50
         Top             =   600
         Width           =   3975
         Begin VB.CommandButton bcalcular 
            Caption         =   "&Ecuaciones"
            Height          =   615
            Left            =   1440
            Picture         =   "frmsurcosinfiltrometros.frx":29F2
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtaabb 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   240
            TabIndex        =   56
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtaa 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   480
            TabIndex        =   55
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtbb 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2280
            TabIndex        =   54
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txta 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   480
            TabIndex        =   53
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtb 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2280
            TabIndex        =   52
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtab 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   240
            TabIndex        =   51
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "li=cm/hr"
            Height          =   255
            Left            =   2880
            TabIndex        =   64
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "t=min"
            Height          =   255
            Left            =   2880
            TabIndex        =   63
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "Icum=mm"
            Height          =   255
            Left            =   2880
            TabIndex        =   62
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "a"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "b"
            Height          =   255
            Left            =   1920
            TabIndex        =   60
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "A"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   "B"
            Height          =   255
            Left            =   1920
            TabIndex        =   58
            Top             =   480
            Width           =   375
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   3480
            Y1              =   2160
            Y2              =   2160
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Resultados de la Prueba Surcos Infiltrómetros"
         ForeColor       =   &H00800000&
         Height          =   3135
         Left            =   5280
         TabIndex        =   33
         Top             =   600
         Width           =   3975
         Begin VB.CommandButton bcalculo 
            Caption         =   "Calcular"
            Height          =   615
            Left            =   1560
            Picture         =   "frmsurcosinfiltrometros.frx":315C
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txttava 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2040
            TabIndex        =   38
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txttinf 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2040
            TabIndex        =   37
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtlon 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2040
            TabIndex        =   36
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtqava 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2040
            TabIndex        =   35
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtqinf 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2040
            TabIndex        =   34
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Tiempo avance"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Tiempo infiltración"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "min"
            Height          =   255
            Left            =   3360
            TabIndex        =   47
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label23 
            Caption         =   "min"
            Height          =   255
            Left            =   3360
            TabIndex        =   46
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label29 
            Caption         =   "Longitud de Surco"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label28 
            Caption         =   "mts"
            Height          =   255
            Left            =   3360
            TabIndex        =   44
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label27 
            Caption         =   "Caudal avance"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Caudal infiltración"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label25 
            Caption         =   "l/min"
            Height          =   255
            Left            =   3360
            TabIndex        =   41
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label24 
            Caption         =   "l/min"
            Height          =   255
            Left            =   3360
            TabIndex        =   40
            Top             =   2040
            Width           =   375
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Método de Surcos Infiltrómetros"
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
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   4215
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mcrear 
         Caption         =   "Guardar como"
         Shortcut        =   ^G
      End
      Begin VB.Menu maccesar 
         Caption         =   "Abrir proyecto"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mpasuel 
      Caption         =   "Parámetros Suelo - Clima"
      Begin VB.Menu mpasuel1 
         Caption         =   "Parámetros Suelo"
      End
      Begin VB.Menu mtex 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcond 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu mevap 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu motros 
      Caption         =   "Otros Cálculos en Surcos"
      Begin VB.Menu masis 
         Caption         =   "Asistente de Diseño"
      End
      Begin VB.Menu mpa 
         Caption         =   "Prueba de Avance"
      End
   End
   Begin VB.Menu massi 
      Caption         =   "Asistente Matemático"
      Begin VB.Menu mconvr 
         Caption         =   "Convertidor de Unidades"
      End
      Begin VB.Menu h 
         Caption         =   "Hidráulica de Canales"
      End
      Begin VB.Menu mregpot 
         Caption         =   "Regresión Potencial Simple"
      End
   End
   Begin VB.Menu mmp 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmsurcosinfiltrometros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim valort(1 To 300) As Double
Dim valorta(1 To 300) As Double
Dim valorqs(1 To 300) As Double
Dim valorqi(1 To 300) As Double
Dim valori(1 To 300) As Double
Dim valorl(1 To 300) As Double
Dim valoria(1 To 300) As Single
Dim km(1 To 7) As Double
Dim u As Integer


Dim row, col, numero, num(0 To 1000)
Dim n As Integer, i, ii, punto

Private Sub bimprimir_Click()
Print Form
End Sub





Private Sub grdD_Click()
i = ""
punto = 0
End Sub

Private Sub grdD_KeyPress(KeyAscii As Integer)

If grdD.col <> col Or grdD.row <> row Then
i = ""
punto = 0
For s = 1 To numero
num(s) = ""
Next s
numero = 0
End If
num(0) = ""
If KeyAscii = 45 And i = "" Then
i = i + "-"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 48 Then
    i = i + "0"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If


If punto <> 1 Then
If KeyAscii = 44 Or KeyAscii = 46 Then
    numero = numero + 1
    If i = "" Then
    i = i + "0."
    grdD.text = i
    num(numero) = i
    punto = 1
Else
    i = i + "."
    grdD.text = i
    num(numero) = i
    punto = 1
End If
End If
End If


If KeyAscii = 49 Then
    i = i + "1"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 50 Then
    i = i + "2"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 51 Then
    i = i + "3"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 52 Then
    i = i + "4"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 53 Then
    i = i + "5"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 54 Then
    i = i + "6"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 55 Then
    i = i + "7"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 56 Then
    i = i + "8"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 57 Then
    i = i + "9"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If
Rem tecla de borrado
If numero >= 1 Then
If KeyAscii = 8 Then
i = num(numero - 1)
numero = numero - 1
grdD.text = i
End If
Else
grdD.text = ""
End If

'ingresar nueva linea
If KeyAscii = 13 Then
u = u + 1
grdD.Rows = u + 2
End If

If KeyAscii = 42 Then
If u >= 2 Then
    u = u - 1
    grdD.Rows = u + 2
End If
End If



Rem pruebas grid1.TextMatrix(numero, 6) = num(numero)

Rem grid1.Text = KeyAscii
col = grdD.col
row = grdD.row

End Sub





Private Sub bagregar_Click()
On Error GoTo mensaje
    
    n = u + 1
    qe = Val(txtQE.text)
    w = Val(txtW.text)
    L = Val(txtL.text)

If n < 2 Then
MsgBox "Ingrese al menos un par de datos de la prueba", 64, "Imposible Calcular"
grdD.SetFocus
Exit Sub
End If
If qe = 0 Then
MsgBox "Ingrese el valor del caudal de entrada", 64, "Surcos Infiltrómetros"
txtQE.SetFocus
Exit Sub
End If
If w = 0 Then
MsgBox "Ingrese el valor del ancho del surco", 64, "Surcos Infiltrómetros"
txtW.SetFocus
Exit Sub
End If
If L = 0 Then
MsgBox "Ingrese el valor de la longitud del surco", 64, "Surcos Infiltrómetros"
txtL.SetFocus
Exit Sub
End If

    'definición de parametros con los que se llena el grid
    st = 0
    sl = 0
    If Option1.Value = True Then
          qe = qe * 60
    End If
    
    For j% = 1 To n
        valort(j%) = Val(grdD.TextMatrix(j%, 0))
        valorqs(j%) = Val(grdD.TextMatrix(j%, 1))
        'deficion de unidades
        If Option1.Value = True Then
            valorqs(j%) = valorqs(j%) * 60
        End If
        valorqi(j%) = qe - valorqs(j%)
        valori(j%) = valorqi(j%) / (w * L) * 6
        valorl(j%) = valori(j%) * valort(j%) * 10 / 60
        st = st + valort(j%)
        sl = sl + valorl(j%)
        valorta(j%) = st
        valoria(j%) = sl
        With grdD
            .Cols = 7
            .TextMatrix(0, 2) = "Tacu min"
            .TextMatrix(0, 3) = "Qinf. l/min"
            .TextMatrix(0, 4) = "li cm/h"
            .TextMatrix(0, 5) = "L parc mm"
            .TextMatrix(0, 6) = "lacum mm"

            .TextMatrix(j%, 2) = Format(st, "##0.00#")
            .TextMatrix(j%, 3) = Format(valorqi(j%), "##0.00#")
            .TextMatrix(j%, 4) = Format(valori(j%), "##0.00#")
            .TextMatrix(j%, 5) = Format(valorl(j%), "##0.00#")
            .TextMatrix(j%, 6) = Format(sl, "##0.00#")
        End With
    Next j%
Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Riego por Surcos"

End Sub

Private Sub Bcalcular_Click()
If grdD.TextMatrix(1, 1) <> "" Then
On Error GoTo mensaje
Rem *************CALCULO DE ICUM***********************
    n = u + 1
    st = 0
    st2 = 0
    si = 0
    sit = 0
    ss1 = 0
    ssit = 0
    
    For j% = 1 To n
        st = st + Log(valorta(j%))
        st2 = st2 + Log(valorta(j%)) ^ 2
        si = si + Log(valoria(j%))
        sit = sit + Log(valorta(j%)) * Log(valoria(j%))
        ssi = ssi + Log(valori(j%))
        ssit = ssit + Log(valorta(j%)) * Log(valori(j%))
    Next j%
    b = ((n * sit - st * si) / (n * st2 - st ^ 2))
    a = (si / n - st / n * b)
    a1 = Exp(a)
    bb = ((n * ssit - st * ssi) / (n * st2 - st ^ 2))
    aa = (ssi / n - st / n * bb)
    aa1 = Exp(aa)
    
    
Rem ******************CALCULO DE LI***************************

    
    txtA.text = Format(a1, "###0.0###")
    txtb.text = Format(b, "###0.0###")
    txtaa.text = Format(aa1, "###0.0###")
    txtbb.text = Format(bb, "###0.0###")
    txtaabb.text = "li=" + Format(aa1, "###0.0###") + "*t^" + Format(bb, "###0.0###")
    txtab.text = "lacum=" + Format(a1, "###0.0###") + "*t^" + Format(b, "###0.0###")

Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Surcos Infiltrómetros"
Else
    MsgBox "Ingrese al menos un par de datos", 64, "Surcos Infiltrómetros"

End If

End Sub

Private Sub bcalculo_Click()
  On Error GoTo mensaje
    b = Val(txtb.text)
    a = Val(txtA.text)
    ln = Val(txtLn.text)
    ef = Val(txtEf.text)
    aaa = Val(Text1.text)
    bbb = Val(Text2.text)
    aa = Val(txtaa.text)
    bb = Val(txtbb.text)
    w = Val(txtW.text)
    qe = Val(txtQE.text)

    If Option1.Value = True Then
        qe = qe * 60
    Else
    End If
    
    lb = (ln / ef * 100)
    
    
    tinf = (lb / a) ^ (1 / b)
    tava = tinf / 4
    
    lon = (tava / aaa) ^ (1 / bbb)
    
   
    
    Qinf = aa * tava ^ bb * w * lon * 10 / 60
    
    
    txtqava.text = Format(qe, "###0.0##")
    txtqinf.text = Format(Qinf, "###0.0##")
    txtlon.text = Format(lon, "###0.0##")
    txttava.text = Format(tava, "###0.0#")
    txttinf.text = Format(tinf, "###0.0#")

Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Surcos Infiltrómetros"


End Sub


Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
txtQE.text = ""
txtW.text = ""
txtL.text = ""
txtLn.text = ""
txtEf.text = ""
Text1.text = ""
Text2.text = ""
txtaa.text = ""
txtA.text = ""
txtbb.text = ""
txtb.text = ""
txtab.text = ""
txtaabb.text = ""
txttava.text = ""
txttinf.text = ""
txtlon.text = ""
txtqava.text = ""
txtqinf.text = ""
Graph1.DrawMode = gphClear
Graph2.DrawMode = gphClear
Graph3.DrawMode = gphClear
grdD.Clear
With grdD
    .TextMatrix(0, 0) = "T Parcial min"
    .TextMatrix(0, 1) = "Qsal l/min"
    .Rows = 2
    .Cols = 2
End With
Option2.Value = True
lblund.Caption = "l/min"

u = 0


End Sub


Private Sub Form_Load()
u = 0
With grdD
    .ColAlignment(0) = 4
    .ColAlignment(1) = 4
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .TextMatrix(0, 0) = "T Parcial min"
    .TextMatrix(0, 1) = "Qsal. l/min"
End With
 StatusBar1.Panels(1).text = "Digite los datos generales del sistema de riego e introduzca los valores de la prueba de Surco Infiltrómetro"
End Sub

Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub maccesar_Click()
On Error GoTo SinArchivo:
cdAccesar.ShowOpen
 NombreArch = cdAccesar.FileName
 u = 0
 Open NombreArch For Random As #1 Len = Len(Paresqst1)
 NumReg = LOF(1) \ Len(Paresqst1)
 grdD.Rows = NumReg + 1
 For j% = 1 To NumReg
  Get #1, j%, Paresqst1
  txtQE.text = Val(Paresqst1.kkqe)
  txtEf.text = Val(Paresqst1.kkef)
  txtLn.text = Val(Paresqst1.kkln)
  txtW.text = Val(Paresqst1.kkw)
  txtL.text = Val(Paresqst1.kkl)
  Text1.text = Val(Paresqst1.kktext1)
  Text2.text = Val(Paresqst1.kktext2)
  LL = Paresqst1.qs
  tT = Paresqst1.tT
  
  xl = Format(LL, "#0.0#######")
  grdD.TextMatrix(j%, 1) = xl
  xt = Format(tT, "#0.0#######")
  grdD.TextMatrix(j%, 0) = xt
  
  'valort(j%) = val(grdD.TextMatrix(j%, 0))
  'valorqs(j%) = (grdD.TextMatrix(j%, 1))
  
 Next j%
 Close
 u = NumReg - 1

 Exit Sub
 
SinArchivo:
 If Errorumber = 32755 Then
  MsgBox "Error desconocido al abrir el archivo " & NombreArch
 End If
End Sub

Private Sub masis_Click()
frmriegosurcos.Show
End Sub

Private Sub mcond_Click()
frmconductividad.Show
End Sub

Private Sub mconvr_Click()
frmconvertidor.Show
End Sub

Private Sub mcrear_Click()
 On Error GoTo SinArchivo:
 ChDir App.Path
 cdCrear.ShowSave
 NombreArch = cdCrear.FileName
 ' Salvar archivo
 Open NombreArch For Random As #1 Len = Len(Paresqst1)
 If (LOF(1) <> 0) Then
  Close #1
  Kill NombreArch
 Open NombreArch For Random As #1 Len = Len(Paresqst1)
 End If
 For j% = 1 To (u + 1)
  Paresqst1.qs = valorqs(j%)
  Paresqst1.tT = valort(j%)
  km(1) = Val(txtQE.text)
  Paresqst1.kkqe = km(1)
  km(2) = Val(txtEf.text)
  Paresqst1.kkef = km(2)
  km(3) = Val(txtLn.text)
  Paresqst1.kkln = km(3)
  km(4) = Val(txtW.text)
  Paresqst1.kkw = km(4)
  km(5) = Val(txtL.text)
  Paresqst1.kkl = km(5)
  km(6) = Val(Text1.text)
  Paresqst1.kktext1 = km(6)
  km(7) = Val(Text2.text)
  Paresqst1.kktext2 = km(7)
  Put #1, j%, Paresqst1
 Next j%
 Close

 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al salvar el archivo " & NombreArch
 End If
End Sub


Private Sub mevap_Click()
frmETO.Show
End Sub

Private Sub mmp_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub mpa_Click()
frmpruebaavance.Show

End Sub

Private Sub mpasuel1_Click()
frmgenerales.Show
End Sub

Private Sub mregpot_Click()
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

Private Sub mtex_Click()
frmtextura.Show
End Sub

Private Sub Option1_Click()
lblund.Caption = "l/s"

grdD.TextMatrix(0, 1) = "Qsal l/s"
End Sub
Private Sub Option2_Click()
lblund.Caption = "l/min"

grdD.TextMatrix(0, 1) = "Qsal l/min"


End Sub


Private Sub TabStrip2_Click()
'On Error GoTo mensaje:
s = TabStrip2.SelectedItem.Index
Select Case s
    Case 1
    Frame2.Visible = True
    Frame4.Visible = False
    Frame8.Visible = False
    StatusBar1.Panels(1).text = "Digite los datos generales del sistema de riego e introduzca los valores de la prueba de Surco Infiltrómetro"
    
    Case 2
    Frame2.Visible = False
    Frame4.Visible = True
    Frame8.Visible = False
    StatusBar1.Panels(1).text = "Oprima el botón de Ecuaciones y Calcular, respectivamente, para determinar los resultados de la Prueba"
    
    Case 3

    Frame2.Visible = False
    Frame8.Visible = True
    Frame4.Visible = False
    StatusBar1.Panels(1).text = ""
    If n <= 1 Then
        Graph1.DrawMode = gphClear
        Graph2.DrawMode = gphClear
        Graph3.DrawMode = gphClear
        Exit Sub
    End If
    numpuntos = n
    
    'grafico de caudal de salida
     Graph1.FontUse = 4
       Graph1.GraphType = 6
       Graph1.GraphStyle = 5
       Graph1.BorderStyle = 1
       Graph1.AutoInc = 0
       Graph1.NumPoints = numpuntos
       Graph1.NumSets = 1
       For j% = 1 To numpuntos
        Graph1.ThisPoint = j%
        Graph1.GraphData = valorqs(j%)
        Graph1.XPosData = valorta(j%)
       Next j%
       Graph1.DrawMode = 2

    
    'grafico de caudal infiltrado
     Graph2.FontUse = 4
       Graph2.GraphType = 6
       Graph2.GraphStyle = 5
       Graph1.BorderStyle = 1
       Graph2.AutoInc = 0
       Graph2.NumPoints = numpuntos
       Graph2.NumSets = 1
       For j% = 1 To numpuntos
        Graph2.ThisPoint = j%
        Graph2.GraphData = valorqi(j%)
        Graph2.XPosData = valorta(j%)
       Next j%
       Graph2.DrawMode = 2
       
    'grafico de caudal de infiltracion acumulada
     Graph3.FontUse = 4
       Graph3.GraphType = 6
       Graph3.GraphStyle = 5
       Graph1.BorderStyle = 1
       Graph3.AutoInc = 0
       Graph3.NumPoints = numpuntos
       Graph1.NumSets = 1
       For j% = 1 To numpuntos
        Graph3.ThisPoint = j%
        Graph3.GraphData = valoria(j%)
        Graph3.XPosData = valorta(j%)
       Next j%
       Graph3.DrawMode = 2
  
    
End Select
End Sub

