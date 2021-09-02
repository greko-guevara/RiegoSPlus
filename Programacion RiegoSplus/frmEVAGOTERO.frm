VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmEVAGOTERO 
   Caption         =   "Evaluación de riego por goteo"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   Icon            =   "frmEVAGOTERO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11850
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Evaluación de Sistema"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Evaluación de goteo"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   2918
      TabIndex        =   3
      Top             =   6480
      Width           =   5775
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   3960
         Picture         =   "frmEVAGOTERO.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2040
         Picture         =   "frmEVAGOTERO.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   240
         Picture         =   "frmEVAGOTERO.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid griDord 
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   250
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7785
      Width           =   11850
      _ExtentX        =   20902
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
            TextSave        =   "08/07/2005"
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
   Begin MSComDlg.CommonDialog cdCrear 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.LT)"
   End
   Begin MSComDlg.CommonDialog cdAccesar 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar el archivo a cargar"
      Filter          =   "(*.LT)"
   End
   Begin VB.Frame Fgotero 
      Caption         =   "Evaluación del gotero"
      ForeColor       =   &H00800000&
      Height          =   5415
      Left            =   960
      TabIndex        =   45
      Top             =   960
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CommandButton BcalGOTERO 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   480
         Picture         =   "frmEVAGOTERO.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3840
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   4815
         Left            =   2280
         TabIndex        =   46
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   8493
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         GridColor       =   16761024
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Abrir"
         Height          =   375
         Left            =   1200
         TabIndex        =   62
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   4680
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resultados estadísticos del gotero"
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   5160
         TabIndex        =   48
         Top             =   600
         Width           =   3975
         Begin VB.TextBox txtCVG 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   52
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtMG 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   51
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtdSg 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   50
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtVcIg 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   49
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Media"
            Height          =   255
            Left            =   360
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "C. V. del fabricante"
            Height          =   375
            Left            =   360
            TabIndex        =   55
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Valor del cuarto inferior"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label16 
            Caption         =   "Desviación estándar"
            Height          =   375
            Left            =   360
            TabIndex        =   53
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.TextBox txtPg 
         Height          =   285
         Left            =   360
         TabIndex        =   57
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2325
         Left            =   5880
         Picture         =   "frmEVAGOTERO.frx":29F2
         Top             =   2760
         Width           =   2850
      End
      Begin VB.Label Label25 
         Caption         =   "Presión de prueba"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label24 
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
         Height          =   615
         Left            =   240
         TabIndex        =   59
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label23 
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
         Height          =   615
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Fsistema 
      Caption         =   "Evaluación del sistema"
      ForeColor       =   &H00800000&
      Height          =   5415
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   9855
      Begin VB.Frame Frame6 
         Height          =   3855
         Left            =   1560
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   6255
         Begin VB.CommandButton Command5 
            Caption         =   "Regresar"
            Height          =   735
            Left            =   2280
            Picture         =   "frmEVAGOTERO.frx":17484
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   3000
            Width           =   1575
         End
         Begin GraphLib.Graph Graph1 
            Height          =   2415
            Left            =   240
            TabIndex        =   67
            Top             =   480
            Visible         =   0   'False
            Width           =   5535
            _Version        =   65536
            _ExtentX        =   9763
            _ExtentY        =   4260
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
         Begin VB.Label Label27 
            Caption         =   "Caudal"
            Height          =   255
            Left            =   840
            TabIndex        =   70
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Presión"
            Height          =   255
            Left            =   4680
            TabIndex        =   69
            Top             =   3000
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Graficar"
         Height          =   615
         Left            =   3120
         Picture         =   "frmEVAGOTERO.frx":1814E
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Abrir"
         Height          =   375
         Left            =   1200
         TabIndex        =   64
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtCVG1 
         Height          =   285
         Left            =   360
         TabIndex        =   43
         Text            =   "1"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ecuación del gotero"
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   5400
         TabIndex        =   31
         Top             =   240
         Width           =   3975
         Begin VB.TextBox Text1 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   4800
            TabIndex        =   35
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtab 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1560
            TabIndex        =   34
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtb 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   3240
            TabIndex        =   33
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txta 
            BackColor       =   &H80000016&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1320
            TabIndex        =   32
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1080
            TabIndex        =   42
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "H"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   41
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "K"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   720
            TabIndex        =   40
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Q="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Correlación"
            Height          =   255
            Left            =   3600
            TabIndex        =   38
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Valor de b"
            Height          =   255
            Left            =   2160
            TabIndex        =   37
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Valor de a"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.CommandButton bcalcular 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   480
         Picture         =   "frmEVAGOTERO.frx":188B8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Frame frame5 
         Caption         =   "Resultados estadísticos para el caudal"
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   5400
         TabIndex        =   17
         Top             =   1560
         Width           =   3975
         Begin VB.TextBox txtCI 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   21
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtdesvio 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   20
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtmedia 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtDS 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   18
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Desviación estándar"
            Height          =   375
            Left            =   360
            TabIndex        =   25
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label20 
            Caption         =   "Valor del cuarto inferior"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label15 
            Caption         =   "C. V. del sistema"
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Media"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Coeficientes de Uniformidad"
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   5400
         TabIndex        =   10
         Top             =   3600
         Width           =   3975
         Begin VB.TextBox txtCUC 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TXTCUK 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   12
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtCUP 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2280
            TabIndex        =   11
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "CU de  Caudales"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label13 
            Caption         =   "CU Christiansen (Caudal)"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label11 
            Caption         =   "CU de Presiones"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   2055
         End
      End
      Begin VB.TextBox txtE 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Text            =   "1"
         Top             =   2400
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid grdDatos 
         Height          =   4215
         Left            =   2400
         TabIndex        =   26
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7435
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   16761024
      End
      Begin VB.Label Label10 
         Caption         =   "Coeficiente varianza del gotero"
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label8 
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
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1935
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
         Height          =   615
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label37 
         Caption         =   "# de Goteros por planta"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   1815
      End
   End
   Begin VB.Label Label18 
      Caption         =   "Evaluación de riego por goteo"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Menu marchivo 
      Caption         =   "Diseño agronómico goteo"
   End
   Begin VB.Menu mhidr 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu mla 
         Caption         =   "Cálculos en laterales"
      End
      Begin VB.Menu gg 
         Caption         =   "Cálculos en principales"
      End
   End
   Begin VB.Menu qpqpq 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmEVAGOTERO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X1(1 To 2) As Double
Dim x2(1 To 2) As Double
Dim x3(1 To 2) As Double
Dim valorl(1 To 300) As Double
Dim valort(1 To 300) As Double
Dim valorQ(0 To 300) As Double
Dim ppp(1 To 250) As Double
Dim pPord(1 To 250) As Double
Dim ppp1(1 To 250) As Double
Dim pPord1(1 To 250) As Double
Dim w As Integer
Dim u As Integer
Dim a As Double
Dim b As Double
Rem funcion de la grid
Dim row, col, numero, num(0 To 1000)
Dim n As Integer, i, ii, punto
Dim g As Integer



Private Sub Bcalcular_Click()
On Error GoTo mensaje:
'ECUACION DEL GOTERO***************************************
If u + 1 < 4 Then
   MsgBox "Ingrese al menos cuatro pares de valores", 64, "Imposible Calcular"
Exit Sub
End If

    n = u + 1
    NN = CInt(n / 4)
    sl = 0
    sl2 = 0
    st = 0
    st2 = 0
    slt = 0
    For j% = 1 To n
        valorl(j%) = Val(grdDatos.TextMatrix(j%, 0))
        valort(j%) = Val(grdDatos.TextMatrix(j%, 1))
        sl = sl + Log(valorl(j%))
        sl2 = sl2 + Log(valorl(j%)) ^ 2
        st = st + Log(valort(j%))
        st2 = st2 + Log(valort(j%)) ^ 2
        slt = slt + Log(valorl(j%)) * Log(valort(j%))
    Next j%
    b = (n * slt - sl * st) / (n * sl2 - sl ^ 2)
    a = st / n - sl / n * b
    aa = Exp(a)
    rr = (slt - n * sl / n * st / n) / Sqr((st2 - n * (st / n) ^ 2) * (sl2 - n * (sl / n) ^ 2))
     Text1.text = Format(rr, "###0.0###")
    txta.text = Format(aa, "###0.0###")
    txtb.text = Format(b, "###0.0###")
        txtab.text = "Q=" + Format(aa, "###0.0###") + "*P^" + Format(b, "###0.0###")

'PARAMETROS ESTADISTICOS*****************************************
sq = 0
For j% = 1 To n
    sq = sq + Val(grdDatos.TextMatrix(j%, 1))
Next j%
media = sq / n
txtmedia = Format(media, "##0.00#")
desvio = 0
For j% = 1 To n
        desvio1 = (media - Val(grdDatos.TextMatrix(j%, 1))) ^ 2
        desvio = desvio1 + desvio
Next j%
desvio = (desvio / (n - 1)) ^ 0.5
txtdesvio = Format(desvio, "##0.00#")
cv = desvio / media
txtDS = Format(cv, "##0.00#")

'******************cuato inferior*********************
 For j% = 1 To n
  pPord(j%) = valort(j%)
 Next j%
' Ordenar de menor a mayor
 m = n - 1
 For j% = 1 To m
  k% = j% + 1
  For jj% = j% To n
   If pPord(jj%) < pPord(k%) Then k% = jj%
  Next jj%
  temp = pPord(j%)
  pPord(j%) = pPord(k%)
  pPord(k%) = temp
 Next j%


cuarto = 0
For j% = 1 To NN
    cuarto = cuarto + pPord(j%)
Next j%
cuainf = cuarto / (NN)
txtCI = Format(cuainf, "##0.00#")

'cuarto inferior de presion******************************
 For j% = 1 To n
  ppp(j%) = valorl(j%)
 Next j%
' Ordenar de menor a mayor
 m = n - 1
 For j% = 1 To m
  k% = j% + 1
  For jj% = j% To n
   If ppp(jj%) < ppp(k%) Then k% = jj%
  Next jj%
  temp1 = ppp(j%)
  ppp(j%) = ppp(k%)
  ppp(k%) = temp1
 Next j%

cuartop = 0
For j% = 1 To NN
    cuartop = cuartop + ppp(j%)
Next j%
cuainfp = cuartop / (NN)
'coeficientes de uniformidad********************************
cuc = cuainf / media * 100
txtCUC = Format(cuc, "##0.00#")
E = Val(txtE.text)
ccvv = Val(txtCVG1.text)
If ccvv = 0 Then
    MsgBox "Ingrese el coeficiente de variación del gotero", 64, "Evaluación riego por goteo"
    txtCVG1.SetFocus
    Exit Sub
End If
cuk = 100 * (1 - 1.27 * ccvv / E ^ 0.5) * (pPord(1) / media)
TXTCUK = Format(cuk, "##0.00#")
Sp = 0
For j% = 1 To n
    Sp = Sp + Val(grdDatos.TextMatrix(j%, 0))
Next j%
mediap = Sp / n
cup = (cuainfp / mediap) ^ b * 100
txtCUP = Format(cup, "##0.00#")
Exit Sub
mensaje:
MsgBox "Ingrese datos adecuados", 64, "Imposible Calcular"

End Sub

Private Sub BcalGOTERO_Click()
If w + 1 < 4 Then
   MsgBox "Ingrese al menos cuatro pares de valores", 64, "Evaluación de riego por gotei"
Exit Sub
End If
    g = w + 1
    NNg = CInt(g / 4)
    sqg = 0
    For Y% = 1 To g
        valorQ(Y%) = Val(grid1.TextMatrix(Y%, 0))
        sqg = sqg + valorQ(Y%)
    Next Y%
'PARAMETROS ESTADISTICOS*****************************************
mg = sqg / g
txtMG = Format(mg, "##0.00#")
desvioo = 0
For j% = 1 To g
        desvioo1 = (mg - Val(grid1.TextMatrix(j%, 0))) ^ 2
        desvioo = desvioo1 + desvioo
Next j%
desvioo = (desvioo / (g - 1)) ^ 0.5
TxtdSg = Format(desvioo, "##0.00#")
cvg = desvioo / mg
txtCVG = Format(cvg, "##0.00#")

'******************cuato inferior*********************
 For j% = 1 To g
  pPord1(j%) = valorQ(j%)
 Next j%
' Ordenar de menor a mayor
 m = g - 1
 For j% = 1 To m
  k% = j% + 1
  For jj% = j% To g
   If pPord1(jj%) < pPord1(k%) Then k% = jj%
  Next jj%
  temp2 = pPord1(j%)
  pPord1(j%) = pPord1(k%)
  pPord1(k%) = temp2
 Next j%
cuartoo = 0
For j% = 1 To NNg
    cuartoo = cuartoo + pPord1(j%)
Next j%
vci = cuartoo / (NNg)
txtVcIg = Format(vci, "##0.00#")
Exit Sub
End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
w = 0
txtCVG1.text = ""
txtPg.text = ""
txta.text = ""
txtb.text = ""
txtab.text = ""
Text1.text = ""
txtmedia.text = ""
txtdesvio.text = ""
txtDS.text = ""
txtCI.text = ""
txtCUC.text = ""
TXTCUK.text = ""
txtCUP.text = ""
txtMG.text = ""
TxtdSg.text = ""
txtCVG.text = ""
txtVcIg.text = ""
grdDatos.Clear
grdDatos.Rows = 2
grdDatos.TextMatrix(0, 0) = "Presión"
grdDatos.TextMatrix(0, 1) = "Caudal"
grid1.Clear
grid1.Rows = 2
grid1.TextMatrix(0, 0) = "Caudales"
u = 0
End Sub



Private Sub Command1_Click()
On Error GoTo SinArchivo
 ChDir App.Path
 cdCrear.ShowSave
 NombreArch = cdCrear.FileName
 ' Salvar archivo
 Open NombreArch For Random As #1 Len = Len(Parestl)
 If (LOF(1) <> 0) Then
  Close #1
  Kill NombreArch
  Open NombreArch For Random As #1 Len = Len(Parestl)
 End If
 For j% = 1 To (u + 1)
  X1(1) = Val(txtCVG1.text)
  x2(1) = Val(txtE.text)
  Parestl.xx1 = X1(1)
  Parestl.xx2 = x2(1)
  Parestl.L = valorl(j%)
  Parestl.T = valort(j%)
  Put #1, j%, Parestl
 Next j%
 Close

 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al salvar el archivo " & NombreArch
 End If
End Sub

Private Sub Command2_Click()
On Error GoTo SinArchivo
 ChDir App.Path
 cdCrear.ShowSave
 NombreArch = cdCrear.FileName
 ' Salvar archivo
 Open NombreArch For Random As #1 Len = Len(caudales)
 If (LOF(1) <> 0) Then
  Close #1
  Kill NombreArch
  Open NombreArch For Random As #1 Len = Len(caudales)
 End If
 For j% = 1 To (w + 1)
  x3(1) = Val(txtPg.text)
  caudales.xx3 = x3(1)
  caudales.q = valorQ(j%)
  Put #1, j%, caudales
 Next j%
 Close

 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al salvar el archivo " & NombreArch
 End If
End Sub

Private Sub Command3_Click()
On Error GoTo SinArchivo
cdAccesar.ShowOpen
 NombreArch = cdAccesar.FileName
 w = 0
 Open NombreArch For Random As #1 Len = Len(caudales)
 NumReg = LOF(1) \ Len(caudales)
 grid1.Rows = NumReg + 1
 For j% = 1 To NumReg
  Get #1, j%, caudales
  LL = caudales.q
  txtPg = caudales.xx3
  xl = Format(LL, "#0.0#######")
  grid1.TextMatrix(j%, 0) = xl
  valorQ(j%) = Val(grid1.TextMatrix(j%, 0))
  
 Next j%
 Close
 w = NumReg - 1
 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al abrir el archivo " & NombreArch
 End If
End Sub

Private Sub Command4_Click()
cdAccesar.ShowOpen
 NombreArch = cdAccesar.FileName
 u = 0
 Open NombreArch For Random As #1 Len = Len(Parestl)
 NumReg = LOF(1) \ Len(Parestl)
 grdDatos.Rows = NumReg + 1
 For j% = 1 To NumReg
  Get #1, j%, Parestl
  LL = Parestl.L
  tT = Parestl.T
  txtE = Parestl.xx2
  txtCVG1 = Parestl.xx1
  xl = Format(LL, "#0.0#######")
  grdDatos.TextMatrix(j%, 0) = xl
  xt = Format(tT, "#0.0#######")
  grdDatos.TextMatrix(j%, 1) = xt
  
  valorl(j%) = Val(grdDatos.TextMatrix(j%, 0))
  valort(j%) = Val(grdDatos.TextMatrix(j%, 1))
  
 Next j%
 Close
 u = NumReg - 1

 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al abrir el archivo " & NombreArch
 End If
End Sub

Private Sub Command5_Click()
Frame6.Visible = False
End Sub

Private Sub Command6_Click()
'On Error GoTo mensaje:
Rem gráfico de la curva
    numpuntos = n
    If numpuntos <= 1 Then
      temp = MsgBox("Para graficar se requiere por lo menos dos puntos", 64, "Imposible Graficar")
    Else
     Graph1.FontUse = 4
       Graph1.GraphType = 9
       Graph1.GraphStyle = 0
       Graph1.BorderStyle = 1
       Graph1.Visible = True
       Label11.Visible = True
       Label10.Visible = True
       Graph1.AutoInc = 0
       Graph1.NumPoints = numpuntos
       Graph1.NumSets = 1
       For j% = 1 To numpuntos
        Graph1.ThisPoint = j%
        Graph1.GraphData = valort(j%)
        Graph1.XPosData = valorl(j%)
       Next j%

  Graph1.DrawMode = 2
  Frame6.Visible = True
     End If
Exit Sub
mensaje:
   MsgBox "Imposible graficar", 16, " Error"
End Sub

Private Sub Form_Load()
u = 0
w = 0

grid1.ColWidth(0) = 1300
grid1.ColAlignment(0) = 4
grid1.TextMatrix(0, 0) = "Caudales"
grdDatos.ColWidth(0) = 1300
grdDatos.ColWidth(1) = 1300
grdDatos.ColAlignment(0) = 4
grdDatos.ColAlignment(1) = 4
grdDatos.TextMatrix(0, 0) = "Presión"
grdDatos.TextMatrix(0, 1) = "Caudal"
StatusBar1.Panels(1).text = "Ingrese los datos de presión vrs caudal así como las características del sistema"
End Sub

Private Sub gg_Click()
frmHprincipal.Show
End Sub

Private Sub grddatos_Click()
i = ""
punto = 0

End Sub
Private Sub grddatos_KeyPress(KeyAscii As Integer)

If grdDatos.col <> col Or grdDatos.row <> row Then
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
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 48 Then
    i = i + "0"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If


If punto <> 1 Then
If KeyAscii = 44 Or KeyAscii = 46 Then
    numero = numero + 1
    If i = "" Then
    i = i + "0."
    grdDatos.text = i
    num(numero) = i
    punto = 1
Else
    i = i + "."
    grdDatos.text = i
    num(numero) = i
    punto = 1
End If
End If
End If


If KeyAscii = 49 Then
    i = i + "1"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 50 Then
    i = i + "2"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 51 Then
    i = i + "3"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 52 Then
    i = i + "4"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 53 Then
    i = i + "5"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 54 Then
    i = i + "6"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 55 Then
    i = i + "7"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 56 Then
    i = i + "8"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 57 Then
    i = i + "9"
    grdDatos.text = i
    numero = numero + 1
    num(numero) = i
End If
Rem tecla de borrado
If numero >= 1 Then
If KeyAscii = 8 Then
i = num(numero - 1)
numero = numero - 1
grdDatos.text = i
End If
Else

grdDatos.text = ""
End If

Rem tecla para eliminar
If KeyAscii = 42 Then
If u >= 2 Then
    u = u - 1
    grdDatos.Rows = u + 2
End If
End If
If KeyAscii = 13 Then
u = u + 1
grdDatos.Rows = u + 2
End If

Rem pruevas grid1.TextMatrix(numero, 6) = num(numero)

Rem grdDatos.Text = KeyAscii
col = grdDatos.col
row = grdDatos.row

End Sub


Private Sub marchivo_Click()
frmDAgoteo.Show
End Sub

Private Sub mla_Click()
FrmHLaterales.Show
End Sub

Private Sub qpqpq_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    Fgotero.Visible = False
    Fsistema.Visible = True
    StatusBar1.Panels(1).text = "Ingrese los diferentes datos de caudal para un mismo caudal"
    txtCVG1.text = txtCVG.text
    Case 2
    Fsistema.Visible = False
    Fgotero.Visible = True
    StatusBar1.Panels(1).text = "Ingrese los datos de presión vrs caudal así como las características del sistema"
End Select

End Sub
Private Sub Grid1_Click()
i = ""
punto = 0

End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)

If grid1.col <> col Or grid1.row <> row Then
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
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 48 Then
    i = i + "0"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If


If punto <> 1 Then
If KeyAscii = 44 Or KeyAscii = 46 Then
    numero = numero + 1
    If i = "" Then
    i = i + "0."
    grid1.text = i
    num(numero) = i
    punto = 1
Else
    i = i + "."
    grid1.text = i
    num(numero) = i
    punto = 1
End If
End If
End If


If KeyAscii = 49 Then
    i = i + "1"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 50 Then
    i = i + "2"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 51 Then
    i = i + "3"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 52 Then
    i = i + "4"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 53 Then
    i = i + "5"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 54 Then
    i = i + "6"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 55 Then
    i = i + "7"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 56 Then
    i = i + "8"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 57 Then
    i = i + "9"
    grid1.text = i
    numero = numero + 1
    num(numero) = i
End If
Rem tecla de borrado
If numero >= 1 Then
If KeyAscii = 8 Then
i = num(numero - 1)
numero = numero - 1
grid1.text = i
End If
Else
grid1.text = ""
End If

Rem tecla para eliminar
If KeyAscii = 42 Then
If w >= 2 Then
    w = w - 1
    grid1.Rows = w + 2
End If
End If
If KeyAscii = 13 Then
w = w + 1
grid1.Rows = w + 2
End If

Rem pruevas grid1.TextMatrix(numero, 6) = num(numero)

Rem grdDatos.Text = KeyAscii
col = grid1.col
row = grid1.row

End Sub

