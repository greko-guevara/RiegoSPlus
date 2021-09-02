VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmevalucion 
   Caption         =   "Evaluación de sistemas de riego por aspersión"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   Icon            =   "frmevalucionaspersion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11850
   Begin VB.Frame Frame2 
      ForeColor       =   &H00800000&
      Height          =   6375
      Left            =   4920
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Frame Frame9 
         Height          =   1815
         Left            =   120
         TabIndex        =   40
         Top             =   4440
         Width           =   6495
         Begin VB.TextBox txtCI 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   46
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtdesvio 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   960
            TabIndex        =   45
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCU 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   4920
            TabIndex        =   44
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCD 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   4920
            TabIndex        =   43
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtmedia 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   960
            TabIndex        =   42
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtIP 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   41
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label25 
            Caption         =   "%"
            Height          =   255
            Left            =   6240
            TabIndex        =   55
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label24 
            Caption         =   "%"
            Height          =   255
            Left            =   6240
            TabIndex        =   54
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label40 
            Caption         =   "mm/hr"
            Height          =   255
            Left            =   3360
            TabIndex        =   53
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label21 
            Caption         =   "Intensidad de aplicación promedio"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label15 
            Caption         =   "Desvio"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Coeficiente de distribución"
            Height          =   255
            Left            =   2760
            TabIndex        =   50
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
            Caption         =   "Coeficiente de Uniformidad"
            Height          =   255
            Left            =   2760
            TabIndex        =   49
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label14 
            Caption         =   "Media"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Valor del cuarto inferior"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   6375
         Begin MSFlexGridLib.MSFlexGrid grid1 
            Height          =   2175
            Left            =   480
            TabIndex        =   6
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3836
            _Version        =   393216
            FixedRows       =   0
            FixedCols       =   0
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   360
            Y1              =   120
            Y2              =   2040
         End
         Begin VB.Line Line3 
            X1              =   240
            X2              =   2160
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Label7 
            Caption         =   "(0,0)"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "X"
            Height          =   255
            Left            =   2160
            TabIndex        =   16
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "Y"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2040
            Width           =   255
         End
      End
      Begin VB.CommandButton bcalcular 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   1680
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmevalucionaspersion.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         Caption         =   "¿Qué está ingresando?"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   360
         TabIndex        =   35
         Top             =   240
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3960
            TabIndex        =   39
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Volúmenes"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Láminas"
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   1560
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Unidades en:"
            Height          =   255
            Left            =   2760
            TabIndex        =   38
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Graficar"
         Height          =   615
         Left            =   3720
         Picture         =   "frmevalucionaspersion.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Posición aspersor"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   6960
      TabIndex        =   56
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   720
         TabIndex        =   58
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   2760
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "En Y"
         Height          =   255
         Left            =   2280
         TabIndex        =   60
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "En X"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid griDord 
      Height          =   1215
      Left            =   8160
      TabIndex        =   33
      Top             =   -120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   250
   End
   Begin VB.Frame Frame8 
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   4695
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   3120
         Picture         =   "frmevalucionaspersion.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   120
         Picture         =   "frmevalucionaspersion.frx":2308
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   1680
         Picture         =   "frmevalucionaspersion.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Entrada"
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4695
      Begin VB.CommandButton bingresar 
         Caption         =   "Ingresar datos de la prueba"
         Height          =   495
         Left            =   1440
         TabIndex        =   5
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtnf 
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtnc 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtt 
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtec 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtar 
         Height          =   285
         Left            =   2640
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "mim"
         Height          =   255
         Left            =   3960
         TabIndex        =   32
         Top             =   1080
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8040
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label10 
         Caption         =   "Número de columnas"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Número de filas"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label43 
         Caption         =   "Espaciamiento cuadrículas"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label42 
         Caption         =   "Tiempo de la prueba"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label39 
         Caption         =   "m"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label37 
         Caption         =   "Area del recipiente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "cm2"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   7770
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
            TextSave        =   "14/05/2007"
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
      Left            =   0
      Top             =   360
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
   Begin VB.Frame frame5 
      Caption         =   "Modelaciones"
      ForeColor       =   &H00800000&
      Height          =   6135
      Left            =   4920
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   6255
         Begin VB.CommandButton batras 
            Caption         =   "Atras"
            Height          =   375
            Left            =   4080
            TabIndex        =   87
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton bmodular 
            Caption         =   "Modelar"
            Height          =   375
            Left            =   4080
            TabIndex        =   24
            Top             =   0
            Width           =   1455
         End
         Begin VB.TextBox txtAA 
            Height          =   285
            Left            =   1920
            TabIndex        =   21
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtLL 
            Height          =   285
            Left            =   1920
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Entre Laterales"
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label16 
            Caption         =   "Entre Aspersores"
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1815
         Left            =   120
         TabIndex        =   68
         Top             =   4200
         Width           =   6495
         Begin VB.TextBox Tia 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   74
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox tMed 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   960
            TabIndex        =   73
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox tCD 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   4920
            TabIndex        =   72
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Tcu 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   4920
            TabIndex        =   71
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox Tdes 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   960
            TabIndex        =   70
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox Tci 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   69
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "Valor del cuarto inferior"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label41 
            Caption         =   "Media"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label38 
            Caption         =   "Coeficiente de Uniformidad"
            Height          =   255
            Left            =   2760
            TabIndex        =   84
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label36 
            Caption         =   "Coeficiente de distribución"
            Height          =   255
            Left            =   2760
            TabIndex        =   83
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label35 
            Caption         =   "Desvio"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label34 
            Caption         =   "Intensidad de aplicación promedio"
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label33 
            Caption         =   "mm/hr"
            Height          =   255
            Left            =   3360
            TabIndex        =   80
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label32 
            Caption         =   "mm"
            Height          =   255
            Left            =   3360
            TabIndex        =   79
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label31 
            Caption         =   "%"
            Height          =   255
            Left            =   6240
            TabIndex        =   78
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label30 
            Caption         =   "%"
            Height          =   255
            Left            =   6240
            TabIndex        =   77
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label29 
            Caption         =   "min"
            Height          =   255
            Left            =   2280
            TabIndex        =   76
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label28 
            Caption         =   "mm"
            Height          =   255
            Left            =   2280
            TabIndex        =   75
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   2775
         Left            =   120
         TabIndex        =   63
         Top             =   1200
         Width           =   6375
         Begin MSFlexGridLib.MSFlexGrid Grid2 
            Height          =   2175
            Left            =   480
            TabIndex        =   64
            Top             =   480
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3836
            _Version        =   393216
            FixedRows       =   0
            FixedCols       =   0
         End
         Begin VB.Label Label23 
            Caption         =   "Y"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "X"
            Height          =   255
            Left            =   2160
            TabIndex        =   66
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "(0,0)"
            Height          =   255
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   375
         End
         Begin VB.Line Line5 
            X1              =   240
            X2              =   2160
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line4 
            X1              =   360
            X2              =   360
            Y1              =   120
            Y2              =   2040
         End
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2190
      Left            =   480
      Picture         =   "frmevalucionaspersion.frx":315C
      Top             =   4200
      Width           =   3825
   End
   Begin VB.Label Label18 
      Caption         =   "Evaluación de riego por aspersión"
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
      Left            =   840
      TabIndex        =   29
      Top             =   360
      Width           =   5175
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mabrir 
         Caption         =   "Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mguardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu daa 
      Caption         =   "Diseño agronómico en aspersión"
   End
   Begin VB.Menu sksksksk 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu leds 
         Caption         =   "Cálculos en laterales"
      End
      Begin VB.Menu wqw 
         Caption         =   "Cálculos en principales"
      End
   End
   Begin VB.Menu hghghg 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmevalucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem funcion de la grid
Dim row, col, numero, num(0 To 1000)
Dim n As Integer, i, ii, punto
Dim j As Integer
Dim k As Integer
Dim hhh As Double
Dim ppp(1 To 250) As Double
Dim ppp1(1 To 250) As Double
Dim pPord(1 To 250) As Double
Dim pPord1(1 To 250) As Double
Dim cccc(1 To 5) As Double


Private Sub batras_Click()
Frame5.Visible = False
Frame2.Visible = True

End Sub

Private Sub Bcalcular_Click()
On Error GoTo mensaje:

ar = Val(txtAR.text)
tT = Val(txtT.text)
nc = Val(txtnc.text)
nf = Val(txtnf.text)
ec = Val(txtec.text)
If ar = 0 Then
MsgBox "Ingrese el área del recipiente", 64, "Evaluación en riego por aspersión"
txtAR.SetFocus
Exit Sub
End If
If ec = 0 Then
MsgBox "Ingrese el espaciamiento entre recipiente", 64, "Evaluación en riego por aspersión"
txtec.SetFocus
Exit Sub
End If
If tT = 0 Then
MsgBox "Ingrese el tiempo de la prueba", 64, "Evaluación en riego por aspersión"
txtT.SetFocus
Exit Sub
End If
If nc = 0 Then
MsgBox "Ingrese el número de columnas", 64, "Evaluación en riego por aspersión"
txtnc.SetFocus
Exit Sub
End If
If nf = 0 Then
MsgBox "Ingrese el número de filas", 64, "Evaluación en riego por aspersión"
txtnf.SetFocus
Exit Sub
End If


n = nc * nf
hhh = n
NN = CInt(n / 4)
If NN = 0 Then
MsgBox "Ingrese al menos cuatro datos del Grid", 64, "Evaluación en riego por aspersión"
txtnf.SetFocus
Exit Sub
End If
'media
suma = 0
m = 0
For j% = 1 To nc
    For k% = 1 To nf
        If Grid1.TextMatrix(k% - 1, j% - 1) = "" Then
        MsgBox "Ingrese los valores de la prueba", 64, "Evaluación en riego por aspersión"
        Grid1.SetFocus
        Exit Sub
        End If
        suma = suma + Val(Grid1.TextMatrix(k% - 1, j% - 1))
        m = m + 1
        ppp(m) = Val(Grid1.TextMatrix(k% - 1, j% - 1))
    Next k%
Next j%
       
media = suma / n
txtmedia = Format(media, "##0.00#")
'desvio
desvio = 0
desvio3 = 0
For j% = 1 To nc
    For k% = 1 To nf
        desvio1 = Abs(media - Grid1.TextMatrix(k% - 1, j% - 1))
        desvio2 = (media - Val(Grid1.TextMatrix(k% - 1, j% - 1))) ^ 2
        desvio3 = (desvio2 + desvio3)
        desvio = desvio + desvio1
    Next k%
Next j%
desvio5 = (desvio3 / (n - 1)) ^ 0.5

'******************cuato inferior*********************

 For j% = 1 To n
  pPord(j%) = ppp(j%)
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
'Mostrar datos ordenados
For j% = 1 To n
 vx1 = pPord(j%)
 griDord.TextMatrix(j%, 0) = Str$(j%)
 xr = Format(vx1, "#0.0#######")
 griDord.TextMatrix(j%, 1) = xr
Next j%
cuarto = 0
For j% = 1 To NN
    cuarto = cuarto + pPord(j%)
Next j%
cuainf = cuarto / (NN)
'**************************
cd = cuainf / media * 100

txtCD = Format(cd, "##0.00#")
txtdesvio = Format(desvio5, "##0.00#")
cu = (1 - desvio / (media * n)) * 100
txtCU = Format(cu, "##0.00#")

'calculo de la lámina promedio aplicada
ar = ar * 100
UNIDADES = (Combo1.ListIndex)

If Option1.Value = True Then
    Select Case UNIDADES
        Case Is = 0
            med = media * 1000000
        Case Is = 2
            med = media * 1000
        Case Is = 1
            med = media
    End Select
    lpa = (med / ar) / tT / 60
Else
    Select Case UNIDADES
        Case Is = 2
        med = media * 1000
        Case Is = 1
        med = media * 10
        Case Is = 0
        med = media
    End Select
    lpa = med / tT * 60
End If

txtCI = Format(cuainf, "##0.00#")
txtIP = Format(lpa, "##0.00#")

StatusBar1.Panels(1).text = ""
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
Frame2.Visible = False
Frame5.Visible = False
txtAR.text = ""
txtec.text = ""
txtT.text = ""
txtnc.text = ""
txtnf.text = ""
txtX.text = ""
txtY.text = ""
txtmedia.text = ""
txtdesvio.text = ""
txtCD.text = ""
txtCU.text = ""
txtCI.text = ""
txtIP.text = ""
txtaa.text = ""
txtLL.text = ""
With Grid1
    .Clear
    .Cols = 0
    .Rows = 0
End With
End Sub

Private Sub bmodular_Click()
On Error GoTo mensaje:
Exit Sub
X = Val(txtX.text)
Y = Val(txtY.text)
tT = Val(txtT.text)
ec = Val(txtec.text)
aa = Val(txtaa.text)
LL = Val(txtLL.text)
nc = Val(txtnc.text)
nf = Val(txtnf.text)
a = aa / ec
L = LL / ec
    
With grid2
   .Clear
   .Rows = nf
   .Cols = nc
    For j = 1 To nc
        .ColWidth(j - 1) = 700
    Next j
'modulacion aspersor
    If txtaa <> "" Then
    aa1 = 1
      For j% = 1 To nc
         For k% = 1 To nf
            If k% <= nf / 2 Then
             grid2.TextMatrix(k% - 1, j% - 1) = Val(Grid1.TextMatrix(k% - 1, j% - 1)) * 2 + Val(Grid1.TextMatrix(k% + a - 1, j% - 1)) * 2
             Else
             grid2.TextMatrix(k% - 1, j% - 1) = Val(Grid1.TextMatrix(k% - 1, j% - 1)) * 2 + Val(Grid1.TextMatrix(k% - a - 1, j% - 1)) * 2
             End If
         Next k%
      Next j%
     End If
     'modulacion de laterales
     If txtLL <> "" Then
    LL = ec / LL
     For j% = 1 To nc
         For k% = 1 To nf
             If j% < nc / 2 Then
             grid2.TextMatrix(k% - 1, j% - 1) = Val(grid2.TextMatrix(k% - 1, j% - 1)) + Val(Grid1.TextMatrix(k% - 1, j% - 1)) * 2 + Val(Grid1.TextMatrix(k% - 1, j% + L - 1)) * 2
             Else
             grid2.TextMatrix(k% - 1, j% - 1) = Val(grid2.TextMatrix(k% - 1, j% - 1)) + Val(Grid1.TextMatrix(k% - 1, j% - 1)) * 2 + Val(Grid1.TextMatrix(k% - 1, j% - L - 1)) * 2
             End If
         Next k%
     Next j%
    End If
End With
'********************************************************************
n = nc * nf
hhh = n
NN = CInt(n / 4)
If NN = 0 Then
MsgBox "Ingrese al menos cuatro datos del Grid", 64, "Evaluación en riego por aspersión"
txtnf.SetFocus
Exit Sub
End If
'media
suma1 = 0
m1 = 0
For j% = 1 To nc
    For k% = 1 To nf
        If grid2.TextMatrix(k% - 1, j% - 1) = "" Then
        MsgBox "Ingrese los valores de la prueba", 64, "Evaluación en riego por aspersión"
        grid2.SetFocus
        Exit Sub
        End If
        suma1 = suma1 + grid2.TextMatrix(k% - 1, j% - 1)
        m = m + 1
        ppp1(m) = grid2.TextMatrix(k% - 1, j% - 1)
    Next k%
Next j%
       
media1 = suma1 / n
tMed = Format(media1, "##0.00#")
'desvio
desvio1 = 0
desvio31 = 0
For j% = 1 To nc
    For k% = 1 To nf
        desvio11 = Abs(media1 - grid2.TextMatrix(k% - 1, j% - 1))
        desvio21 = (media1 - grid2.TextMatrix(k% - 1, j% - 1)) ^ 2
        desvio31 = (desvio21 + desvio31)
        desvio1 = desvio1 + desvio11
    Next k%
Next j%
desvio51 = (desvio31 / (n - 1)) ^ 0.5

'******************cuato inferior*********************

 For j% = 1 To n
  pPord1(j%) = ppp1(j%)
 Next j%
' Ordenar de menor a mayor
 m1 = n - 1
 For j% = 1 To m1
  k% = j% + 1
  For jj% = j% To n
   If pPord1(jj%) < pPord1(k%) Then k% = jj%
  Next jj%
  temp1 = pPord1(j%)
  pPord1(j%) = pPord1(k%)
  pPord1(k%) = temp1
 Next j%
'Mostrar datos ordenados

cuarto1 = 0
For j% = 1 To NN
    cuarto1 = cuarto1 + pPord1(j%)
Next j%
cuainf1 = cuarto1 / (NN)
'**************************
cd1 = cuainf1 / media1 * 100

tCD = Format(cd1, "##0.00#")
Tdes = Format(desvio51, "##0.00#")
cu1 = (1 - desvio1 / (media1 * n)) * 100
Tcu = Format(cu1, "##0.00#")

'calculo de la lámina promedio aplicada
ar = ar * 100
UNIDADES = (Combo1.ListIndex)

If Option1.Value = True Then
    Select Case UNIDADES
        Case Is = 0
            med1 = media1 * 1000000
        Case Is = 2
            med1 = media1 * 1000
        Case Is = 1
            med1 = media1
    End Select
    lpa1 = (med1 / ar) / tT / 60
Else
    Select Case UNIDADES
        Case Is = 2
        med1 = media1 * 1000
        Case Is = 1
        med1 = media1 * 10
        Case Is = 0
        med1 = media1
    End Select
    lpa1 = med1 / tT * 60
End If

Tci = Format(cuainf1, "##0.00#")
Tia = Format(lpa1, "##0.00#")

StatusBar1.Panels(1).text = ""
'**************************************************************************************
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub



Private Sub Command1_Click()
Frame5.Visible = True
Frame2.Visible = False

End Sub

Private Sub daa_Click()
frmDAaspersion.Show
End Sub

Private Sub Form_Load()
StatusBar1.Panels(1).text = "Ingrese los datos básicos y oprima el botón de Ingresar datos para cargas los datos obtenidos durante la prueba"
With Combo1
    .AddItem "mm"
    .AddItem " cm"
    .AddItem " mts"
    .ListIndex = 0
    
End With

End Sub

Private Sub Grid1_Click()
i = ""
punto = 0
End Sub


Private Sub grid1_KeyPress(KeyAscii As Integer)

If Grid1.col <> col Or Grid1.row <> row Then
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
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 48 Then
    i = i + "0"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If


If punto <> 1 Then
If KeyAscii = 44 Or KeyAscii = 46 Then
    numero = numero + 1
    If i = "" Then
    i = i + "0."
    Grid1.text = i
    num(numero) = i
    punto = 1
Else
    i = i + "."
    Grid1.text = i
    num(numero) = i
    punto = 1
End If
End If
End If


If KeyAscii = 49 Then
    i = i + "1"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 50 Then
    i = i + "2"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 51 Then
    i = i + "3"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 52 Then
    i = i + "4"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 53 Then
    i = i + "5"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 54 Then
    i = i + "6"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 55 Then
    i = i + "7"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 56 Then
    i = i + "8"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 57 Then
    i = i + "9"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If
Rem tecla de borrado
If numero >= 1 Then
If KeyAscii = 8 Then
i = num(numero - 1)
numero = numero - 1
Grid1.text = i
End If
Else
Grid1.text = ""
End If

Rem pruebas grid1.TextMatrix(numero, 6) = num(numero)

Rem grid1.Text = KeyAscii
col = Grid1.col
row = Grid1.row

End Sub

Private Sub bingresar_Click()
On Error GoTo mensaje:
nc = Val(txtnc.text)
nf = Val(txtnf.text)
With Grid1
    .Cols = nc
    .Rows = nf
    For j = 1 To nc
    .ColWidth(j - 1) = 700
    Next j
End With
If nc < 2 And nf < 2 Then
MsgBox "La prueba debe contener al menos 2 filas y dos colomnnas", 64, "Evaluación en aspersión"
txtnc.SetFocus
Exit Sub
End If
Frame2.Visible = True
Frame3.Visible = False

Grid1.SetFocus
StatusBar1.Panels(1).text = "Ingrese los datos de la prueba para determinar sus coeficientes de uniformidad y distribución"
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub

Private Sub hghghg_Click()
Unload Me
frmGeneral.Show
End Sub



Private Sub leds_Click()
FrmHLaterales.Show
End Sub

Private Sub mabrir_Click()
On Error GoTo mensaje:
 cdAccesar.ShowOpen
 NombreArch = cdAccesar.FileName

 Open NombreArch For Random As #1 Len = Len(Paresqst)
 NumReg = LOF(1) \ Len(Paresqst)
 Get #1, (2), Paresqst
 txtAR.text = Paresqst.nar
 txtec.text = Paresqst.nec
 txtT.text = Paresqst.ntp
 With Grid1
    .Rows = Paresqst.numfila
    .Cols = Paresqst.numcol
    For jj = 1 To Paresqst.numcol
    .ColWidth(jj - 1) = 700
    Next jj
    End With
m = 0
 For j% = 1 To Paresqst.numcol
 For k% = 1 To Paresqst.numfila
 m = m + 1
  Get #1, (m), Paresqst
  tT = Paresqst.tT
  
  xt = Format(tT, "#0.0")
  Grid1.TextMatrix(k% - 1, j% - 1) = xt
  Next k%
 Next j%
 Close
 txtnc.text = Grid1.Cols
 txtnf.text = Grid1.Rows
 Frame2.Visible = True
 Exit Sub
 
mensaje:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al abrir el archivo " & NombreArch
 End If
End Sub

Private Sub mguardar_Click()
On Error GoTo SinArchivo:
 ChDir App.Path
 cdCrear.ShowSave
 NombreArch = cdCrear.FileName
 ' Salvar archivo
 Open NombreArch For Random As #1 Len = Len(Paresqst)
 If (LOF(1) <> 0) Then
  Close #1
  Kill NombreArch
  Open NombreArch For Random As #1 Len = Len(Paresqst)
 End If
 
 For j% = 1 To (hhh)
  Paresqst.tT = ppp(j%)
  cccc(1) = Val(Grid1.Cols)
 Paresqst.numcol = cccc(1)
  cccc(2) = Val(Grid1.Rows)
 Paresqst.numfila = cccc(2)
 cccc(3) = Val(txtAR.text)
 Paresqst.nar = cccc(3)
 cccc(4) = Val(txtec.text)
 Paresqst.nec = cccc(4)
 cccc(5) = Val(txtT.text)
 Paresqst.ntp = cccc(5)
  Put #1, j%, Paresqst
 Next j%
 Close

 Exit Sub
 
SinArchivo:

  MsgBox "Error desconocido al salvar el archivo " & NombreArch

End Sub

Private Sub Option1_Click()
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
With Combo1
    .Clear
    .AddItem "lts"
    .AddItem " cm3"
    .AddItem " mm3"
    .ListIndex = 0
End With
    
End Sub

Private Sub Option2_Click()
Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
With Combo1
    .Clear
    .AddItem "mm"
    .AddItem "cm"
    .AddItem "mts"
    .ListIndex = 0
End With

End Sub

Private Sub wqw_Click()
frmHprincipal.Show
End Sub
