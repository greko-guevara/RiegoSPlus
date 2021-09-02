VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmpruebaavance 
   Caption         =   "Prueba de avance en surcos"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmpruebaavance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   1200
      TabIndex        =   14
      Top             =   6240
      Width           =   9255
      Begin VB.CommandButton Command1 
         Caption         =   "&Graficar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1920
         Picture         =   "frmpruebaavance.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bcalcular 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   120
         Picture         =   "frmpruebaavance.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   3720
         Picture         =   "frmpruebaavance.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   5520
         Picture         =   "frmpruebaavance.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   7440
         Picture         =   "frmpruebaavance.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdDatos 
      Height          =   5295
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   16761024
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ecuación Potencial"
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   5520
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      Begin VB.TextBox Text1 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4080
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtab 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtb 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txta 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "B"
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
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "L"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "A"
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
         Left            =   2160
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "T="
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
         Left            =   1680
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Correlación"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Valor de b"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Valor de a"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   7815
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
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
      Left            =   960
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.LT)"
   End
   Begin MSComDlg.CommonDialog cdAccesar 
      Left            =   840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar el archivo a cargar"
      Filter          =   "(*.LT)"
   End
   Begin GraphLib.Graph Graph1 
      Height          =   2415
      Left            =   5520
      TabIndex        =   23
      Top             =   3360
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
   Begin VB.Label Label11 
      Caption         =   "Longitud"
      Height          =   255
      Left            =   9840
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Tiempo"
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   480
      TabIndex        =   22
      Top             =   1440
      Width           =   1935
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
      Height          =   1215
      Left            =   480
      TabIndex        =   21
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lbltitulo 
      Caption         =   "Prueba de Avance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   240
      Width           =   5415
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mcrear 
         Caption         =   "Guardar como"
         Shortcut        =   ^G
      End
      Begin VB.Menu maccesar 
         Caption         =   "Abrir Archivo"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mpasuelcli 
      Caption         =   "Parámetros Suelo -Clima"
      Begin VB.Menu mgesuel 
         Caption         =   "Generales Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcond 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meva 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu motrscal 
      Caption         =   "Otros Cálculos en Surcos"
      Begin VB.Menu masis 
         Caption         =   "Asistente de Diseño"
      End
      Begin VB.Menu msinf 
         Caption         =   "Surcos Infiltrómetros"
      End
   End
   Begin VB.Menu masismat 
      Caption         =   "Asistente Matémático"
      Begin VB.Menu h 
         Caption         =   "Hidráulica de Canales"
      End
      Begin VB.Menu mconv 
         Caption         =   "Convertidor de Unidades"
      End
   End
   Begin VB.Menu mm 
      Caption         =   "Menú Principal "
   End
End
Attribute VB_Name = "frmpruebaavance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valorl(1 To 300) As Double
Dim valort(1 To 300) As Double


Dim u As Integer
Dim a As Double
Dim b As Double
Rem funcion de la grid
Dim row, col, numero, num(0 To 1000)
Dim n As Integer, i, ii, punto

Private Sub grddatos_Click()
i = ""
punto = 0

End Sub


Private Sub Bcalcular_Click()
On Error GoTo mensaje
If u + 1 < 2 Then
   MsgBox "Ingrese al menos un par de valores", 64, "Imposible Calcular"
Exit Sub
End If
    n = u + 1
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
    txtA.text = Format(aa, "###0.0###")
    txtb.text = Format(b, "###0.0###")
    If lbltitulo.Caption <> "Regresión Potencial Simple" Then
        txtab.text = "T=" + Format(aa, "###0.0###") + "*L^" + Format(b, "###0.0###")
    Else
        txtab.text = "Y=" + Format(aa, "###0.0###") + "*X^" + Format(b, "###0.0###")
    End If
    Command1.Enabled = True
    
Exit Sub
mensaje:
   MsgBox "Ingrese datos adecuadamente", 64, "Imposible Calcular"


End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub bimprimir_Click()
Print Form

End Sub

Private Sub blimpiar_Click()
u = 0
txtA.text = ""
txtb.text = ""
txtab.text = ""
Text1.text = ""
 Graph1.Visible = False
       Label11.Visible = False
       Label10.Visible = False
grdDatos.Clear
grdDatos.Rows = 2
If lbltitulo.Caption = "Regresión Potencial Simple" Then
    grdDatos.TextMatrix(0, 0) = "X"
    grdDatos.TextMatrix(0, 1) = "Y"
    StatusBar1.Panels(1).text = "Introduzca los datos de la variable dependiente (X) y la independiente (Y), oprima el Botón Calcular para encontrar la regresión"
Else
    
    grdDatos.TextMatrix(0, 0) = "Longitud"
    grdDatos.TextMatrix(0, 1) = "Tiempo"
    StatusBar1.Panels(1).text = "Introduzca los datos de Tiempo y Longitud recopilados durante la prueba y oprima el Botón Calcular para encontrar la regresión"

End If
u = 0

End Sub


Private Sub Command1_Click()
On Error GoTo mensaje:
Rem gráfico de la curva
    numpuntos = n
    If numpuntos <= 1 Then
      temp = MsgBox("Para graficar se requiere por lo menos dos puntos", 64, "Imposible Graficar")
    Else
     Graph1.FontUse = 4
       Graph1.GraphType = 6
       Graph1.GraphStyle = 5
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
     End If
Exit Sub
mensaje:
   MsgBox "Imposible graficar", 16, " Error"
End Sub

Private Sub Form_Load()
u = 0

grdDatos.ColWidth(0) = 1300
grdDatos.ColWidth(1) = 1300
grdDatos.ColAlignment(0) = 4
grdDatos.ColAlignment(1) = 4
grdDatos.TextMatrix(0, 0) = "Longitud"
grdDatos.TextMatrix(0, 1) = "Tiempo"
StatusBar1.Panels(1).text = "Introduzca los datos de Tiempo y Longitud recopilados durante la prueba y oprima el Botón Calcular para encontrar la regresión"

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

Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub maccesar_Click()
On Error GoTo SinArchivo:
 cdAccesar.ShowOpen
 NombreArch = cdAccesar.FileName
 u = 0
 Open NombreArch For Random As #1 Len = Len(Pares)
 NumReg = LOF(1) \ Len(Pares)
 grdDatos.Rows = NumReg + 1
 For j% = 1 To NumReg
  Get #1, j%, Pares
  LL = Pares.L
  tT = Pares.T
  
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

Private Sub masis_Click()
frmriegosurcos.Show
End Sub

Private Sub mcond_Click()
frmconductividad.Show
End Sub

Private Sub mconv_Click()
frmconvertidor.Show
End Sub

Private Sub mcrear_Click()
 On Error GoTo SinArchivo
 ChDir App.Path
 cdCrear.ShowSave
 NombreArch = cdCrear.FileName
 ' Salvar archivo
 Open NombreArch For Random As #1 Len = Len(Pares)
 If (LOF(1) <> 0) Then
  Close #1
  Kill NombreArch
  Open NombreArch For Random As #1 Len = Len(Pares)
 End If
 For j% = 1 To (u + 1)
  Pares.L = valorl(j%)
  Pares.T = valort(j%)
  Put #1, j%, Pares
 Next j%
 Close

 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al salvar el archivo " & NombreArch
 End If
End Sub

Private Sub meva_Click()
frmETO.Show
End Sub

Private Sub mgesuel_Click()
frmgenerales.Show
End Sub

Private Sub mm_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub msinf_Click()
frmsurcosinfiltrometros.Show
End Sub

Private Sub mt_Click()
frmtextura.Show
End Sub
