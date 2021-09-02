VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmpruebaavancerecesion 
   Caption         =   "Pruebas de avance recesión"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmpruebaavancerecesion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      Caption         =   "Datos Fijos"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   840
      TabIndex        =   7
      Top             =   960
      Width           =   3855
      Begin VB.TextBox txtln 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txta 
         BackColor       =   &H80000004&
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtb 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "li=cm/h y t=min"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Lámina neta"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "cm"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "li="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label32 
         Caption         =   "li = a x t^b"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2295
      Left            =   840
      TabIndex        =   29
      Top             =   5160
      Width           =   3855
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2040
         Picture         =   "frmpruebaavancerecesion.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   240
         Picture         =   "frmpruebaavancerecesion.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   2040
         Picture         =   "frmpruebaavancerecesion.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   240
         Picture         =   "frmpruebaavancerecesion.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   840
      TabIndex        =   16
      Top             =   3000
      Width           =   3855
      Begin VB.TextBox txtLP 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtEF 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtTI 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "cm"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Lámina percolada prueba actual"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "%"
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Eficiencia prueba actual"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "min"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Tiempo de contacto mínimo"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   28
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
   Begin MSComDlg.CommonDialog cdCrear 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "(*.DTaTc)"
   End
   Begin MSComDlg.CommonDialog cdAccesar 
      Left            =   0
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar el archivo a cargar"
      Filter          =   "(*.DTaTc)"
   End
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   5160
      TabIndex        =   6
      Top             =   960
      Width           =   6495
      Begin VB.TextBox txtlapr 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txttcpr 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Height          =   5055
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   6015
         Begin VB.CommandButton Command2 
            Caption         =   "&Graficar"
            Height          =   735
            Left            =   4200
            Picture         =   "frmpruebaavancerecesion.frx":29F2
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid grdD 
            Height          =   3495
            Left            =   0
            TabIndex        =   40
            Top             =   1320
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            GridColor       =   16761024
         End
         Begin VB.Label Label11 
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
            Left            =   360
            TabIndex        =   42
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label12 
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
            Left            =   2400
            TabIndex        =   41
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5055
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   6015
         Begin VB.CommandButton Command3 
            Caption         =   "Regresar"
            Height          =   735
            Left            =   4200
            Picture         =   "frmpruebaavancerecesion.frx":315C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   480
            Width           =   1575
         End
         Begin GraphLib.Graph Graph1 
            Height          =   3255
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   5775
            _Version        =   65536
            _ExtentX        =   10186
            _ExtentY        =   5741
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
         Begin VB.Label Label17 
            Caption         =   "Tiempo"
            Height          =   255
            Left            =   720
            TabIndex        =   37
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Longitud"
            Height          =   255
            Left            =   4560
            TabIndex        =   36
            Top             =   4560
            Width           =   1215
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Lámina acumulada promedio"
         Height          =   375
         Left            =   3480
         TabIndex        =   32
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "cm"
         Height          =   255
         Left            =   5280
         TabIndex        =   31
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "minutos"
         Height          =   375
         Left            =   2280
         TabIndex        =   30
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Tiempo de contacto promedio "
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   5760
         Width           =   2295
      End
   End
   Begin VB.Label Label16 
      Caption         =   "Pruebas de avance- recesión "
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
      Left            =   480
      TabIndex        =   27
      Top             =   240
      Width           =   5295
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mguardar 
         Caption         =   "Guardar como"
         Shortcut        =   ^G
      End
      Begin VB.Menu mabriuri 
         Caption         =   "Abrir proyecto"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mpsc 
      Caption         =   "Parámetros Suelo - Clima"
      Begin VB.Menu mps 
         Caption         =   "Parámetros del Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcond 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meto 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu motrs 
      Caption         =   "Otros Cálculos en Melgas"
      Begin VB.Menu mmcp 
         Caption         =   "Melgas con Pendiente"
      End
      Begin VB.Menu mmsp 
         Caption         =   "Melgas sin Pendiente"
      End
      Begin VB.Menu marr 
         Caption         =   "Arroceras"
      End
   End
   Begin VB.Menu mass 
      Caption         =   "Asistente Matemático"
      Begin VB.Menu mconvert 
         Caption         =   "Convertidor de Unidades"
      End
      Begin VB.Menu h 
         Caption         =   "Hidráulica de Canales"
      End
      Begin VB.Menu mreg 
         Caption         =   "Regresión Potencial Simple"
      End
   End
   Begin VB.Menu mmp 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmpruebaavancerecesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem funcion de la grid
Dim row, col, numero, num(0 To 1000)
Dim n As Integer, i, ii, punto
Dim u As Integer
Dim valord(1 To 300) As Double
Dim valorta(1 To 300) As Double
Dim valortr(1 To 300) As Double
Dim kmkm(1 To 3) As Double


Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtLn.text = ""
txtA.text = ""
txtb.text = ""
txtlapr.text = ""
txttcpr.text = ""
Command2.Visible = False
Frame5.Visible = False
Frame6.Visible = True
txtTI.text = ""
txtEf.text = ""
txtLP.text = ""
u = 0
grdD.Clear
With grdD
    .Cols = 3
    .Rows = 2
    .TextMatrix(0, 0) = "Distancia"
    .TextMatrix(0, 1) = "T avance"
    .TextMatrix(0, 2) = "T recesión"
End With

End Sub


Private Sub Command1_Click()
On Error GoTo mensaje

a = Val(txtA.text)
b = Val(txtb.text)
ln = Val(txtLn.text)
n = u + 1
    
If n < 2 Then
MsgBox "Ingrese al menos un par de datos de la prueba de avance", 64, "Imposible Calcular"
grdD.SetFocus
Exit Sub
End If
If ln = 0 Then
MsgBox "Ingrese el valor de la lámina neta", 64, "Prueba de avance en melgas"
txtLn.SetFocus
Exit Sub
End If
If a = 0 Then
MsgBox "Ingrese valores de la ecuación de infitración instantánea", 64, "Prueba de avance en melgas"
txtA.SetFocus
Exit Sub
End If
If b = 0 Then
MsgBox "Ingrese valores de la ecuación de infitración instantánea", 64, "Prueba de avance en melgas"
txtb.SetFocus
Exit Sub
End If

For j% = 1 To n
valord(j%) = grdD.TextMatrix(j%, 0)
valorta(j%) = grdD.TextMatrix(j%, 1)
valortr(j%) = grdD.TextMatrix(j%, 2)
Next j%

With grdD
.Cols = 5
.ColAlignment(3) = 4
.ColAlignment(4) = 4
.ColWidth(3) = 1150
.ColWidth(4) = 1150
End With
grdD.TextMatrix(0, 3) = "T Contacto"
grdD.TextMatrix(0, 4) = "Lámina Acum"
    
    aa = a / ((b + 1) * 60)
    bb = b + 1
    
    stc = 0
    sla = 0
    For j% = 1 To n
        TC = Val(grdD.TextMatrix(j%, 2)) - Val(grdD.TextMatrix(j%, 1))
        If TC >= 0 Then
            la = aa * TC ^ bb
            TC = Format(TC, "##0.00#")
            la = Format(la, "##0.00#")
            grdD.TextMatrix(j%, 3) = TC
            grdD.TextMatrix(j%, 4) = la
            
            stc = stc + Val(grdD.TextMatrix(j%, 3))
            sla = sla + Val(grdD.TextMatrix(j%, 4))
        Else
        MsgBox "El tiempo recesión no puede ser mayor al de avance", 16, "Prueba de avance en Melgas"
        End If
        Next j%
        tcp = stc / n
        lap = sla / n
            
        txtlapr.text = Format(lap, "##0.00#")
        txttcpr.text = Format(tcp, "##0.00#")
            
    
    
    ti = (ln / aa) ^ (1 / bb)
    
    ef = ln / lap * 100
    LP = lap - ln
    
    txtTI.text = Format(ti, "##0.00#")
    If ef > 100 Then
        txtEf.text = "lámina insuficiente"
    Else
        txtEf.text = Format(ef, "##0.00#")
    End If
    txtLP.text = Format(LP, "##0.00#")
 Command2.Visible = True
 
Exit Sub
mensaje:
  MsgBox "Ingrese adecuadamente los datos", 64, "Prueba de Avance Recesión"

End Sub

Private Sub Command2_Click()
'On Error GoTo mensaje:
Frame6.Visible = False
Frame5.Visible = True
    numpuntos = n
    If numpuntos <= 1 Then
      temp = MsgBox("Para graficar se requiere por lo menos dos puntos", 64, "Imposible Graficar")
    Else
     Graph1.FontUse = 4
       Graph1.GraphType = 6
       Graph1.GraphStyle = 0
       Graph1.BorderStyle = 1
       Graph1.Visible = True
       Frame6.Visible = False
       Frame5.Visible = True
       Graph1.AutoInc = 1
       Graph1.NumPoints = numpuntos
       Graph1.NumSets = 2
        With Graph1
        For j% = 1 To .NumSets
        .ThisSet = j%
         If j% = 1 Then
           .LegendText = "Avance"
         Else
           .LegendText = "Recesión"
         End If
        Next j%
       
 'Ingresar valores del eje X
   
       For j% = 1 To numpuntos
          .XPosData = valord(j%)
       Next j%
       
 'Ingresar valores del eje Y
      For j% = 1 To 2 * numpuntos
       Select Case j%
        Case Is <= numpuntos
          .GraphData = valorta(j%)
        Case Is <= 2 * numpuntos
          .GraphData = valortr(j% - numpuntos)
       End Select
      Next j%
        .DrawMode = 2
     
    End With
     End If
Exit Sub
mensaje:
  MsgBox "Ingrese adecuadamente los datos", 64, "Prueba de Avance Recesión"
End Sub

Private Sub Command3_Click()
Frame5.Visible = False
Frame6.Visible = True

End Sub

Private Sub Form_Load()
StatusBar1.Panels(1).text = "Favor digitar los datos fijos e introducir los datos recopilados durante la prueba, luego oprima el Botón de Calcular  "
u = 0
With grdD
    .ColAlignment(0) = 4
    .ColAlignment(1) = 4
    .ColAlignment(2) = 4
    .ColWidth(0) = 1150
    .ColWidth(1) = 1150
    .ColWidth(2) = 1150
    .TextMatrix(0, 0) = "Distancia"
    .TextMatrix(0, 1) = "T avance"
    .TextMatrix(0, 2) = "T recesión"
    
End With
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


'enter
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

Rem pruevas grid1.TextMatrix(numero, 6) = num(numero)

Rem grdDatos.Text = KeyAscii
col = grdD.col
row = grdD.row

End Sub



Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub mabriuri_Click()
On Error GoTo SinArchivo:
 cdAccesar.ShowOpen
 NombreArch = cdAccesar.FileName
 u = 0
 Open NombreArch For Random As #1 Len = Len(TrioDTaTR)
 NumReg = LOF(1) \ Len(TrioDTaTR)
 grdD.Rows = NumReg + 1
 For j% = 1 To NumReg
  Get #1, j%, TrioDTaTR
  txtLn.text = TrioDTaTR.kln
  txtA.text = TrioDTaTR.ka
  txtb.text = TrioDTaTR.kb
  dd = TrioDTaTR.d
  tta = TrioDTaTR.ta
  ttr = TrioDTaTR.tr
  With grdD
   .TextMatrix(j%, 0) = dd
   .TextMatrix(j%, 1) = tta
   .TextMatrix(j%, 2) = ttr
  End With
 Next j%
 Close
 u = NumReg - 1
 Exit Sub
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al abrir el archivo " & NombreArch
 End If
End Sub

Private Sub marr_Click()
frmarroceras.Show
End Sub

Private Sub mcond_Click()
frmconductividad.Show
End Sub

Private Sub mconvert_Click()
frmconvertidor.Show
End Sub

Private Sub meto_Click()
frmETO.Show
End Sub

Private Sub mguardar_Click()
On Error GoTo SinArchivo
 ChDir App.Path
 cdCrear.ShowSave
 NombreArch = cdCrear.FileName
 
 ' Salvar archivo
 Open NombreArch For Random As #1 Len = Len(TrioDTaTR)
 If (LOF(1) <> 0) Then
  Close #1
  Kill NombreArch
  Open NombreArch For Random As #1 Len = Len(TrioDTaTR)
 End If
 For j% = 1 To (u + 1)
  valord(j%) = Val(grdD.TextMatrix(j%, 0))
  valorta(j%) = Val(grdD.TextMatrix(j%, 1))
  valortr(j%) = Val(grdD.TextMatrix(j%, 2))
  TrioDTaTR.d = valord(j%)
  TrioDTaTR.ta = valorta(j%)
  TrioDTaTR.tr = valortr(j%)
  kmkm(1) = Val(txtLn.text)
  TrioDTaTR.kln = kmkm(1)
  kmkm(2) = Val(txtA.text)
  TrioDTaTR.ka = kmkm(2)
  kmkm(3) = Val(txtb.text)
  TrioDTaTR.kb = kmkm(3)
  Put #1, j%, TrioDTaTR
 Next j%
 Close

 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al salvar el archivo " & NombreArch
 End If
End Sub

Private Sub mmcp_Click()
frmmelgaspendiente.Show
End Sub

Private Sub mmp_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub mmsp_Click()
frmmelgassinpendiente.Show
End Sub

Private Sub mps_Click()
frmgenerales.Show
End Sub

Private Sub mreg_Click()
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

Private Sub mt_Click()
frmtextura.Show
End Sub
