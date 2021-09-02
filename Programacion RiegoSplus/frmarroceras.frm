VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmarroceras 
   Caption         =   "Melgas para Arroz"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmarroceras.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      Height          =   2055
      Left            =   6360
      TabIndex        =   66
      Top             =   5520
      Width           =   4575
      Begin VB.CommandButton Bevaluar 
         Caption         =   "&Evaluar"
         Height          =   735
         Left            =   480
         Picture         =   "frmarroceras.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2520
         Picture         =   "frmarroceras.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   480
         Picture         =   "frmarroceras.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   2520
         Picture         =   "frmarroceras.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Etapa de Reposición"
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   6360
      TabIndex        =   59
      Top             =   4320
      Width           =   4575
      Begin VB.TextBox txtqt3 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   61
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtqmelga3 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   60
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label44 
         Caption         =   "Caudal total necesario"
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label45 
         Caption         =   "m3/s"
         Height          =   255
         Left            =   3840
         TabIndex        =   64
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label46 
         Caption         =   "Caudal por melga"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label47 
         Caption         =   "m3/s"
         Height          =   255
         Left            =   1800
         TabIndex        =   62
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Etepa de Inundación"
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   6360
      TabIndex        =   49
      Top             =   2640
      Width           =   4575
      Begin VB.TextBox txtq2 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   52
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtqmelga2 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   51
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtTinund 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Caudal total necesario"
         Height          =   255
         Left            =   2400
         TabIndex        =   58
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label37 
         Caption         =   "m3/s"
         Height          =   255
         Left            =   3960
         TabIndex        =   57
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Caudal por melga"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label41 
         Caption         =   "m3/s"
         Height          =   255
         Left            =   1800
         TabIndex        =   55
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "días"
         Height          =   255
         Left            =   1800
         TabIndex        =   54
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label43 
         Caption         =   "Tiempo de infiltración"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Etapa de Mojado"
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   6360
      TabIndex        =   36
      Top             =   960
      Width           =   4575
      Begin VB.TextBox txtqmelga 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtqt 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   39
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtti 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtTa 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   "m3/s"
         Height          =   255
         Left            =   1800
         TabIndex        =   48
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label28 
         Caption         =   "Caudal por melga"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "m3/s"
         Height          =   255
         Left            =   3960
         TabIndex        =   46
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label26 
         Caption         =   "Caudal total necesario"
         Height          =   255
         Left            =   2400
         TabIndex        =   45
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "Tiempo de infiltración"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "min"
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   600
         Width           =   255
      End
      Begin VB.Label sfsf 
         Caption         =   "Tiempo de avance"
         Height          =   255
         Left            =   2400
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "min"
         Height          =   255
         Left            =   3960
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para el diseño"
      ForeColor       =   &H00800000&
      Height          =   4575
      Left            =   840
      TabIndex        =   14
      Top             =   960
      Width           =   4575
      Begin VB.TextBox txtZ 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtper 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtETR 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtLN 
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtAgot 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtNM 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txta 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtb 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4560
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label36 
         Caption         =   "cm"
         Height          =   255
         Left            =   3840
         TabIndex        =   34
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label35 
         Caption         =   "Diferencia entre bordos (z)"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Lámina permanente (H)"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "mm/día"
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "%"
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "has"
         Height          =   255
         Left            =   3840
         TabIndex        =   29
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "Area de la melga"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "mm/dia"
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "mm"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Agotamiento"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Evapotranspiración real"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Lámina Neta "
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Número de melgas"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Icum=mm     y    t=min"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label16 
         Caption         =   "A"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   "B"
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label32 
         Caption         =   "Icum= A x B^b"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label33 
         Caption         =   "Percolación promedio"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label34 
         Caption         =   "cm"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   1800
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   35
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
            TextSave        =   "07/07/2005"
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   840
      Picture         =   "frmarroceras.frx":29F2
      Top             =   5640
      Width           =   4560
   End
   Begin VB.Label Label10 
      Caption         =   "Melgas para Arroz"
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
      Left            =   1200
      TabIndex        =   28
      Top             =   240
      Width           =   2775
   End
   Begin VB.Menu mgsc 
      Caption         =   "Generales Suelo - Clima"
      Begin VB.Menu mgs 
         Caption         =   "Generales Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mconduct 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meto 
         Caption         =   "Evapotranspiración "
      End
   End
   Begin VB.Menu motros 
      Caption         =   "Otros Cálculos Melgas"
      Begin VB.Menu mmcp 
         Caption         =   "Melgas con Pendiente"
      End
      Begin VB.Menu mprueba 
         Caption         =   "Prueba de Campo en melgas con pendiente"
      End
      Begin VB.Menu mmsp 
         Caption         =   "Melgas Sin Pendiente"
      End
   End
   Begin VB.Menu mam 
      Caption         =   "Asistente Matemático"
      Begin VB.Menu mconvert 
         Caption         =   "Conversión de Unidades"
      End
      Begin VB.Menu h 
         Caption         =   "Hidráulica de Canales"
      End
      Begin VB.Menu mrefpot 
         Caption         =   "Regresión Potencial Simple"
      End
   End
   Begin VB.Menu mmp 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmarroceras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bevaluar_Click()
On Error GoTo mensaje
ln = Val(txtLN.Text)
agot = Val(txtAgot.Text)
area = Val(txtArea.Text)
nm = Val(txtNM.Text)
h = Val(txtH.Text)
z = Val(txtZ.Text)
a = Val(txta.Text)
b = Val(txtb.Text)
etr = Val(txtETR.Text)
per = Val(txtper.Text)

If ln = 0 Then
MsgBox "Ingrese el valor de la lámina neta", 64, "Arroceras"
txtLN.SetFocus
Exit Sub
End If
If agot = 0 Then
MsgBox "Ingrese el valor de agotamiento", 64, "Arroceras"
txtAgot.SetFocus
Exit Sub
End If
If area = 0 Then
MsgBox "Ingrese el valor del área", 64, "Arroceras"
txtArea.SetFocus
Exit Sub
End If
If nm = 0 Then
MsgBox "Ingrese el valor del número de melgas", 64, "Arroceras"
txtNM.SetFocus
Exit Sub
End If
If h = 0 Then
MsgBox "Ingrese el valor de la lámina permanente", 64, "Arroceras"
txtH.SetFocus
Exit Sub
End If
If z = 0 Then
MsgBox "Ingrese el valor de la diferencia entre bordos", 64, "Arroceras"
txtZ.SetFocus
Exit Sub
End If
If a = 0 Then
MsgBox "Ingrese los parámetros de la ecuación de infiltración", 64, "Arroceras"
txta.SetFocus
Exit Sub
End If
If b = 0 Then
MsgBox "Ingrese los parámetros de la ecuación de infiltración", 64, "Arroceras"
txtb.SetFocus
Exit Sub
End If
If etr = 0 Then
MsgBox "Ingrese el valor de la evapotranspiración real", 64, "Arroceras"
txtETR.SetFocus
Exit Sub
End If
If per = 0 Then
MsgBox "Ingrese el valor de la percolación promedio", 64, "Arroceras"
txtper.SetFocus
Exit Sub
End If

Rem -*-***-*-*-*-*-*-*-etapa de mojado *******************************

Rem calculo de tiempos
ti = (ln / a) ^ (1 / b)
tA = ti / 4
txtti.Text = Format(ti, "##0.00##")
txtTa.Text = Format(tA, "##0.00##")


Rem caudal

d = z / 2
bb = b - 1
aa = a * b
R1 = (1 - bb) / 2
f = (bb - R1 * bb + 2) / (1 + R1)
icum = f * aa * tA ^ (b) / ((b) * (bb + 2))
qo = (d / 100 + icum / 1000) * 10000 / (60 * tA)
qmelga1 = area * qo
qt1 = qmelga1 * nm

txtqmelga.Text = Format(qmelga1, "##0.00##")
txtqt.Text = Format(qt1, "##0.00##")

Rem -*-***-*-*-*-*-*-*-etapa de inundacion *******************************

lsat = 2 * ln / (agot * 10)
d1 = d / 100
tinund = ln / (agot * 2 * etr / 100)
lper = per * tinund / 1000
qmelga2 = (lsat + d1 + h / 100 + lper) * area * 10000 / (tinund * 86400)
qt2 = qmelga2 * nm

txtqmelga2.Text = Format(qmelga2, "##0.00##")
txtq2.Text = Format(qt2, "##0.00##")
txtTinund.Text = Format(tinund, "##0.00##")

Rem -*-***-*-*-*-*-*-*-etapa de reposición *******************************

qmelga3 = (per + etr) * area / 8640
qt3 = qmelga3 * nm

txtqmelga3.Text = Format(qmelga3, "##0.00##")
txtqt3.Text = Format(qt3, "##0.00##")
Exit Sub
mensaje:
MsgBox "ingrese valores adecuados", 64, "Arroceras"
End Sub

Private Sub bfinailizar_Click()
Unload Me
End Sub



Private Sub Command2_Click()
txtLN.Text = ""
txtAgot.Text = ""
txtArea.Text = ""
txtNM.Text = ""
txtH.Text = ""
txtZ.Text = ""
txta.Text = ""
txtb.Text = ""
txtETR.Text = ""
txtper.Text = ""
txtti.Text = ""
txtTa.Text = ""
txtqt.Text = ""
txtqmelga.Text = ""
txtqmelga2.Text = ""
txtqmelga3.Text = ""
txtq2.Text = ""

txtTinund.Text = ""
txtLN.SetFocus
End Sub

Private Sub Command3_Click()
Print Form
End Sub

Private Sub Command4_Click()
Unload Me: frmGeneral.Show
End Sub

Private Sub Form_Load()
StatusBar1.Panels(1).Text = "Digite los Datos Básicos de Diseño y oprima el botón de Evaluar para realizar los cálculos"
End Sub


Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub mconduct_Click()
frmconductividad.Show
End Sub

Private Sub mconvert_Click()
frmconvertidor.Show
End Sub

Private Sub meto_Click()
frmETO.Show
End Sub

Private Sub mgs_Click()
frmgenerales.Show
End Sub

Private Sub mmcp_Click()
frmmelgaspendiente.Show
End Sub

Private Sub mmp_Click()
frmGeneral.Show
End Sub

Private Sub mmsp_Click()
frmmelgassinpendiente.Show
End Sub

Private Sub mprueba_Click()
frmpruebaavancerecesion.Show
End Sub

Private Sub mrefpot_Click()
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



