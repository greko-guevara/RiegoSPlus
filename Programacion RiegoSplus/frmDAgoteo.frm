VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDAgoteo 
   Caption         =   "Diseño agronómico en Goteo"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11400
   Icon            =   "frmDAgoteo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11400
   Begin VB.Frame Frame6 
      Caption         =   "¿Calcular...?"
      Height          =   615
      Left            =   6600
      TabIndex        =   48
      Top             =   360
      Width           =   4695
      Begin VB.OptionButton Option3 
         Caption         =   "Lámina Bruta"
         Height          =   255
         Left            =   3240
         TabIndex        =   60
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Lámina Neta"
         Height          =   255
         Left            =   1920
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Consumo por planta"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1215
      Left            =   1080
      TabIndex        =   8
      Top             =   6480
      Width           =   6015
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   4080
         Picture         =   "frmDAgoteo.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2160
         MaskColor       =   &H000000FF&
         Picture         =   "frmDAgoteo.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   240
         Picture         =   "frmDAgoteo.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   240
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox Text1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5820
         TabIndex        =   64
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtFR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   61
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtLB 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5820
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtK 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtTR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5820
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtetrK 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   32
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtEf 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5820
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtRT 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Consumo por planta "
         Height          =   255
         Left            =   4200
         TabIndex        =   66
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "l/dia"
         Height          =   255
         Left            =   7140
         TabIndex        =   65
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Frecuencia de riego"
         Height          =   255
         Left            =   300
         TabIndex        =   62
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label30 
         Caption         =   "Lámina bruta"
         Height          =   255
         Left            =   4200
         TabIndex        =   47
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label34 
         Caption         =   "Caudal por hectarea"
         Height          =   255
         Left            =   300
         TabIndex        =   46
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "m3/h"
         Height          =   255
         Left            =   3540
         TabIndex        =   45
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Tiempo de riego"
         Height          =   255
         Left            =   4200
         TabIndex        =   44
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label h 
         Caption         =   "hrs"
         Height          =   255
         Left            =   7140
         TabIndex        =   43
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Precipitación "
         Height          =   255
         Left            =   300
         TabIndex        =   42
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "mm/h"
         Height          =   255
         Left            =   3540
         TabIndex        =   41
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label23 
         Caption         =   "%"
         Height          =   255
         Left            =   7140
         TabIndex        =   40
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "Eficiencia"
         Height          =   255
         Left            =   4200
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "Relación de transpiración"
         Height          =   255
         Left            =   300
         TabIndex        =   38
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label28 
         Caption         =   "días"
         Height          =   255
         Left            =   3540
         TabIndex        =   37
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "mm"
         Height          =   255
         Left            =   7140
         TabIndex        =   36
         Top             =   720
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   7815
      Width           =   11400
      _ExtentX        =   20108
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
            TextSave        =   "8/23/2008"
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos"
      ForeColor       =   &H00800000&
      Height          =   3015
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   11055
      Begin VB.CommandButton Command1 
         Caption         =   "Qué es esto?"
         Height          =   255
         Left            =   4560
         TabIndex        =   63
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   360
         TabIndex        =   51
         Top             =   1800
         Width           =   7095
         Begin VB.TextBox txtEg 
            Height          =   285
            Left            =   3360
            TabIndex        =   54
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox extEP 
            Height          =   285
            Left            =   5760
            TabIndex        =   53
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtNF 
            Height          =   285
            Left            =   1200
            TabIndex        =   52
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "m"
            Height          =   255
            Left            =   4320
            TabIndex        =   59
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label32 
            Caption         =   "Esp. goteros"
            Height          =   255
            Left            =   2400
            TabIndex        =   58
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label33 
            Caption         =   "m"
            Height          =   255
            Left            =   6720
            TabIndex        =   57
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label35 
            Caption         =   "Esp. manguera"
            Height          =   375
            Left            =   4920
            TabIndex        =   56
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label39 
            Caption         =   "N° de filas"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.ComboBox ctextura 
         Height          =   315
         Left            =   8640
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cRaices 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Gotero"
         Height          =   975
         Left            =   6240
         TabIndex        =   25
         Top             =   240
         Width           =   4095
         Begin VB.TextBox txtQG 
            Height          =   285
            Left            =   2280
            TabIndex        =   4
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCU 
            Height          =   285
            Left            =   2280
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "%"
            Height          =   255
            Left            =   3600
            TabIndex        =   29
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label24 
            Caption         =   "Caudal del gotero"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "lts/h"
            Height          =   255
            Left            =   3600
            TabIndex        =   27
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "Coeficiente de uniformidad"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.TextBox txtLN 
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtMP 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPAR 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Text            =   "100"
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton bCalcular 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   8760
         Picture         =   "frmDAgoteo.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Textura"
         Height          =   255
         Left            =   6360
         TabIndex        =   24
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "l/día"
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Consumo por planta"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "m2"
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Marco de plantación"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Profundidad de raices"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblunidades 
         Caption         =   "mts"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label36 
         Caption         =   "%"
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label37 
         Caption         =   "Porcentaje de área humedecida"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2625
      Left            =   8280
      Picture         =   "frmDAgoteo.frx":29F2
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   3435
   End
   Begin VB.Label Label17 
      Caption         =   "Diseño agronómico en Riego por Goteo"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   360
      Width           =   6255
   End
   Begin VB.Menu qpqpqpp 
      Caption         =   "Parámetros Suelo- clima"
      Begin VB.Menu mgesu 
         Caption         =   "Generales del suelo"
      End
      Begin VB.Menu mtex 
         Caption         =   "Textura"
      End
      Begin VB.Menu mconhi 
         Caption         =   "Conductividad hidráulica"
      End
      Begin VB.Menu meva 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu owowo 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu mlat 
         Caption         =   "Cálculos en el lateral"
      End
      Begin VB.Menu mprin 
         Caption         =   "Cálculos en la principal"
      End
      Begin VB.Menu mse 
         Caption         =   "Selección de bomba"
      End
      Begin VB.Menu mconbi 
         Caption         =   "Combinación de diámetros"
      End
   End
   Begin VB.Menu tututuutut 
      Caption         =   "Menú principal"
   End
End
Attribute VB_Name = "frmDAgoteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bcalcular_Click()
On Error GoTo mensaje:

mp = Val(txtMP.text)
ln = Val(txtLn.text)
PAR = Val(txtpar.text)
cu = Val(txtCU.text)
qg = Val(txtQG.text)
rai = Val(cRaices.ListIndex)
tex = Val(CTextura.ListIndex)
eg = Val(txteg.text)
ep = Val(extEP.text)
nf = Val(txtnf.text)

If mp = 0 Then
    MsgBox "Ingrese el valor del Marco de Plantación", 64, "Riego por goteo"
    txtMP.SetFocus
    Exit Sub
End If
If ln = 0 Then
    If Option1.Value = True Then
    MsgBox "Ingrese el consumo por planta", 64, "Riego por goteo"
    Else
    If Option2.Value = True Then
    MsgBox "Ingrese la lámina neta", 64, "Riego por goteo"
    Else
    MsgBox "Ingrese la lámina bruta", 64, "Riego por goteo"
    End If
    End If
    txtLn.SetFocus
    Exit Sub
End If
    
If PAR = 0 Then
    MsgBox "Ingrese el valor del Porcentaje de área humedecida", 64, "Riego por goteo"
    txtpar.SetFocus
    Exit Sub
End If
If cu = 0 Then
    MsgBox "Ingrese el valor del coeficiente de uniformidad del gotero", 64, "Riego por goteo"
   txtCU.SetFocus
    Exit Sub
End If
If qg = 0 Then
    MsgBox "Ingrese el valor del caudal del gotero", 64, "Riego por goteo"
    txtQG.SetFocus
    Exit Sub
End If
If eg = 0 Then
    If Option4.Value = True Then
    MsgBox "Ingrese el espaciamiento entre goteros", 64, "Riego por goteo"
    txteg.SetFocus
    Exit Sub
    End If
End If
If ep = 0 Then
    If Option4.Value = True Then
    MsgBox "Ingrese el espaciamiento entre plantas", 64, "Riego por goteo"
    extEP.SetFocus
    Exit Sub
    End If
End If
If nf = 0 Then
    MsgBox "Ingrese el número de líneas de gotero por cada línea de plantas ", 64, "Riego por microaspersión"
   txtnf.SetFocus
    Exit Sub
End If
If rai = -1 Then
    MsgBox "Seleccione la profundidad de raices", 64, "Riego por goteo"
    cRaices.SetFocus
    Exit Sub
End If
If tex = -1 Then
    MsgBox "Seleccione la textura", 64, "Riego por microaspersión"
   CTextura.SetFocus
    Exit Sub
End If
'**

'RELACION TRANSPIRACIION
Select Case tex
Case 0
 Select Case rai
    Case 0
    rt = 0.9
    Case 1
    rt = 0.9
    Case 2
    rt = 95
 End Select
Case 1
 Select Case rai
    Case 0
    rt = 0.9
    Case 1
    rt = 0.95
    Case 2
    rt = 1
 End Select
Case 2
 Select Case rai
    Case 0
    rt = 0.95
    Case 1
    rt = 1
    Case 2
    rt = 1
 End Select
Case 3
 Select Case rai
    Case 0
    rt = 1
    Case 1
    rt = 1
    Case 2
    rt = 1
 End Select
End Select

ef = cu / 100 * rt



'**
'frecuencia lamina
If Option1.Value = True Then
    prec = mp / ep / eg * nf * qg
    tr = ln / prec
    lb = ln / mp
    Label10.Caption = "l/h"
    Text1.text = ln
Else
    If Option2.Value = True Then
        prec = qg / (ep * eg * PAR / 100)
        lb = ln / ef
        tr = lb / prec
        Label10.Caption = "mm/h"
        Text1.text = Format(lb * mp, "##0.00##")
    Else
        prec = qg / (ep * eg * PAR / 100)
        tr = ln / prec
        lb = ln
        Label10.Caption = "mm/h"
        Text1.text = Format(lb * mp, "##0.00##")
    End If
End If






'caudal por hectarea
qhect = 100 / (eg) * 100 / ep * nf * qg / 1000


txtK.text = Format(qhect, "##0.00##")
txtetrK.text = Format(prec, "##0.00##")

txtRT.text = Format(rt, "##0.00##")
txtLB.text = Format(lb, "##0.00##")
txtEf.text = Format(ef * 100, "##0.00##")
txtTR.text = Format(tr, "##0.00##")

Frame2.Visible = True
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"

End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
txtK.text = ""
txtetrK.text = ""
txtFR.text = ""
txtRT.text = ""
txtLB.text = ""
txtEf.text = ""
txtTR.text = ""
txtMP.text = ""
txtLn.text = ""
txtpar.text = ""
txtCU.text = ""
txtQG.text = ""
txteg.text = ""
extEP.text = ""
txtnf.text = ""
cRaices.text = ""
CTextura.text = ""
Frame2.Visible = False
End Sub


Private Sub Command1_Click()
Dialog.Show
End Sub


Private Sub Form_Load()
With cRaices
    .AddItem "<de 0.75"
    .AddItem "0.75 a 1.50"
    .AddItem ">de 1.50"
End With
With CTextura
    .AddItem "Gruesa"
    .AddItem "Media"
    .AddItem "Fina"
    .AddItem "Muy fina"
End With
StatusBar1.Panels(1).text = "Ingrese los datos básicos y oprima el botón de Calcular "
End Sub


Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    Frame1.Visible = True
    Frame4.Visible = False
    StatusBar1.Panels(1).text = "Ingrese los datos básicos y oprima el botón de Calcular "
    Case 2
    Frame4.Visible = True
    Frame1.Visible = False
    StatusBar1.Panels(1).text = ""
End Select
End Sub



Private Sub mconbi_Click()
frmcombDia.Show
End Sub

Private Sub mconhi_Click()
frmconductividad.Show
End Sub

Private Sub meva_Click()
frmETO.Show
End Sub

Private Sub mgesu_Click()
frmgenerales.Show
End Sub

Private Sub mlat_Click()
FrmHLaterales.Show
End Sub

Private Sub mprin_Click()
frmHprincipal.Show
End Sub

Private Sub mse_Click()
frmbomba.Show
End Sub

Private Sub mtex_Click()
frmtextura.Show
End Sub

Private Sub Option1_Click()
Command1.Visible = False
Label37.Visible = False
Label36.Visible = False
txtpar.Visible = False
txtpar.text = 100

Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
Option3.ForeColor = &H80000012
Label6.Caption = "Frecuencia de riego"
Label2.Caption = "Consumo por planta"
Label3.Caption = "l/día"
End Sub

Private Sub Option2_Click()
Command1.Visible = True
Label37.Visible = True
Label36.Visible = True
txtpar.Visible = True
txtpar.text = 100

Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
Option3.ForeColor = &H80000012
Label2.Caption = "Lámina Neta"
Label3.Caption = "mm"
Label6.Caption = "lámina neta"
End Sub
Private Sub Option3_Click()
Command1.Visible = True
Label37.Visible = True
Label36.Visible = True
txtpar.Visible = True
txtpar.text = 100

Option3.ForeColor = &HC0&
Option1.ForeColor = &H80000012
Option2.ForeColor = &H80000012
Label2.Caption = "Lámina Bruta"
Label3.Caption = "mm"
Label6.Caption = "lámina neta"

End Sub



Private Sub tututuutut_Click()
Unload Me
frmGeneral.Show
End Sub
