VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmelgassinpendiente 
   Caption         =   "Melgas Sin Pendiente ni Salida de Agua"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmmelgassinpendiente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   1080
      TabIndex        =   57
      Top             =   5520
      Width           =   4095
      Begin VB.CommandButton Bevaluar 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   240
         Picture         =   "frmmelgassinpendiente.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2160
         Picture         =   "frmmelgassinpendiente.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   240
         Picture         =   "frmmelgassinpendiente.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   2160
         Picture         =   "frmmelgassinpendiente.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos conociendo:"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   1920
      TabIndex        =   55
      Top             =   600
      Width           =   7695
      Begin VB.OptionButton Option2 
         Caption         =   "Caudal Unitario por metro de melga"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tirante medio de la melga (D)"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.Line Line3 
         X1              =   2400
         X2              =   2640
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   5880
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtava 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   59
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txttinf 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   58
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox TxtDo 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtLre 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   44
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtTapl 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtY 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtLB 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtLB1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtLN1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "min"
         Height          =   255
         Left            =   3240
         TabIndex        =   63
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "min"
         Height          =   255
         Left            =   4440
         TabIndex        =   62
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "T.infiltración"
         Height          =   255
         Left            =   2400
         TabIndex        =   61
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "T. avance"
         Height          =   255
         Left            =   3600
         TabIndex        =   60
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label text 
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   2280
         TabIndex        =   52
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "cm"
         Height          =   255
         Left            =   1440
         TabIndex        =   51
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Tirante Máximo al inicio "
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5160
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label24 
         Caption         =   "mts"
         Height          =   255
         Left            =   3720
         TabIndex        =   48
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Longitud Recomenda "
         Height          =   255
         Left            =   2400
         TabIndex        =   47
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label38 
         Caption         =   "min"
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label39 
         Caption         =   "Tiempo de Aplicación"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblunidades1 
         Height          =   255
         Left            =   1680
         TabIndex        =   42
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lbletiqueta1 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "mm"
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Lámina Bruta"
         Height          =   255
         Left            =   2400
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "día"
         Height          =   255
         Left            =   1440
         TabIndex        =   38
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "Frecuencia de Riego"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Lámina Bruta corregida"
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "mm"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "Lámina Neta corregida"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "mm"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   1200
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para el diseño"
      ForeColor       =   &H00800000&
      Height          =   3975
      Left            =   1080
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   64
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtb 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txta 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtEFAP 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtQdis 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtLN 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtETR 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Coeficente de Manning n (Puede ser digitado o seleccionado)"
         Height          =   495
         Left            =   240
         TabIndex        =   65
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label34 
         Caption         =   "cm"
         Height          =   255
         Left            =   3600
         TabIndex        =   54
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label33 
         Caption         =   "Altura de Camellones"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "Icum= A x B^b"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "B"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "A"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "Icum=cm/h y t=min"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "%"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Eficiencia de Aplicación"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lbletiqueta 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Lámina Neta "
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Evapotranspiración real"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblunidades 
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "mm"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "mm/dia"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   56
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Left            =   6000
      Picture         =   "frmmelgassinpendiente.frx":29F2
      Top             =   5400
      Width           =   4560
   End
   Begin VB.Label Label10 
      Caption         =   "Diseño de Riego por Melgas sin Pendiente y sin Salida de Agua"
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
      TabIndex        =   20
      Top             =   120
      Width           =   8175
   End
   Begin VB.Menu psc 
      Caption         =   "Parámetros del Suelo - Clima"
      Begin VB.Menu ps 
         Caption         =   "Parámetros Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcionduct 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meto 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu motros 
      Caption         =   "Otros Cálculos con Melgas"
      Begin VB.Menu mmcp 
         Caption         =   "Melgas con pendiente"
      End
      Begin VB.Menu mpcmcp 
         Caption         =   "Pruebas de Campo Melgas con pendiente"
      End
      Begin VB.Menu marro 
         Caption         =   "Arroceras"
      End
   End
   Begin VB.Menu masismat 
      Caption         =   "Asistente Matemáticos "
      Begin VB.Menu mconvertidor 
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
Attribute VB_Name = "frmmelgassinpendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bevaluar_Click()
On Error GoTo mensaje
If (Option1.Value = False) And (Option2.Value = False) Then
MsgBox "Seleccione la opción de cálculo según los datos conocidos", 64, "Melgas sin pendiente y sin salida"
Exit Sub
End If

qwy = Val(txtQdis.text)
etr = Val(txtetr.text)
n = Combo1.text
a = Val(txtA.text)
b = Val(txtb.text)
h1 = Val(txtH.text)
efap = Val(txtEFAP.text)
ln = Val(txtLn.text)
h = Val(txtH.text)

If etr = 0 Then
MsgBox "Ingrese el valor de evapotranspiración", 64, "Melgas sin pendiente y sin salida"
txtetr.SetFocus
Exit Sub
End If
If ln = 0 Then
MsgBox "Ingrese el valor de lámina neta", 64, "Melgas sin pendiente y sin salida"
txtLn.SetFocus
Exit Sub
End If
If qwy = 0 Then
MsgBox "Ingrese el valor conocido", 64, "Melgas sin pendiente y sin salida"
txtQdis.SetFocus
Exit Sub
End If
If Combo1.text = "" Then
MsgBox "Ingrese el valor de rugosidad", 64, "Melgas sin pendiente y sin salida"
Combo1.SetFocus
Exit Sub
End If
If efap = 0 Then
MsgBox "Ingrese el valor de eficiencia", 64, "Melgas sin pendiente y sin salida"
txtEFAP.SetFocus
Exit Sub
End If
If h = 0 Then
MsgBox "Ingrese el valor de la altura de los camellones", 64, "Melgas sin pendiente y sin salida"
txtH.SetFocus
Exit Sub
End If
If a = 0 Then
MsgBox "Ingrese valores a la ecuación de infiltración acumulada", 64, "Melgas sin pendiente y sin salida"
txtA.SetFocus
Exit Sub
End If
If b = 0 Then
MsgBox "Ingrese valores a la ecuación de infiltración acumulada", 64, "Melgas sin pendiente y sin salida"
txtb.SetFocus
Exit Sub
End If

Select Case n
    Case "Superficie lisa desnuda 0.04"
    n = 0.04
    Case "Cultivos en líneas, melgas a nivel 0.04"
    n = 0.04
    Case "Cereales sembrados en hiler, en dirección al flujo 0.10"
    n = 0.1
    Case "Alfalfa y cereales similares sembrados al voleo 0.15"
    n = 0.15
    Case "Cultivos que forman una densa unión del cultivo con el suelo, y cereales sembrados transversalmente a la dirección de flujo 0.25"
    n = 0.25
End Select
n = Val(n)
Rem lamina bruta
lb = ln / efap * 100

Rem de Frecuencia de riego
fr = ln / etr
If fr <= 1 Then
fr = 1
End If
fr1 = Int(fr)
Rem*********************
txtLB.text = Format(lb, "#0.0#")
txtFR.text = Format(fr1, "#0.0#")
Rem laminas corregidas
ln1 = fr1 * etr
lb1 = ln1 / efap * 100

txtLN1.text = Format(ln1, "#0.0#")
txtLB1.text = Format(lb1, "#0.0#")

Rem Calculo del tiempo de infiltración y avance
tinf = (ln / a) ^ (1 / b)

Select Case efap
    Case Is >= 95
        R = 6.27
    Case Is >= 90
        X = efap - 90
        R = 3.57 + 2.68 * X / 5
    Case Is >= 85
        X = efap - 85
        R = 2.5 + 1.07 * X / 5
    Case Is >= 80
        X = efap - 80
        R = 1.72 + 0.75 * X / 5
    Case Is >= 75
        X = efap - 75
        R = 1.25 + 0.45 * X / 5
    Case Is >= 70
        X = efap - 70
        R = 0.93 + 0.32 * X / 5
    Case Is >= 65
        X = efap - 65
        R = 0.69 + 0.27 * X / 5
    Case Is >= 60
        X = efap - 60
        R = 0.53 + 0.16 * X / 5
    Case Is >= 55
        X = efap - 55
        R = 0.41 + 0.12 * X / 5
    Case Is >= 50
        X = efap - 50
        R = 0.31 + 0.1 * X / 5
    Case efap < 50
        R = 0.31
End Select

ta = tinf / R
txttinf.text = Format(tinf, "###0.0##")
txtava.text = Format(ta, "####0.0#")

If Option2.Value = True Then
Rem tirante medio o caudal unitario
    qw = qwy / 1000
    Y = 1.7974 * qw ^ (9 / 16) * n ^ (6 / 16) * ta ^ (3 / 16)
    txtY.text = Format(Y * 100, "##0.0#")
Else
    Y = qwy / 100
    qw = (Y / (1.7974 * n ^ (6 / 16) * ta ^ (3 / 16))) ^ (16 / 9)
    txtY.text = Format(qw * 1000, "##0.0#")
End If
Rem Icum
bb = b - 1
aa = a * b
R1 = (1 - bb) / 2
f = (bb - R1 * bb + 2) / (1 + R1)
icum = f * aa * ta ^ (b) / ((b) * (bb + 2))

Rem longitud y tiempo de aplicación
L = qw * ta * 60 / (Y + icum / 1000)
tapl = lb / 1000 * L / (qw * 60)
txtLre.text = Format(L, "###0.0##")
txtTapl.text = Format(tapl, "####0.0#")

Rem calculo de Do
d = qw ^ (6 / 13) * n ^ (6 / 13) * L ^ (3 / 13)
If h1 <= (d * 100) Then
    text.Caption = " Alerta de desbordamiento"
Else
    text.Caption = " No existen problemas de desbordamiento"
End If

TxtDo.text = Format(d * 100, "##0.00#")

Frame3.Visible = True

Exit Sub
mensaje:
MsgBox "Introduzca adecuadamente los datos", 64, " Melgas sin Pendiente"


End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtetr.text = ""
txtLn.text = ""
txtQdis.text = ""

txtEFAP.text = ""
txtA.text = ""
txtb.text = ""
txtH.text = ""
txtFR.text = ""
txtLB.text = ""
txtLB1.text = ""
txtLN1.text = ""
txtLre.text = ""
txtY.text = ""
txtTapl.text = ""
TxtDo.text = ""

Frame3.Visible = False
Frame1.Visible = False
Option1.Value = False
Option2.Value = False
text.Caption = ""
Option1.ForeColor = &H80000012
Option2.ForeColor = &H80000012

StatusBar1.Panels(1).text = "Seleccione la opción de cálculo según los datos que Usted posea"


End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
With Combo1
    .AddItem "Superficie lisa desnuda 0.04"
    .AddItem "Cultivos en líneas, melgas a nivel 0.04"
    .AddItem "Cereales sembrados en hiler, en dirección al flujo 0.10"
    .AddItem "Alfalfa y cereales similares sembrados al voleo 0.15"
    .AddItem "Cultivos que forman una densa unión del cultivo con el suelo, y cereales sembrados transversalmente a la dirección de flujo 0.25"
End With
StatusBar1.Panels(1).text = "Seleccione la opción de cálculo según los datos que Usted posea"
End Sub



Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub marro_Click()
frmarroceras.Show
End Sub

Private Sub mcionduct_Click()
frmconductividad.Show
End Sub

Private Sub mconvertidor_Click()
frmconvertidor.Show

End Sub

Private Sub meto_Click()
frmETO.Show
End Sub

Private Sub mmcp_Click()
frmmelgaspendiente.Show
End Sub

Private Sub mmp_Click()
frmGeneral.Show
Unload Me

End Sub

Private Sub mpcmcp_Click()
frmpruebaavancerecesion.Show
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

Private Sub mt_Click()
frmtextura.Show
End Sub

Private Sub Option1_Click()
lbletiqueta.Caption = "Tirante promedio"
lbletiqueta1.Caption = "Caudal Unitario"
lblunidades.Caption = "cm"
lblunidades1.Caption = "l/s_m"
Frame1.Visible = True
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
txtetr.SetFocus

StatusBar1.Panels(1).text = "Digite los datos básicos para el diseño y oprima el botón de Evaluar "
End Sub

Private Sub Option2_Click()
lbletiqueta.Caption = "Caudal Unitario"
lbletiqueta1.Caption = "Tirante promedio"
lblunidades.Caption = "l/s_m"
lblunidades1.Caption = "cm"
Frame1.Visible = True
Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
txtetr.SetFocus
StatusBar1.Panels(1).text = "Digite los datos básicos para el diseño y oprima el botón de Evaluar "
End Sub

Private Sub ps_Click()
frmgenerales.Show
End Sub

