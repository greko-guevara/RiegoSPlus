VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmelgaspendiente 
   Caption         =   "Melgas Con Pendiente"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmmelgaspendiente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame10 
      Height          =   1215
      Left            =   840
      TabIndex        =   69
      Top             =   6000
      Width           =   5295
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   120
         Picture         =   "frmmelgaspendiente.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   1800
         MaskColor       =   &H000000FF&
         Picture         =   "frmmelgaspendiente.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   3600
         Picture         =   "frmmelgaspendiente.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   5040
      TabIndex        =   49
      Top             =   840
      Width           =   6135
      Begin VB.TextBox txtLre 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4200
         TabIndex        =   55
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtQor 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4200
         TabIndex        =   54
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtLB 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   53
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtLB1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   51
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtLN1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton BEFAP 
         Caption         =   "&Evaluar"
         Height          =   615
         Left            =   240
         Picture         =   "frmmelgaspendiente.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   2160
         Picture         =   "frmmelgaspendiente.frx":29F2
         Top             =   1440
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label Label24 
         Caption         =   "mts"
         Height          =   255
         Left            =   5520
         TabIndex        =   67
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Longitud Recomenda "
         Height          =   255
         Left            =   4200
         TabIndex        =   66
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "l/s x 100m2"
         Height          =   375
         Left            =   5520
         TabIndex        =   65
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Caudal Unitario real"
         Height          =   375
         Left            =   4200
         TabIndex        =   64
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "mm"
         Height          =   255
         Left            =   3600
         TabIndex        =   63
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Lámina Bruta"
         Height          =   255
         Left            =   2280
         TabIndex        =   62
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "día"
         Height          =   255
         Left            =   1560
         TabIndex        =   61
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "Frecuencia de Riego"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Lámina Bruta corregida"
         Height          =   255
         Left            =   2280
         TabIndex        =   59
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "mm"
         Height          =   255
         Left            =   3600
         TabIndex        =   58
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "Lámina Neta corregida"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "mm"
         Height          =   255
         Left            =   1560
         TabIndex        =   56
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2415
      Left            =   5040
      TabIndex        =   30
      Top             =   3000
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtL 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtQm 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   360
         TabIndex        =   70
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtQmin 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4320
         TabIndex        =   36
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtAM 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtY 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4320
         TabIndex        =   34
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtQmax 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   33
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton bcalcular2 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   240
         Picture         =   "frmmelgaspendiente.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtB 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtTR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   2160
         X2              =   2160
         Y1              =   1560
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   2160
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label45 
         Caption         =   "mts"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1680
         TabIndex        =   75
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "lps"
         Height          =   255
         Left            =   1680
         TabIndex        =   74
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label39 
         Caption         =   "Longitud Revaluada según terreno"
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label38 
         Caption         =   "mts"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1680
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label40 
         Caption         =   "Caudal aplicar"
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "lps"
         Height          =   255
         Left            =   5640
         TabIndex        =   48
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   "Caudal mínimo"
         Height          =   255
         Left            =   4320
         TabIndex        =   47
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "m2"
         Height          =   255
         Left            =   3600
         TabIndex        =   46
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label33 
         Caption         =   "Area de la Melga"
         Height          =   255
         Left            =   2280
         TabIndex        =   45
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label txtYfdgfdg 
         Caption         =   "Tirante Y"
         Height          =   255
         Left            =   4320
         TabIndex        =   44
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label35 
         Caption         =   "cm"
         Height          =   255
         Left            =   5640
         TabIndex        =   43
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label36 
         Caption         =   "Caudal máximo"
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label37 
         Caption         =   "lps"
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label41 
         Caption         =   "Bordo B"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label42 
         Caption         =   "cm"
         Height          =   255
         Left            =   5640
         TabIndex        =   39
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label43 
         Caption         =   "Tiempo de Riego"
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   "horas"
         Height          =   255
         Left            =   3600
         TabIndex        =   37
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para el diseño"
      ForeColor       =   &H00800000&
      Height          =   4455
      Left            =   600
      TabIndex        =   14
      Top             =   840
      Width           =   4215
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   360
         TabIndex        =   27
         Top             =   3480
         Width           =   3375
         Begin VB.TextBox txtEFAP 
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton beficiencia 
            Caption         =   "E&stimar"
            Height          =   375
            Left            =   2160
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "%"
            Height          =   255
            Left            =   1560
            TabIndex        =   29
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label18 
            Caption         =   "Eficiencia de Aplicación"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox txtIb 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtETR 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtLN 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtWC 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtS 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Coeficente de Manning n (Puede ser digitado o seleccionado)"
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "cm/hr"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "mm/dia"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "mm"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "mts"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Infiltración base"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Evapotranspiración real"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Lámina Neta "
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Ancho de melga W "
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Pendiente "
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   68
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
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   6600
      Picture         =   "frmmelgaspendiente.frx":335C
      Top             =   5640
      Width           =   4185
   End
   Begin VB.Label Label10 
      Caption         =   "Diseño de  Melgas con Pendiente"
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
      TabIndex        =   20
      Top             =   240
      Width           =   4575
   End
   Begin VB.Menu psc 
      Caption         =   "Parámetros Suelo - Clima"
      Begin VB.Menu ps 
         Caption         =   "Parámetros Suelo"
      End
      Begin VB.Menu mte 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcond 
         Caption         =   "Conductividad"
      End
      Begin VB.Menu meto 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu mocm 
      Caption         =   "Otros Cálculos con Melgas"
      Begin VB.Menu mpcmcp 
         Caption         =   "Pruebas de Campo Melgas Con Pendiente"
      End
      Begin VB.Menu mmsp 
         Caption         =   "Melgas Sin Pendiente"
      End
      Begin VB.Menu marrr 
         Caption         =   "Arroceras"
      End
   End
   Begin VB.Menu mmat 
      Caption         =   "Asistente Matemático"
      Begin VB.Menu mconv 
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
Attribute VB_Name = "frmmelgaspendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bcalcular2_Click()
On Error GoTo mensaje
L = Val(txtL.text)
w = Val(txtWC.text)
s = Val(txtS.text)
n = Combo1.text
lb = Val(txtLB.text)
qor = Val(txtQor.text)
If w = 0 Then
MsgBox "Ingrese el valor del ancho de melga", 64, "Melgas con Pendiente"
txtWC.SetFocus
Exit Sub
End If
If Combo1.text = "" Then
MsgBox "selecione el valor de la rugosidad o digítelo", 64, "Melgas con Pendiente"
Combo1.SetFocus
Exit Sub
End If
If L = 0 Then
MsgBox "Ingrese el valor de la longitud de la melga", 64, "Melgas con Pendiente"
txtL.SetFocus
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
Rem Calculo del area
a = L * w
Rem caudal por melga
qm = qor * a / 100
Rem calculo del caudal maximo
If n >= 0.15 Then
    qmax = 0.354 * (s / 100) ^ (-0.75) * w
Else
    qmax = 0.177 * (s / 100) ^ (-0.75) * w
End If
Rem caudal minimo
qmin = 0.0195 * L * (s / 100) ^ 0.5 * w / n
Rem caudal recesion
Qred = qmax / 3



Rem tirante
Y = ((qm / 1000 * n / (w * (s / 100) ^ 0.5)) ^ (3 / 5)) * 100
Rem tiempo de riego
tr = lb / 10 * a / (qm * 360)
Rem bordo
b = 1.2 * Y
Rem numero de melgas


txtAM.text = Format(a, "#0.00#")
txtQmax.text = Format(qmax, "#0.00#")
txtQmin.text = Format(qmin, "#0.0#")
txtQm.text = Format(qm, "#0.00#")
txtY.text = Format(Y, "#0.0#")
txtTR.text = Format(tr, "#0.0#")
txtb.text = Format(b, "#0.0#")



Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Melgas con Pendiente"




End Sub


Private Sub BEFAP_Click()
On Error GoTo mensaje

s = Val(txtS.text)
ln = Val(txtLn.text)
etr = Val(txtetr.text)
ib = Val(txtIb.text)
If ib = 0 Then
MsgBox "Ingrese el valor de la Infiltración básica", 64, "Melgas con Pendiente"
txtIb.SetFocus
Exit Sub
End If
If etr = 0 Then
MsgBox "Ingrese el valor de la evapotranspiración", 64, "Melgas con Pendiente"
txtetr.SetFocus
Exit Sub
End If
If ln = 0 Then
MsgBox "Ingrese el valor de la lámina neta", 64, "Melgas con Pendiente"
txtLn.SetFocus
Exit Sub
End If
If s = 0 Then
MsgBox "Ingrese el valor de la pendiente", 64, "Melgas con Pendiente"
txtS.SetFocus
Exit Sub
End If


Rem lamina bruta
efap = Val(txtEFAP.text)
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
Rem calculo de caudales unitarios y longitud maxima recomendada
qo = 10.9875681 * ib ^ 1.0080137 * (lb / 10) ^ -1.1050021
qor = 0.86214 * s ^ -0.214 * qo
lre = 554.1758954 * s ^ -0.7843749 * qor ^ -1.0036026
txtQor.text = Format(qor, "##0.00##")
txtLre.text = Format(lre, "##0.00##")
Frame5.Visible = True
Image1.Visible = True

txtL.SetFocus
StatusBar1.Panels(1).text = "Seleccione la distancia definitiva de las melgas y oprima el Botón de Calcular"
Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Melgas con Pendiente"

End Sub

Private Sub beficiencia_Click()
On Error GoTo mensaje
Rem Eficiencia en la aplicacion
s = Val(txtS.text)
lb = Val(txtIb.text)
If s = 0 Then
MsgBox "Ingrese el valor de la pendiente", 64, "Melgas con Pendiente"
txtS.SetFocus
Exit Sub
End If
If lb = 0 Then
MsgBox "Ingrese el valor de la infiltracion básica", 64, "Melgas con Pendiente"
txtIb.SetFocus
Exit Sub
End If

If s <= 0.5 Then
    Select Case lb
        Case Is < 1.27
        txtEFAP.text = 80
        Case Is >= 1.27
        txtEFAP.text = 70
    End Select
    
Else
    If s <= 1 Then
        Select Case lb
            Case Is < 0.76
            txtEFAP.text = 65
            Case Is >= 0.76
            txtEFAP.text = 70
        End Select
    Else
        If s <= 2 Then
            Select Case lb
                Case Is < 0.76
                txtEFAP.text = 60
                Case Is >= 1.27
                txtEFAP.text = 65
                Case Is >= 5.08
                txtEFAP.text = 70
                Case Is >= 10.16
                txtEFAP.text = 75
            End Select
        Else
            If s >= 4 Then
                Select Case lb
                    Case Is < 0.76
                    txtEFAP.text = 55
                    Case Is >= 1.27
                    txtEFAP.text = 60
                    Case Is >= 5.08
                    txtEFAP.text = 65
                    Case Is >= 10.16
                    txtEFAP.text = 60
                End Select
             End If
        End If
    End If
End If
Exit Sub
mensaje:
   MsgBox "Ingrese adecuadamente los datos", 64, "Melgas con Pendiente"

End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtIb.text = ""
txtetr.text = ""
txtLn.text = ""
txtWC.text = ""
txtS.text = ""

txtEFAP.text = ""
txtFR.text = ""
txtLN1.text = ""
txtLB.text = ""
txtLB1.text = ""
txtQor.text = ""
txtLre.text = ""
txtL.text = ""
txtQm.text = ""
txtAM.text = ""
txtQmax.text = ""
txtQmin.text = ""
txtY.text = ""
txtb.text = ""
txtIb.SetFocus
Frame5.Visible = False
Image1.Visible = False

End Sub

Private Sub Form_Load()
With Combo1
    .AddItem "Superficie lisa desnuda 0.04"
    .AddItem "Cultivos en líneas, melgas a nivel 0.04"
    .AddItem "Cereales sembrados en hiler, en dirección al flujo 0.10"
    .AddItem "Alfalfa y cereales similares sembrados al voleo 0.15"
    .AddItem "Cultivos que forman una densa unión del cultivo con el suelo, y cereales sembrados transversalmente a la dirección de flujo 0.25"
End With
StatusBar1.Panels(1).text = "Digite los datos básicos para el diseño y oprima el botón de Evaluar para Iniciar el proceso de cálculo  "
End Sub




Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub marrr_Click()
frmarroceras.Show

End Sub

Private Sub mcond_Click()
frmconductividad.Show
End Sub

Private Sub mconv_Click()
frmconvertidor.Show
End Sub

Private Sub meto_Click()
frmETO.Show
End Sub

Private Sub mmp_Click()
frmGeneral.Show
Unload Me
End Sub

Private Sub mmsp_Click()
frmmelgassinpendiente.Show

End Sub

Private Sub mpcmcp_Click()
frmpruebaavancerecesion.Show
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

Private Sub mte_Click()
frmtextura.Show
End Sub

Private Sub ps_Click()
frmgenerales.Show
End Sub
