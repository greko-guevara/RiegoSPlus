VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmhidraulica 
   Caption         =   "Cálculos del tirante de los diferentes parámetros hidráulicos"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "FrmhidraulicaCa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.OptionButton O2 
      Caption         =   "Parabólica"
      Height          =   255
      Left            =   6600
      TabIndex        =   48
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton O1 
      Caption         =   "Trapezoidal - Rectangular -  triangular  "
      Height          =   255
      Left            =   1920
      TabIndex        =   47
      Top             =   840
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Canal"
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   1673
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox Txtq 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Text            =   " "
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Txtb 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Text            =   " "
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Txtz 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Text            =   " "
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Txtn 
         Height          =   285
         Left            =   6480
         TabIndex        =   4
         Text            =   " "
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Left            =   6480
         TabIndex        =   5
         Text            =   " "
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "m"
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "m^3/s:"
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Caudal "
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Talud:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Rugosidad:"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Pendiente:"
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultados"
      ForeColor       =   &H00800000&
      Height          =   3255
      Left            =   1673
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox Txtv 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6360
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Txte 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Txttf 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Txtf 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Txtt 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6360
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Txtr 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6360
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Txta 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Txtp 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Txty 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1560
         Left            =   4320
         Picture         =   "FrmhidraulicaCa.frx":0CCA
         Top             =   1560
         Visible         =   0   'False
         Width           =   3810
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1560
         Left            =   4320
         Picture         =   "FrmhidraulicaCa.frx":132CC
         Top             =   1560
         Visible         =   0   'False
         Width           =   3810
      End
      Begin VB.Label Label14 
         Caption         =   "Energía Específica "
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Flujo"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Número de Froude:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Velocidad (V)"
         Height          =   255
         Left            =   4320
         TabIndex        =   42
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Espejo de Agua  (T)"
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Radio Hidráulico (R)"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Área Hidráulica (A)"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Perímetro Mojado (P) "
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Tirante Normal (Y)"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "m/s"
         Height          =   375
         Left            =   7680
         TabIndex        =   36
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "m"
         Height          =   375
         Left            =   7680
         TabIndex        =   35
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "m"
         Height          =   375
         Left            =   7680
         TabIndex        =   34
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "m Kg/Kg"
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label25 
         Height          =   375
         Left            =   3480
         TabIndex        =   32
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "m2"
         Height          =   375
         Left            =   3480
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "m"
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "m"
         Height          =   375
         Left            =   3480
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label17 
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   2273
      TabIndex        =   0
      Top             =   6480
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   3840
         Picture         =   "FrmhidraulicaCa.frx":258CE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton bejecutar 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   240
         Picture         =   "FrmhidraulicaCa.frx":26038
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2040
         Picture         =   "FrmhidraulicaCa.frx":267A2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cmdsalir 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   5640
         Picture         =   "FrmhidraulicaCa.frx":26E8C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   49
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
            TextSave        =   "07/06/2005"
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
   Begin VB.Label Label26 
      Caption         =   "Cálculo del tirante y otros parámetros hidráulicos (flujo uniforme)"
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
      TabIndex        =   20
      Top             =   240
      Width           =   8895
   End
   Begin VB.Menu in 
      Caption         =   "Infiltración en canales"
   End
   Begin VB.Menu mm 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "Frmhidraulica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public q As Double
Public b As Double
Public z As Double
Public n As Double
Public s As Double
Public Y As Double
Public P As Double
Public a As Single
Public T As Double
Public R As Double
Public V As Double
Public f As Double
Public E As Double
Public L As Double
Public X1 As Double
Public M2 As Double
Public dy As Double
Public k As Integer
Public i As Integer
Sub trapecio()

a = (b + z * Y) * Y
P = b + 2 * Y * L
T = b + 2 * z * Y
X1 = q - ((a) ^ (5 / 3) / P ^ (2 / 3) / n * Sqr(s))
End Sub
Sub parábola()
If b = 0 Then
mensaje = MsgBox("Ingrese el valor del espejo de agua", 64, "Parámetros Hidráulicos")
txtb.SetFocus
Exit Sub
End If

a = 2 / 3 * b * Y
u = 4 * Y / b
If (u > 0) And (u <= 1) Then
    P = b + (8 * (Y ^ 2) / (3 * b))
Else
    P = b / 2 * (((1 + u ^ 2) ^ 0.5) + (1 / u * Log(u + (1 + u ^ 2) ^ 0.5)))
End If
X1 = q - ((a) ^ (5 / 3) / P ^ (2 / 3) / n * Sqr(s))
T = b
End Sub
Private Sub bejecutar_Click()
On Error GoTo mensaje:
'ingreso de valores
q = Val(txtQ.text)
b = Val(txtb.text)
z = Val(txtZ.text)
n = Val(Txtn.text)
s = Val(txtS.text)
'verificando valores
If q = 0 Then
mensaje = MsgBox("Ingrese el valor del caudal", 64, "Parámetros Hidráulicos")
txtQ.SetFocus
Exit Sub
End If
If n = 0 Then
mensaje = MsgBox("Ingrese el valor de rugosidad", 64, "Parámetros Hidráulicos")
Txtn.SetFocus
Exit Sub
End If
If s = 0 Then
mensaje = MsgBox("Ingrese el valor de pendiente", 64, "Parámetros Hidráulicos")
txtS.SetFocus
Exit Sub
End If

'determinación del tipo de sección
If b = 0 Then
Sección = "TRIANGULAR"
End If
If z = 0 Then
Sección = "RECTANGULAR"
Else
Sección = "TRAPEZOIDAL"
End If
Y = 1
L = Sqr(1 + z ^ 2)
calculod:
If O1.Value = True Then
trapecio
Else
parábola
End If
M2 = X1
Y = Y - 0.001
If O1.Value = True Then
trapecio
Else
parábola
End If
dy = 0.001 * M2 / (M2 - X1)
Y = Y - dy + 0.001
If Abs(dy) >= 0.00001 Then GoTo calculod:
If O1.Value = True Then
trapecio
Else
parábola
End If
R = a / P
V = q / a
f = V / Sqr(9.81 * a / T)
E = Y + V ^ 2 / 19.62
'tipo de flujo
If f > 1 Then
Txttf.text = "SUPERCRÍTICO"
End If
If f = 1 Then
Txttf.text = "CRÍTICO"
End If
If f < 1 Then
Txttf.text = "SUBCRÍTICO"
End If
txtY.text = Format(Y, "###0.00##")
txtA.text = Format(a, "###0.00##")
txtp.text = Format(P, "###0.00##")
txtR.text = Format(R, "###0.00##")
txtt.text = Format(T, "###0.00##")
Txtv.text = Format(V, "###0.00##")
Txtf.text = Format(f, "###0.00##")
Txte.text = Format(E, "###0.00##")
Frame2.Visible = True
Exit Sub
mensaje:
    MsgBox "Ingrese los datos en forma correcta", 16, "Parámetros Hidráulicos"
End Sub



Private Sub Cmdnuevo_Click()
txtQ.text = ""
txtb.text = ""
Image1.Visible = False
Image2.Visible = False

StatusBar1.Panels(1).text = " Seleccione el tipo de sección del canal"
txtZ.text = ""
Txtn.text = ""
txtS.text = ""
txtY.text = ""
txtA.text = ""
txtp.text = ""
txtR.text = ""
txtt.text = ""
Txttf.text = ""
Txte.text = ""
Txtv.text = ""
Txtf.text = ""
txtQ.SetFocus
O1.ForeColor = &H80000012
O2.ForeColor = &H80000012
O1.Value = False
O2.Value = False
Frame1.Visible = False
Frame2.Visible = False
Y = 0
q = 0
b = 0
z = 0

End Sub

Private Sub Cmdsalir_Click()
If Y > 0 Then frminfiltracioncanales.tT.text = Format(Y, "0.0#####")
If q > 0 Then frminfiltracioncanales.tQ.text = Format(q, "0.0#####")
If b > 0 Then frminfiltracioncanales.tB.text = Format(b, "0.0#####")
If z > 0 Then frminfiltracioncanales.tZ.text = Format(z, "0.0#####")
Unload Me
End Sub

Private Sub Command1_Click()
Print Form
End Sub

Private Sub Form_Load()
Y = 0
q = 0
b = 0
z = 0
O1.Value = False
O2.Value = False
O1.ForeColor = &H80000012
O2.ForeColor = &H80000012

StatusBar1.Panels(1).text = " Seleccione el tipo de sección del canal"
End Sub



Private Sub in_Click()
frminfiltracioncanales.Show
End Sub

Private Sub mm_Click()
If Y > 0 Then frminfiltracioncanales.tT.text = Format(Y, "0.0#####")
If q > 0 Then frminfiltracioncanales.tQ.text = Format(q, "0.0#####")
If b > 0 Then frminfiltracioncanales.tB.text = Format(b, "0.0#####")
If z > 0 Then frminfiltracioncanales.tZ.text = Format(z, "0.0#####")
Unload Me
End Sub

Private Sub O1_Click()
Frame1.Visible = True
Label2.Caption = "Ancho de Solera (b)"
Label3.Caption = "Talud:"
txtZ.Visible = True
Image2.Visible = True
Image1.Visible = False
Frame2.Visible = True

StatusBar1.Panels(1).text = "Digite los datos del Canal"
O1.ForeColor = &HC0&
O2.ForeColor = &H80000012

End Sub

Private Sub O2_Click()
Frame1.Visible = True
Frame2.Visible = True

txtZ.Visible = False
Label2.Caption = "Espejo de agua (T)"
Label3.Caption = ""
O2.ForeColor = &HC0&
O1.ForeColor = &H80000012
Image2.Visible = False
Image1.Visible = True

StatusBar1.Panels(1).text = "Digite los datos del Canal"
End Sub
