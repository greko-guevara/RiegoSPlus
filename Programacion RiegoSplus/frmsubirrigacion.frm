VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsubirrigacion 
   Caption         =   "Subirrigación"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmsubirrigacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   1920
      TabIndex        =   39
      Top             =   6240
      Width           =   7815
      Begin VB.CommandButton BC 
         Caption         =   "&Calcular"
         Height          =   855
         Left            =   240
         Picture         =   "frmsubirrigacion.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   855
         Left            =   6000
         Picture         =   "frmsubirrigacion.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   4080
         Picture         =   "frmsubirrigacion.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   855
         Left            =   2160
         Picture         =   "frmsubirrigacion.frx":2308
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para el diseño"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   9495
      Begin VB.TextBox txtetr 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtPp 
         Height          =   285
         Left            =   4920
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Left            =   7440
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtM 
         Height          =   285
         Left            =   7440
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtK 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   7440
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtZ 
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Evapotranspiración"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "mm/día"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Optativo **Percolación Profunda**"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "m/día"
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Espaciamiento entre drenes (L)"
         Height          =   255
         Left            =   4560
         TabIndex        =   32
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "m"
         Height          =   255
         Left            =   8760
         TabIndex        =   31
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Prof. Crítica del Nivel Freático (M)"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "m"
         Height          =   255
         Left            =   8760
         TabIndex        =   29
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "m/día"
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "m"
         Height          =   255
         Left            =   8760
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Conductividad Hidráulica"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Profundidad al Estrato Impermeable (H)"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lbletiqueta 
         Caption         =   "Bordo libre en el Canal (z)"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label17 
         Caption         =   "m"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   1200
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtr1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   720
         TabIndex        =   37
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label14 
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   2640
         TabIndex        =   40
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblunidades1 
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label respuesta 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label und 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.Label text 
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   2400
         TabIndex        =   5
         Top             =   3000
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccione el parámetro a Calcular"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      Begin VB.OptionButton optM 
         Caption         =   "Profundida Crítica de la Tabla de Agua (M)"
         Height          =   375
         Left            =   5760
         TabIndex        =   28
         Top             =   240
         Width           =   3375
      End
      Begin VB.OptionButton optE 
         Caption         =   "Espaciamiento entre drenes (L)"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton optZ 
         Caption         =   "Bordo libre en el Canal (z)"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   38
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Left            =   6000
      Picture         =   "frmsubirrigacion.frx":29F2
      Top             =   3720
      Width           =   4560
   End
   Begin VB.Label Label10 
      Caption         =   "Parámetros de Riego por Subirrigación"
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
      TabIndex        =   27
      Top             =   240
      Width           =   5175
   End
   Begin VB.Menu mpasc 
      Caption         =   "Parámetros Suelo - Clima"
      Begin VB.Menu mps 
         Caption         =   "Parámetros Suelo"
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
   Begin VB.Menu masism 
      Caption         =   "Asistente Matemático"
      Begin VB.Menu mconv 
         Caption         =   "Convertidor de Unidades "
      End
      Begin VB.Menu as 
         Caption         =   "-"
      End
      Begin VB.Menu hc 
         Caption         =   "Hidráulica de Canales"
      End
      Begin VB.Menu asd 
         Caption         =   "-"
      End
      Begin VB.Menu mregpot 
         Caption         =   "Regresión Potencial Simple"
      End
   End
   Begin VB.Menu mmp 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmsubirrigacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BC_Click()
On Error GoTo mensaje
If Frame1.Visible = False Then
MsgBox "Seleccione el dato que desee encontrar", 64, "Método de Subirrigación"
Exit Sub
End If



k = Val(txtK.text)
h = Val(txth.text)
E = Val(txtetr.text) / 1000
If k = 0 Then
MsgBox "Ingrese el valor de conductividad hidráulica", 64, "Método de Subirrigación"
txtK.SetFocus
Exit Sub
End If
If h = 0 Then
MsgBox "Ingrese el valor de profundidad del estrato impermeable", 64, "Método de Subirrigación"
txth.SetFocus
Exit Sub
End If
If E = 0 Then
MsgBox "Ingrese el valor de evapotranspiración", 64, "Método de Subirrigación"
txtetr.SetFocus
Exit Sub
End If

If optE.Value = True Then
    m = Val(txtM.text)
    z = Val(txtZ.text)
    If m = 0 Then
    MsgBox "Ingrese el valor de profundidad del crítica de la tabla de agua", 64, "Método de Subirrigación"
    txtM.SetFocus
    Exit Sub
    End If
    If z = 0 Then
    MsgBox "Ingrese el valor del nivel a piso de agua en los drenes ", 64, "Método de Subirrigación"
    txtZ.SetFocus
    Exit Sub
    End If
    
    If txtPp.text <> "" Then
        pp = Val(txtPp.text)
        L = ((4 * k * (h - z) ^ 2 - (h - m) ^ 2) / (E + pp)) ^ 0.5
    Else
        If z = 0 Then
            L = ((4 * k * m * (2 * h - m)) / E) ^ 0.5
        Else
            If m = h Then
                L = ((4 * k * (h - z) ^ 2) / E) ^ 0.5
            Else
                If m = h And z = 0 Then
                    L = ((4 * k * h ^ 2) / E) ^ 0.5
                Else
                    If m <> h And z <> 0 Then
                        L = ((4 * k * (m - z) * (2 * h - (z + m))) / E) ^ (1 / 2)
                    End If
                End If
            End If
        End If
    End If
    txtR = Format(L, "##0.0##")
    txtr1 = ""
Else
    If optZ.Value = True Then
        m = Val(txtM.text)
        L = Val(txtL.text)
        If m = 0 Then
        MsgBox "Ingrese el valor de profundidad del crítica de la tabla de agua", 64, "Método de Subirrigación"
        txtM.SetFocus
        Exit Sub
        End If
        If L = 0 Then
        MsgBox "Ingrese el valor del espaciamiento de drenes ", 64, "Método de Subirrigación"
        txtL.SetFocus
        Exit Sub
        End If
        c2 = L ^ 2 * E / (4 * k)
        c1 = (2 * h - m) * m
        c = c1 - c2
        b = -(2 * h)
        Rem****
        s1 = -(b / 2)
        s1 = Format(s1, "##0.0##")
        discriminante = b ^ 2 - 4 * c
        If discriminante < 0 Then
            s2 = Sqr(Abs(discriminante)) / (2)
            s2 = Format(s2, "##0.0##")
            txtR.text = Str(s1) + "+" + Str(s2) + "i"
            txtr1.text = Str(s1) + "" + Str(-s2) + "i"
        Else
            s2 = (Sqr(discriminante)) / 2
            txtr1.text = Format(s1 + s2, "##0.0##")
            txtR.text = Format(s1 - s2, "##0.0##")
            Label14.Caption = "Dos Soluciones matemáticos, solo una física"
        End If
    Else
        z = Val(txtZ.text)
        L = Val(txtL.text)
        If L = 0 Then
        MsgBox "Ingrese el valor de espaciamiento de drenes", 64, "Método de Subirrigación"
        txtL.SetFocus
        Exit Sub
        End If
        If z = 0 Then
        MsgBox "Ingrese el valor del nivel a piso de agua en los drenes ", 64, "Método de Subirrigación"
        txtZ.SetFocus
        Exit Sub
        End If
        c2 = L ^ 2 * E / (4 * k)
        c1 = (2 * h - z) * (z)
        c = c1 + c2
        b = -(2 * h)
        
        Rem****
        s1 = -b / (2)
        s1 = Format(s1, "##0.0##")
        discriminante = b ^ 2 - 4 * c
        If discriminante < 0 Then
            s2 = Sqr(Abs(discriminante)) / (2)
            s2 = Format(s2, "##0.0##")
            txtR.text = Str(s1) + "+" + Str(s2) + "i"
            txtr1.text = Str(s1) + "" + Str(-s2) + "i"
        Else
            s2 = Sqr(discriminante) / (2)
            txtr1.text = Format(s1 + s2, "##0.0##")
            txtR.text = Format(s1 - s2, "##0.0##")
            Label14.Caption = "Dos Soluciones matemáticos, solo una física"
            
        End If
    End If
End If
Exit Sub
mensaje:
    MsgBox "Ingrese adecuadamente todos los datos de entrada", 64, "Método de Subirrigación"
    
End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub



Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtK.text = ""
txtetr.text = ""
txtZ.text = ""
txtPp.text = ""
txtL.text = ""
txtM.text = ""
txth.text = ""
txtR.text = ""
txtr1.text = ""
Label14.Caption = ""

optE.Value = False
optZ.Value = False
optM.Value = False
Frame1.Visible = False
Frame3.Visible = False
optM.ForeColor = &H80000012

optZ.ForeColor = &H80000012

optE.ForeColor = &H80000012

StatusBar1.Panels(1).text = "Seleccione el dato que desea conocer"
End Sub

Private Sub Form_Load()
optE.Value = False
optZ.Value = False
optM.Value = False
StatusBar1.Panels(1).text = "Seleccione el dato que desea conocer"

End Sub

Private Sub hc_Click()
Frmhidraulica.Show
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

Private Sub mps_Click()
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

Private Sub mt_Click()
frmtextura.Show
End Sub

Private Sub optE_Click()
Frame1.Visible = True
Frame3.Visible = True
txtL.Enabled = False
txtL.BackColor = &H8000000F
respuesta.Caption = "Espaciamiento entre canales"
und.Caption = "m"
txtZ.BackColor = &H80000005
txtM.BackColor = &H80000005
txtZ.Enabled = True
txtM.Enabled = True
Label11.Visible = True
Label7.Visible = True
txtPp.Visible = True
StatusBar1.Panels(1).text = "Digite los parámetros de entrada y oprima el Botón de Calcular para encotrar el espaciamiento entre zanjas"
optE.ForeColor = &HC0&
optM.ForeColor = &H80000012
optZ.ForeColor = &H80000012


End Sub

Private Sub optM_Click()
Frame1.Visible = True
Frame3.Visible = True
txtM.Enabled = False
txtM.BackColor = &H8000000F
respuesta.Caption = "Profundidad inferior de la tabla de Agua"
und.Caption = "m"
txtZ.BackColor = &H80000005
txtL.BackColor = &H80000005
txtZ.Enabled = True
txtL.Enabled = True
Label11.Visible = False
Label7.Visible = False
txtPp.Visible = False
optM.ForeColor = &HC0&
optE.ForeColor = &H80000012
optZ.ForeColor = &H80000012

StatusBar1.Panels(1).text = "Digite los parámetros de entrada y oprima el Botón de Calcular para encotrar la máxima profundidad de la tabla de agua"
End Sub

Private Sub optZ_Click()
Frame1.Visible = True
Frame3.Visible = True
txtZ.Enabled = False
txtZ.BackColor = &H8000000F
optZ.ForeColor = &HC0&
respuesta.Caption = "Bordo libre en el dren"
und.Caption = "m"
txtM.BackColor = &H80000005
txtL.BackColor = &H80000005
optM.ForeColor = &H80000012
optE.ForeColor = &H80000012

txtM.Enabled = True
txtL.Enabled = True
Label11.Visible = False
Label7.Visible = False
txtPp.Visible = False
StatusBar1.Panels(1).text = "Digite los parámetros de entrada y oprima el Botón de Calcular para encotrar el bordo libre del canal"
End Sub


