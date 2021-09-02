VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpozas 
   Caption         =   "Riego por pozas"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmpozas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame3 
      Caption         =   "Datos conociendo:"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   4560
      TabIndex        =   38
      Top             =   720
      Width           =   5655
      Begin VB.OptionButton Option1 
         Caption         =   "Longitud de la melga"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Diferencias entre bordos"
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame11 
      Height          =   1215
      Left            =   2280
      TabIndex        =   27
      Top             =   6360
      Width           =   7455
      Begin VB.CommandButton bC 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   240
         Picture         =   "frmpozas.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton BL 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2040
         Picture         =   "frmpozas.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bS 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   5640
         Picture         =   "frmpozas.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton BI 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   3840
         Picture         =   "frmpozas.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      ForeColor       =   &H00800000&
      Height          =   2655
      Left            =   960
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox txtL 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtQo 
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtS 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtLn 
         Height          =   285
         Left            =   2760
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblunidades 
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbletiqueta 
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2310
         Left            =   5280
         Picture         =   "frmpozas.frx":29F2
         Top             =   240
         Width           =   4560
      End
      Begin VB.Label Label16 
         Caption         =   "Caudal (recomendado) "
         Height          =   375
         Left            =   480
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "l/s*10m2"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Pendiente"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "cm"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Lámina neta (ln)"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Textura"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultados"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   1560
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox txtZ 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2760
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtEf 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6360
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtA 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6360
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtQ 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2760
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtT 
         BackColor       =   &H80000016&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "%"
         Height          =   255
         Left            =   7680
         TabIndex        =   37
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "m2"
         Height          =   255
         Left            =   7680
         TabIndex        =   36
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbletiqueta1 
         Height          =   375
         Left            =   720
         TabIndex        =   35
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblunidades1 
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Area"
         Height          =   375
         Left            =   5160
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Eficiencia"
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "l/seg"
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "horas"
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Tiempo aplicación"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Caudal"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
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
            Object.Width           =   2646
            MinWidth        =   2646
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
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Left            =   6240
      Picture         =   "frmpozas.frx":2398C
      Top             =   1800
      Width           =   4560
   End
   Begin VB.Label Label17 
      Caption         =   "Riego por Pozas o Cuadros"
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
      Width           =   3615
   End
   Begin VB.Menu mgsc 
      Caption         =   "Generales Suelo- Clima"
      Begin VB.Menu mgs 
         Caption         =   "General Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mconductividad 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meto 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu mass 
      Caption         =   "Asistente Matemático"
      Begin VB.Menu mconv4r 
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
Attribute VB_Name = "frmpozas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub BC_Click()
On Error GoTo mensaje
If (Option1.Value = False) And (Option2.Value = False) Then
MsgBox "Seleccione la opción de cálculo según los datos conocidos", 64, "Melgas sin pendiente y sin salida"
Exit Sub
End If

ln = Val(txtLn.text)
s = Val(txtS.text)
L = Val(txtL.text)
If ln = 0 Then
MsgBox "Ingrese el valor de lámina neta", 64, " Riego por Pozas"
txtLn.SetFocus
Exit Sub
End If
If s = 0 Then
MsgBox "Ingrese el valor de la pendiente", 64, " Riego por Pozas"
txtS.SetFocus
Exit Sub
End If
If L = 0 Then
MsgBox "Ingrese el valor de la longitud", 64, " Riego por Pozas"
txtL.SetFocus
Exit Sub
End If
If Combo1.text = "" Then
MsgBox "Selecione la textura", 64, " Riego por Pozas"
Combo1.SetFocus
Exit Sub
End If

qo = Val(txtQo.text)
If qo = 0 Then
     qo = Combo1.text
    Select Case qo
        Case "Arenoso"
        qo = 1.5
        Case "Franco Arenoso"
        qo = 0.5
        Case "Franco Arcilloso"
        qo = 0.25
        Case "Arcilloso"
        qo = 0.15
    End Select
End If

If Option1.Value = True Then
z = L * s
a = L * L
ef = 2 / (2 + z / ln)

Else
z = L / s
a = z * z
ef = 2 / (2 + L / ln)

End If

q = a * qo / 10
ta = ln * a / (ef * q) * 10 / 3600
ef1 = ef * 100

txtZ.text = Format(z, "##0.0#")
txtA.text = Format(a, "##0.0#")
txtEf.text = Format(ef1, "##0.0#")
txtQ.text = Format(q, "##0.0#")
txtT.text = Format(ta, "##0.0#")
Frame2.Visible = True

Exit Sub
mensaje:
MsgBox "ingrese valores adecuados", 64, " Riego por Pozas"
End Sub





Private Sub BI_Click()
Print Form
End Sub

Private Sub BL_Click()


txtS.text = ""
txtZ.text = ""
txtLn.text = ""
txtQo.text = ""
txtL.text = ""
txtA.text = ""
txtT.text = ""
txtQ.text = ""
txtEf.text = ""
Combo1.text = ""
Frame2.Visible = False
Frame1.Visible = False
Option1.Value = False
Option2.Value = False
Option1.ForeColor = &H80000012
Option2.ForeColor = &H80000012
StatusBar1.Panels(1).text = "Seleccione la opción de cálculo según los datos que Usted posea"
End Sub

Private Sub bS_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub Combo1_Click()
qo = Combo1.text
Select Case qo
    Case "Arenoso"
    qo = 1.5
    Case "Franco Arenoso"
    qo = 0.5
    Case "Franco Arcilloso"
    qo = 0.25
    Case "Arcilloso"
    qo = 0.15
End Select

txtQo.text = qo
End Sub



Private Sub Form_Load()
Option1.Value = False
Option2.Value = False

With Combo1
    .AddItem "Arenoso"
    .AddItem "Franco Arenoso"
    .AddItem "Franco Arcilloso"
    .AddItem "Arcilloso"
End With
StatusBar1.Panels(1).text = "Seleccione la opción de cálculo según los datos que Usted posea"
End Sub

Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub mconductividad_Click()
frmconductividad.Show
End Sub

Private Sub mconv4r_Click()
frmconvertidor.Show
End Sub

Private Sub meto_Click()
frmETO.Show
End Sub

Private Sub mgs_Click()
frmgenerales.Show

End Sub

Private Sub mmp_Click()
frmGeneral.Show
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

Private Sub Option1_Click()
lbletiqueta.Caption = "Longitud de melga"
lbletiqueta1.Caption = "Diferencia entre bordos"
lblunidades.Caption = "mts"
lblunidades1.Caption = "cm"
Frame1.Visible = True
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
txtLn.SetFocus
StatusBar1.Panels(1).text = "Digite los datos básicos para el diseño y oprima el botón de Evaluar "
End Sub

Private Sub Option2_Click()
lbletiqueta1.Caption = "Longitud de melga"
lbletiqueta.Caption = "Diferencia entre bordos"
lblunidades1.Caption = "mts"
lblunidades.Caption = "cm"
Frame1.Visible = True
Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
txtLn.SetFocus
StatusBar1.Panels(1).text = "Digite los datos básicos para el diseño y oprima el botón de Evaluar "

End Sub
