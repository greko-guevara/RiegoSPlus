VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDAMicro 
   Caption         =   "Diseño agronómico riego por micro- aspersión"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11400
   Icon            =   "frmDAMicro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11400
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   240
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox txtRT 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   58
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtetrK 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   57
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtTR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5820
         TabIndex        =   56
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtK 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   55
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLB 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   5820
         TabIndex        =   54
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtFR 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2220
         TabIndex        =   53
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "mm"
         Height          =   255
         Left            =   7140
         TabIndex        =   70
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label28 
         Caption         =   "días"
         Height          =   255
         Left            =   3540
         TabIndex        =   69
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "Relación de transpiración"
         Height          =   255
         Left            =   300
         TabIndex        =   68
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "mm/h"
         Height          =   255
         Left            =   3540
         TabIndex        =   67
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Precipitación "
         Height          =   255
         Left            =   300
         TabIndex        =   66
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label h 
         Caption         =   "hrs"
         Height          =   255
         Left            =   7140
         TabIndex        =   65
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label26 
         Caption         =   "Tiempo de riego"
         Height          =   255
         Left            =   4140
         TabIndex        =   64
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "m3/h"
         Height          =   255
         Left            =   3540
         TabIndex        =   63
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "Caudal por hectarea"
         Height          =   255
         Left            =   300
         TabIndex        =   62
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label30 
         Caption         =   "Lámina bruta"
         Height          =   255
         Left            =   4140
         TabIndex        =   61
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Frecuencia de riego"
         Height          =   255
         Left            =   300
         TabIndex        =   60
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "días"
         Height          =   255
         Left            =   3540
         TabIndex        =   59
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1095
      Left            =   2798
      TabIndex        =   12
      Top             =   6600
      Width           =   6255
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   4425
         Picture         =   "frmDAMicro.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2385
         MaskColor       =   &H000000FF&
         Picture         =   "frmDAMicro.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   345
         Picture         =   "frmDAMicro.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   7770
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
            TextSave        =   "8/21/2008"
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   6480
      TabIndex        =   42
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Diseño Agronómico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Selección de Microaspersor"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos"
      ForeColor       =   &H00800000&
      Height          =   2655
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   8415
      Begin VB.TextBox txtEM 
         Height          =   285
         Left            =   1920
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "¿Calcular...?"
         Height          =   615
         Left            =   3600
         TabIndex        =   38
         Top             =   120
         Width           =   4695
         Begin VB.OptionButton Option3 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   3240
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Lámina Neta"
            Height          =   255
            Left            =   1920
            TabIndex        =   40
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Consumo por planta"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.TextBox txtEL 
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton bc 
         Caption         =   "&Evaluar"
         Height          =   615
         Left            =   1440
         Picture         =   "frmDAMicro.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtDM 
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtQM 
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtLN 
         Height          =   285
         Left            =   6120
         TabIndex        =   0
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtA 
         Height          =   285
         Left            =   6120
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtEF 
         Height          =   285
         Left            =   6120
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Espaciamiento  micros"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "m"
         Height          =   255
         Left            =   3240
         TabIndex        =   44
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "m"
         Height          =   255
         Left            =   3240
         TabIndex        =   37
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Espaciamiento laterales"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "lts/h"
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Diámetro del micro"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label21 
         Caption         =   "mts"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "Caudal del micro"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "l/día"
         Height          =   255
         Left            =   7440
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Area =marco de cultivo"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Consumo por planta"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblunidades 
         Caption         =   "m2"
         Height          =   255
         Left            =   7440
         TabIndex        =   8
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label36 
         Caption         =   "%"
         Height          =   255
         Left            =   7440
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label37 
         Caption         =   "Eficiencia"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Selección de micro"
      Height          =   2655
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox txtib 
         Height          =   285
         Left            =   2520
         TabIndex        =   49
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtcmax 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6480
         TabIndex        =   46
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPAR1 
         Height          =   285
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtA1 
         Height          =   285
         Left            =   2520
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Bevaluar 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   1440
         Picture         =   "frmDAMicro.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtDMM 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6480
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label48 
         Caption         =   "Infiltración base"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "mm/h"
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label41 
         Caption         =   "lts/h"
         Height          =   255
         Left            =   7920
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "Caudal máximo del micro"
         Height          =   255
         Left            =   4560
         TabIndex        =   47
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label35 
         Caption         =   "El micro a seleccionar debe superar el diámetro calculado y su descarga debe ser menor que la propuesta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   4560
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label45 
         Caption         =   "Porcentaje área regada (PAR)"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label44 
         Caption         =   "%"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "m2"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Area =marco de cultivo"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "mts"
         Height          =   255
         Left            =   7920
         TabIndex        =   19
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Diámetro de mínimo del micro aspersor"
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2460
      Left            =   8760
      Picture         =   "frmDAMicro.frx":315C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2865
   End
   Begin VB.Label Label17 
      Caption         =   "Diseño agronómico Riego por micro- aspersión"
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
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   6855
   End
   Begin VB.Menu hhhhhhhh 
      Caption         =   "Parámetros Suelo- clima"
      Begin VB.Menu mgelsu 
         Caption         =   "Generales Suelo"
      End
      Begin VB.Menu mte 
         Caption         =   "Textura"
      End
      Begin VB.Menu mco 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meva 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu jjjjjjjjjj 
      Caption         =   "Hidráulica de Tuberías"
      Begin VB.Menu mca 
         Caption         =   "Cálculo en la lateral"
      End
      Begin VB.Menu mrpin 
         Caption         =   "Cálculo en la principal"
      End
      Begin VB.Menu mselbo 
         Caption         =   "Selección de la bomba"
      End
      Begin VB.Menu mcomdi 
         Caption         =   "Combinación de diámetros"
      End
   End
   Begin VB.Menu mmmppppssss 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmDAMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double

Private Sub BC_Click()
On Error GoTo mensaje:


dm = Val(txtDM.text)
qm = Val(txtQm.text)
ln = Val(txtLN.text)
a = Val(txtA.text)
EL = Val(txtEL.text)
em = Val(txtEM.text)


If EL = 0 Then
    MsgBox "Ingrese el valor el espaciamiento entre laterales", 64, "Riego por micro- aspersión"
    txtPAR.SetFocus
    Exit Sub
End If
If em = 0 Then
    MsgBox "Ingrese el valor el epaciamiento entre micros", 64, "Riego por micro- aspersión"
    txtEf.SetFocus
    Exit Sub
End If
If ln = 0 Then
    If Option1.Value = True Then
    MsgBox "Ingrese el consumo por planta", 64, "Riego por microaspersión"
    Else
    If Option2.Value = True Then
    MsgBox "Ingrese la lámina neta", 64, "Riego por microaspersión"
    Else
    MsgBox "Ingrese la lámina bruta", 64, "Riego por microaspersión"
    End If
    End If
    txtLN.SetFocus
    Exit Sub
End If
If a = 0 Then
    MsgBox "Ingrese el valor del marco del cultivo", 64, "Riego por micro- aspersión"
    txtA.SetFocus
    Exit Sub
End If
If dm = 0 Then
    MsgBox "Ingrese el valor del diámetro del micro", 64, "Riego por micro- aspersión"
    txtDM.SetFocus
    Exit Sub
End If
If qm = 0 Then
    MsgBox "Ingrese el valor del caudal del micro", 64, "Riego por micro- aspersión"
    tQ.SetFocus
    Exit Sub
End If

parc = 3.14159 * dm ^ 2 / 4 * 1 / a * 100
qhas = qm * 100 / EL * 100 / em / 1000


If Option1.Value = True Then
    esparbol = a / EL
    microsplanta = esparbol / em
    prec = microsplanta * qm
    tr = ln / prec
      lb = ln / a
    Label10.Caption = "l/h"
Else
    If Option2.Value = True Then
        ef = Val(txtEf.text)
        If ef = 0 Then
            MsgBox "Ingrese el valor de la eficiencia", 64, "Riego por micro- aspersión"
            txtEf.SetFocus
            Exit Sub
        End If
        prec = qm / (EL * em)
        lb = ln / ef * 100
        tr = lb / prec
        Label10.Caption = "mm/h"
    Else
        prec = qm / (EL * em)
        tr = ln / prec
        lb = ln
        
        Label10.Caption = "mm/h"
    End If
End If



txtK.text = Format(qhas, "##0.00##")
txtetrK.text = Format(prec, "##0.00##")

txtLB.text = Format(lb, "##0.00##")
txtRT.text = Format(parc, "##0.00##")
txtTR.text = Format(tr, "##0.00##")

Frame2.Visible = True
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"

End Sub

Private Sub bevaluar_Click()
On Error GoTo mensaje:
PAR1 = Val(txtPAR1.text)
a1 = Val(txtA1.text)
ib = Val(txtIb.text)
If PAR1 = 0 Then
    MsgBox "Ingrese el valor del Porcentaje de área humedecida", 64, "Riego por microaspersión"
    txtPAR1.SetFocus
    Exit Sub
End If
If a1 = 0 Then
    MsgBox "Ingrese el valor del marco de plantación", 64, "Riego por microaspersión"
   txtA1.SetFocus
    Exit Sub
End If
If ib = 0 Then
    MsgBox "Ingrese el valor del infiltración básica", 64, "Riego por microaspersión"
   txtIb.SetFocus
    Exit Sub
End If
cmax = ib * a1
d1 = (PAR1 / 100 * a1 * 4 / 3.14159) ^ 0.5
Label35.Visible = True
txtDMM.text = Format(d1, "##0.00##")
txtcmax.text = Format(cmax, "##0.00##")

StatusBar1.Panels(1).text = "Digite los datos del micro y oprima el botón de Evaluar para continuar con el diseño agronómico"
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
txtcmax.text = ""
txtPAR1.text = ""
txtA1.text = ""
txtIb.text = ""
txtEf.text = ""
txtLN.text = ""
txtA.text = ""
txtDMM.text = ""
txtDM.text = ""
txtQm.text = ""
txtLB.text = ""
txtFR.text = ""
txtTR.text = ""
Frame2.Visible = False
bC.Value = False

End Sub

Private Sub Form_Load()
StatusBar1.Panels(1).text = "Ingrese los datos básicos y oprima el botón de Calcular para estimar el diámetro mínimo del micro"
End Sub





Private Sub mca_Click()
FrmHLaterales.Show
End Sub

Private Sub mco_Click()
frmconductividad.Show
End Sub

Private Sub mcomdi_Click()
frmcombDia.Show
End Sub

Private Sub meva_Click()
frmETO.Show
End Sub

Private Sub mgelsu_Click()
frmgenerales.Show
End Sub

Private Sub mmmppppssss_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub mrpin_Click()
frmHprincipal.Show
End Sub

Private Sub mselbo_Click()
frmbomba.Show
End Sub

Private Sub mte_Click()
frmtextura.Show
End Sub

Private Sub Option1_Click()
txtEf.Visible = False
Label36.Visible = False
Label37.Visible = False
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
Label5.Caption = "Consumo por planta"
Label13.Caption = "l/día"
End Sub
Private Sub Option2_Click()
txtEf.Visible = True
Label36.Visible = True
Label37.Visible = True
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
Label5.Caption = "lámina neta"
Label13.Caption = "mm/día"
End Sub

Private Sub Option3_Click()
txtEf.Visible = False
Label36.Visible = False
Label37.Visible = False
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
Label5.Caption = "lámina bruta"
Label13.Caption = "mm/día"
End Sub



Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    Frame1.Visible = True
    Frame6.Visible = False
    StatusBar1.Panels(1).text = "Digite los datos de entrada y oprima el botón de Evaluar"

    Case 2
    Frame6.Visible = True
    Frame1.Visible = False
    StatusBar1.Panels(1).text = "Ingrese los datos de entrada y oprima el botón de Calcular para ver los requerimientos de los aspersores"
End Select
End Sub
