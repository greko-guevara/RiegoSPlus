VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmgenerales 
   Caption         =   "Parámetros Generales del Suelo"
   ClientHeight    =   7050
   ClientLeft      =   1290
   ClientTop       =   1080
   ClientWidth     =   8460
   Icon            =   "frmgenerales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   8460
   Begin VB.Frame Frame10 
      Height          =   1215
      Left            =   120
      TabIndex        =   72
      Top             =   4680
      Width           =   5175
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   3480
         Picture         =   "frmgenerales.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   1800
         MaskColor       =   &H000000FF&
         Picture         =   "frmgenerales.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   120
         Picture         =   "frmgenerales.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   6675
      Width           =   8460
      _ExtentX        =   14923
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
            TextSave        =   "12/06/2005"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Humedad en el Suelo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "C. C y P.M.P"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cálculo de láminas"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame7 
      Caption         =   "Lámina maxima y neta"
      ForeColor       =   &H00800000&
      Height          =   3375
      Left            =   240
      TabIndex        =   52
      Top             =   840
      Visible         =   0   'False
      Width           =   8055
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   120
         TabIndex        =   60
         Top             =   2040
         Width           =   7695
         Begin VB.TextBox txtlnl 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   62
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtlml 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2880
            TabIndex        =   61
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label33 
            Caption         =   "cm"
            Height          =   255
            Left            =   1800
            TabIndex        =   67
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label32 
            Caption         =   "Lámina Neta"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label31 
            Caption         =   "cm"
            Height          =   255
            Left            =   4440
            TabIndex        =   65
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label30 
            Caption         =   "Lámina Maxima"
            Height          =   255
            Left            =   2880
            TabIndex        =   64
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label25 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   2760
            TabIndex        =   63
            Top             =   -240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1575
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   7695
         Begin VB.TextBox txtpl 
            Height          =   285
            Left            =   5520
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtal 
            Height          =   285
            Left            =   5520
            TabIndex        =   13
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtpmpl 
            Height          =   285
            Left            =   1920
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtpel 
            Height          =   285
            Left            =   1920
            TabIndex        =   9
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtccl 
            Height          =   285
            Left            =   1920
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Evaluar"
            Height          =   375
            Left            =   5280
            TabIndex        =   15
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Profundidad"
            Height          =   375
            Left            =   3840
            TabIndex        =   71
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label36 
            Caption         =   "Agotamiento "
            Height          =   375
            Left            =   3840
            TabIndex        =   70
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label35 
            Caption         =   "%"
            Height          =   255
            Left            =   6840
            TabIndex        =   69
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label34 
            Caption         =   "cm"
            Height          =   255
            Left            =   6840
            TabIndex        =   68
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "Capacidad Campo"
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Pto. Marquitez Perma."
            Height          =   375
            Left            =   240
            TabIndex        =   58
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Peso Específico"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "%"
            Height          =   255
            Left            =   3240
            TabIndex        =   56
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "%"
            Height          =   255
            Left            =   3240
            TabIndex        =   55
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "gr/cc"
            Height          =   255
            Left            =   3240
            TabIndex        =   54
            Top             =   1200
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Capacidad de Campo (C.C) y Punto de Marquitez Permanente (P.M.P)"
      ForeColor       =   &H00800000&
      Height          =   3375
      Left            =   240
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   8055
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   7695
         Begin VB.CommandButton bccpmp 
            Caption         =   "&Evaluar"
            Height          =   615
            Left            =   5280
            Picture         =   "frmgenerales.frx":2288
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtArena 
            Height          =   285
            Left            =   1920
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtLimo 
            Height          =   285
            Left            =   1920
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtArcilla 
            Height          =   285
            Left            =   5520
            TabIndex        =   1
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "%"
            Height          =   255
            Left            =   3240
            TabIndex        =   43
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "%"
            Height          =   255
            Left            =   3240
            TabIndex        =   42
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "%"
            Height          =   255
            Left            =   6840
            TabIndex        =   41
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Limo"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Arcilla"
            Height          =   255
            Left            =   4080
            TabIndex        =   39
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Arena"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   7695
         Begin VB.TextBox txtpmp 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2880
            TabIndex        =   31
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtcc 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   2760
            TabIndex        =   36
            Top             =   -240
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Pto. Marquitez Permanente"
            Height          =   255
            Left            =   2880
            TabIndex        =   35
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            Height          =   255
            Left            =   4440
            TabIndex        =   34
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label6 
            Caption         =   "Capacidad de Campo"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "%"
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   720
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Humedad en el Suelo"
      ForeColor       =   &H00800000&
      Height          =   3375
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Width           =   8055
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   7695
         Begin VB.CommandButton bhumedad 
            Caption         =   "&Evaluar"
            Height          =   615
            Left            =   5280
            Picture         =   "frmgenerales.frx":29F2
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtpss 
            Height          =   285
            Left            =   1920
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtpsh 
            Height          =   285
            Left            =   5520
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtv 
            Height          =   285
            Left            =   1920
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Peso del Suelo Seco"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Peso del Suelo húmedo"
            Height          =   255
            Left            =   3720
            TabIndex        =   50
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Volumen"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "gr"
            Height          =   255
            Left            =   3240
            TabIndex        =   48
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label56 
            Caption         =   "cm3"
            Height          =   255
            Left            =   3240
            TabIndex        =   47
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label57 
            Caption         =   "gr"
            Height          =   255
            Left            =   6840
            TabIndex        =   46
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   7695
         Begin VB.TextBox txtpe 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txthv 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   5400
            TabIndex        =   19
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txthg 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2880
            TabIndex        =   18
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "gr/cm3"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Peso Específico"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   2760
            TabIndex        =   25
            Top             =   -240
            Width           =   1575
         End
         Begin VB.Label Label26 
            Caption         =   "% Humedad Volumétrica"
            Height          =   255
            Left            =   5400
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label27 
            Caption         =   "%"
            Height          =   255
            Left            =   6960
            TabIndex        =   23
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label28 
            Caption         =   "% Humedad Gravimétrica"
            Height          =   255
            Left            =   2880
            TabIndex        =   22
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label29 
            Caption         =   "%"
            Height          =   255
            Left            =   4440
            TabIndex        =   21
            Top             =   720
            Width           =   255
         End
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Left            =   5400
      Picture         =   "frmgenerales.frx":315C
      Top             =   4320
      Width           =   2700
   End
   Begin VB.Label lbltitulo 
      Caption         =   "Parámetros Generales del Suelo"
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
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mopar 
      Caption         =   "Otros Parámetros"
      Begin VB.Menu mtecx 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcon 
         Caption         =   "Conductividad Hidráulica"
      End
   End
   Begin VB.Menu masismat 
      Caption         =   "Evapotranspiración"
   End
   Begin VB.Menu mm 
      Caption         =   "Menu Principal"
   End
End
Attribute VB_Name = "frmgenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub bccpmp_Click()
On Error GoTo mensaje
arena = Val(txtArena.text)
arcilla = Val(txtArcilla.text)
limo = Val(txtLimo.text)
If arena + arcilla + limo <> 100 Then
MsgBox "la suma de las partículas debe ser igual a 100", 64, "Determinación de la Lámina"
txtArena.SetFocus
Exit Sub
End If

cc = 0.6382 * arcilla + 0.2845 * limo + 0.0507 * arena
pmp = 0.4897 * cc + 1.16666
txtcc.text = Format(cc, "##0.0##")
txtpmp.text = Format(pmp, "##0.0##")
txtccl.text = Format(cc, "##0.0##")
txtpmpl.text = Format(pmp, "##0.0##")
Exit Sub
mensaje:
 MsgBox "Ingrese adecuadamente los datos", 64, "Estimación de CC y PMP"

End Sub

Private Sub bfinailizar_Click()
Unload Me

End Sub

Private Sub bhumedad_Click()
On Error GoTo mensaje
pss = Val(txtpss.text)
psh = Val(txtpsh.text)
V = Val(Txtv.text)
If pss = 0 Then
MsgBox "Ingrese el valor de Pss", 64, "Determinación de la Lámina"
txtpss.SetFocus
Exit Sub
End If
If psh = 0 Then
MsgBox "Ingrese el valor de Psh", 64, "Determinación de la Lámina"
txtpsh.SetFocus
Exit Sub
End If
If V = 0 Then
MsgBox "Ingrese el valor de V", 64, "Determinación de la Lámina"
Txtv.SetFocus
Exit Sub
End If
If pss > psh Then
MsgBox "El Pss no puede ser mayor al Psh", 64, "Determinación de la Lámina"
txtpss.SetFocus
Exit Sub
End If

pe = pss / V
hg = (psh - pss) / pss * 100
hv = hg * pe

txtpe.text = Format(pe, "##0.0##")
txtpel.text = Format(pe, "##0.0##")

txthg.text = Format(hg, "##0.0##")
txthv.text = Format(hv, "##0.0##")
Exit Sub
mensaje:
     MsgBox "Ingrese adecuadamente los datos", 64, "Cálculo de %HG y %HV"


End Sub



Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtpss.text = ""
txtpsh.text = ""
Txtv.text = ""
txtpe.text = ""
txthg.text = ""
txthv.text = ""
txtArcilla.text = ""
txtArena.text = ""
txtLimo.text = ""
txtpmp.text = ""
txtcc.text = ""
txtccl.text = ""
txtpmpl.text = ""
txtpl.text = ""
txtal.text = ""
txtpel.text = ""
txtlnl.text = ""
txtlml.text = ""


End Sub


Private Sub Command1_Click()
On Error GoTo mensaje
cc = Val(txtccl.text)
pmp = Val(txtpmpl.text)
pe = Val(txtpel.text)
a = Val(txtal.text)
P = Val(txtpl.text)
If cc = 0 Then
MsgBox "Ingrese el valor de CC", 64, "Determinación de la Lámina"
txtccl.SetFocus
Exit Sub
End If
If pmp = 0 Then
MsgBox "Ingrese el valor de PMP", 64, "Determinación de la Lámina"
txtpmpl.SetFocus
Exit Sub
End If
If pe = 0 Then
MsgBox "Ingrese el valor de Peso específico", 64, "Determinación de la Lámina"
txtpel.SetFocus
Exit Sub
End If
If P = 0 Then
MsgBox "Ingrese el valor de la profundidad de raices", 64, "Determinación de la Lámina"
txtpl.SetFocus
Exit Sub
End If
If a = 0 Then
MsgBox "Ingrese el valor del agotamiento", 64, "Determinación de la Lámina"
txtal.SetFocus
Exit Sub
End If
If cc <= pmp Then
MsgBox "La CC es mayor al PMP ", 64, "Determinación de la Lámina"
txtccl.SetFocus
Exit Sub
End If

lm = ((cc - pmp) / 100) * pe * P
ln = lm * a / 100

txtlnl.text = Format(ln, "##0.0##")
txtlml.text = Format(lm, "##0.0##")
Exit Sub
mensaje:
     MsgBox "Ingrese adecuadamente los datos", 64, "Determinación de la Lámina"

End Sub

Private Sub Form_Load()
StatusBar1.Panels(1).text = " Cálculo de porcentaje de Humedad gravimetrica y volumétrica de una muestra de suelo"
End Sub




Private Sub masismat_Click()
frmETO.Show
End Sub

Private Sub mcon_Click()
frmconductividad.Show
End Sub

Private Sub mm_Click()
frmGeneral.Show
Unload Me
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

Private Sub mtecx_Click()
frmtextura.Show

End Sub

Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    Frame5.Visible = True
    txtpss.SetFocus
    Frame6.Visible = False
    Frame7.Visible = False
    Case 2
    Frame6.Visible = True
    Frame5.Visible = False
    Frame7.Visible = False
    txtArena.SetFocus
    Case 3
    Frame6.Visible = False
    Frame5.Visible = False
    Frame7.Visible = True
    txtccl.SetFocus
End Select
If Frame5.Visible = True Then
     StatusBar1.Panels(1).text = "Cálculo de porcentaje de Humedad gravimetrica y volumétrica de una muestra de suelo"
Else
    If Frame6.Visible = True Then
        StatusBar1.Panels(1).text = "Estimación del porcentaje de C.C y P.M.P a partir de la granulometría de la muestra"
    Else
        StatusBar1.Panels(1).text = "Cálculo de la Lámina Máxima y Neta del Suelo"
    End If
End If
End Sub
