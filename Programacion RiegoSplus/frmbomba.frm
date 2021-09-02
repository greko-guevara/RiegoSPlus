VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmbomba 
   Caption         =   "Selección de la bomba"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmbomba.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para la selección"
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   773
      TabIndex        =   22
      Top             =   840
      Width           =   10335
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   9000
         TabIndex        =   32
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txthd 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   7800
         TabIndex        =   31
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtZd 
         Height          =   285
         Left            =   7800
         TabIndex        =   30
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   9000
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtHs 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   7800
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtZs 
         Height          =   285
         Left            =   7800
         TabIndex        =   27
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtPO 
         Height          =   285
         Left            =   3000
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtQ 
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3000
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmbomba.frx":0CCA
         Left            =   3000
         List            =   "frmbomba.frx":0CCC
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   5280
         X2              =   5280
         Y1              =   120
         Y2              =   2040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   5280
         X2              =   10800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "Pérdidas en la descarga"
         Height          =   255
         Left            =   5520
         TabIndex        =   48
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Delta de Z en la descarga"
         Height          =   255
         Left            =   5520
         TabIndex        =   47
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "m"
         Height          =   255
         Left            =   8760
         TabIndex        =   46
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "m"
         Height          =   255
         Left            =   9120
         TabIndex        =   45
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Pérdidas en la succión"
         Height          =   255
         Left            =   5520
         TabIndex        =   44
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Delta de Z en la succión"
         Height          =   255
         Left            =   5520
         TabIndex        =   43
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Presión al final de la descarga"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "m"
         Height          =   255
         Left            =   8760
         TabIndex        =   41
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "m"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "°C"
         Height          =   255
         Left            =   4320
         TabIndex        =   39
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "Elevación sobre el mar"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Caudal requerido"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Temperatura ambiente"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "m3/hr"
         Height          =   255
         Left            =   4320
         TabIndex        =   35
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "m.s.n.m."
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "m"
         Height          =   255
         Left            =   9120
         TabIndex        =   33
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   1193
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   9495
      Begin VB.TextBox txtPA 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtPV 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNPSH 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtCDB 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid grdd 
         Height          =   2295
         Left            =   4560
         TabIndex        =   20
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   9
         Cols            =   4
         BackColorFixed  =   8438015
         ForeColorFixed  =   8388608
         GridColor       =   8438015
      End
      Begin VB.Label Label19 
         Caption         =   "Recuerde: el NPSH disponible debe ser mayor al NPSH requerido"
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label Label43 
         Caption         =   "m"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.Label fsf 
         Caption         =   "Carga dinámica de la bomba"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label34 
         Caption         =   "Presión de vapor"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "m"
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "NPSH disponible"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "m"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Presión atmosférica"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "m"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   600
      TabIndex        =   4
      Top             =   6240
      Width           =   7215
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmbomba.frx":0CCE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   3720
         Picture         =   "frmbomba.frx":13B8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   5520
         Picture         =   "frmbomba.frx":1B22
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular"
         Height          =   735
         Left            =   120
         MaskColor       =   &H008080FF&
         Picture         =   "frmbomba.frx":228C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7935
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
            TextSave        =   "06/07/2005"
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
      Height          =   1650
      Left            =   8160
      Picture         =   "frmbomba.frx":29F6
      Top             =   6000
      Width           =   3120
   End
   Begin VB.Label Label17 
      Caption         =   "Selección de la bomba"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   3495
   End
   Begin VB.Menu m1 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu mcalla 
         Caption         =   "Cálculos en la lateral"
      End
      Begin VB.Menu mcalprin 
         Caption         =   "Cálculos en la principal"
      End
      Begin VB.Menu mcomdia 
         Caption         =   "Combinación de diámetros"
      End
      Begin VB.Menu mdispre 
         Caption         =   "Distribución de presiones"
      End
   End
   Begin VB.Menu msal 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmbomba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pv As Double
Dim pa As Double
Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
Combo1.text = ""
Combo2.text = ""
txtQ.text = ""
txtPO.text = ""
txtHs.text = ""
txtZs.text = ""
txthd.text = ""
txtZd.text = ""
txtPV.text = ""
txtPA.text = ""
txtCDB.text = ""
txtNPSH.text = ""
For j% = 1 To 8
For k% = 1 To 3
    grdD.TextMatrix(j%, k%) = ""
Next k%
Next j%
Frame2.Visible = False

End Sub

Private Sub Command1_Click()
On Error GoTo mensaje:
q = Val(txtQ.text)
po = Val(txtPO.text)
hs = Val(txtHs.text)
hd = Val(txthd.text)
zs = Val(txtZs.text)
zD = Val(txtZd.text)

If q = 0 Then
    MsgBox "Ingrese el valor del caudal", 64, "Selección de la bomba"
    txtQ.SetFocus
Exit Sub
End If
If po = 0 Then
    MsgBox "Ingrese la presión de descarga", 64, "Selección de la bomba"
    txtPO.SetFocus
Exit Sub
End If
temp = Val(Combo4.ListIndex)
Select Case temp
    Case 0
    pv = 0.06
    Case 1
    pv = 0.09
    Case 2
    pv = 0.12
    Case 3
    pv = 0.17
    Case 4
    pv = 0.25
    Case 5
    pv = 0.33
    Case 6
    pv = 0.44
    Case 7
    pv = 0.58
    Case 8
    pv = 0.76
End Select

ele = Val(Combo3.ListIndex)
Select Case ele
    Case 0
    pa = 10.32
    Case 1
    pa = 10.03
    Case 2
    pa = 9.73
    Case 3
    pa = 9.45
    Case 4
    pa = 9.17
    Case 5
    pa = 9.02
    Case 6
    pa = 8.64
    Case 7
    pa = 8.37
    Case 8
    pa = 8.11
    Case 9
    pa = 7.58
    Case 10
    pa = 7.05
End Select

z1 = Val(Combo1.ListIndex)
Select Case z1
    Case 0
        zs1 = (zs)
    Case 1
        zs1 = -(zs)
    Case 2
        zs1 = 0
End Select
z2 = Val(Combo2.ListIndex)
Select Case z2
    Case 0
        zd1 = (zD)
    Case 1
        zd1 = -(zD)
    Case 2
        zd1 = 0
End Select

'***********************

npsh = pa - (zs1 + pv + hs)
cdb = (po + hs + hd + zd1 + zs1) * 1.05
pothp = q * cdb / 274

For j% = 1 To 8
    grdD.TextMatrix(j%, 1) = q
    grdD.TextMatrix(j%, 2) = Format(pothp / Val(grdD.TextMatrix(j%, 0)) * 100, "##0.00")
    grdD.TextMatrix(j%, 3) = Format(pothp / Val(grdD.TextMatrix(j%, 0)) * 100 * 745.69987, "##0.00")
Next j%
txtPV = Format(pv, "##0.0#")
txtPA = Format(pa, "##0.0#")
txtNPSH = Format(npsh, "##0.0##")
txtCDB = Format(cdb, "##0.0##")
Frame2.Visible = True


Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub





Private Sub Form_Load()
With Combo1
    .AddItem "Sube"
    .AddItem "Baja"
    .AddItem "Nivel"
End With
With Combo2
    .AddItem "Sube"
    .AddItem "Baja"
    .AddItem "Nivel"
End With
StatusBar1.Panels(1).text = "Digite los datos básicos para la selección y oprima el botón de Calcular"
With grdD
    .ColWidth(0) = 1000
    .ColWidth(1) = 1200
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    
    .TextMatrix(0, 0) = "Eficiecia (%)"
    .TextMatrix(0, 1) = "Caudal(m3/hr)"
    .TextMatrix(0, 2) = "Potencia (Hp)"
    .TextMatrix(0, 3) = "Potencia (watts)"
    .TextMatrix(1, 0) = "       30"
    .TextMatrix(2, 0) = "       40"
    .TextMatrix(3, 0) = "       50"
    .TextMatrix(4, 0) = "       60"
    .TextMatrix(5, 0) = "       70"
    .TextMatrix(6, 0) = "       80"
    .TextMatrix(7, 0) = "       90"
    .TextMatrix(8, 0) = "      100"
End With

With Combo4
    .AddItem "0"
    st = 0
    For j% = 1 To 8
    st = st + 5
    .AddItem st
    Next j%
End With

With Combo3
    .AddItem "0"
    st = 0
    For j% = 1 To 8
    st = st + 250
    .AddItem st
    Next j%
    .AddItem "2500"
    .AddItem "3000"
End With
Combo3.ListIndex = 4
Combo4.ListIndex = 5



End Sub

Private Sub mcalla_Click()
FrmHLaterales.Show
End Sub

Private Sub mcalprin_Click()
frmHprincipal.Show
End Sub

Private Sub mcomdia_Click()
frmcombDia.Show
End Sub

Private Sub mdispre_Click()
frmPresiones.Show
End Sub

Private Sub msal_Click()
Unload Me
frmGeneral.Show
End Sub
