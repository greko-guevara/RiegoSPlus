VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDAaspersion 
   Caption         =   "Diseño agronómico riego por aspersión"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   Icon            =   "frmDAaspersion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11850
   Begin VB.Frame Frame10 
      Height          =   1095
      Left            =   2760
      TabIndex        =   27
      Top             =   6360
      Width           =   6615
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   360
         Picture         =   "frmDAaspersion.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2520
         MaskColor       =   &H000000FF&
         Picture         =   "frmDAaspersion.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   4560
         Picture         =   "frmDAaspersion.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   1920
      Width           =   3615
      _ExtentX        =   6376
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
            Caption         =   "Selección de Aspersores"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   7770
      Width           =   11850
      _ExtentX        =   20902
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
            TextSave        =   "4/5/2000"
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
      Caption         =   "Diseño Agronómico"
      ForeColor       =   &H00800000&
      Height          =   4215
      Left            =   600
      TabIndex        =   12
      Top             =   1560
      Width           =   7335
      Begin VB.OptionButton Option1 
         Caption         =   "Bruta"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4320
         TabIndex        =   44
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Neta"
         Height          =   255
         Left            =   5640
         TabIndex        =   43
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Bevaluar 
         Caption         =   "&Evaluar"
         Height          =   615
         Left            =   5400
         Picture         =   "frmDAaspersion.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox tlb 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   34
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox ttr 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   33
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox tib 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   32
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox tfr 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   31
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox ta 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox tq 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox tln 
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox tef 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   9000
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "mm/h"
         Height          =   255
         Left            =   4920
         TabIndex        =   42
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Lámina bruta"
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "mm"
         Height          =   255
         Left            =   4920
         TabIndex        =   40
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Caudal por hectárea"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Tiempo de riego"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "hrs"
         Height          =   255
         Left            =   4920
         TabIndex        =   37
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label34 
         Caption         =   "Intensidad de aplicación"
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label43 
         Caption         =   "m3/h-ha"
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "m2"
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "Area de riego"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "m3/h"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "mm"
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Caudal del aspersor"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Lámina "
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblunidades 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lbletiqueta 
         Caption         =   "Eficiencia"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Recomendación del aspersor según datos de campo"
      ForeColor       =   &H00800000&
      Height          =   4215
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Bc 
         Caption         =   "Calcular"
         Height          =   615
         Left            =   5520
         MaskColor       =   &H008080FF&
         Picture         =   "frmDAaspersion.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox CS 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtIb 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cVV 
         Height          =   315
         Left            =   5520
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid grdd 
         Height          =   2535
         Left            =   480
         TabIndex        =   10
         Top             =   1560
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   10
         Cols            =   5
         FixedCols       =   2
         BackColorFixed  =   8438015
         ForeColorFixed  =   8388608
         GridColor       =   8438015
      End
      Begin VB.Label Label1 
         Caption         =   "Pendiente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label44 
         Caption         =   "mm/h"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label45 
         Caption         =   "Infiltración básica"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "Velocidad del viento"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2085
      Left            =   8160
      Picture         =   "frmDAaspersion.frx":315C
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   3210
   End
   Begin VB.Label Label17 
      Caption         =   "Diseño agronómico Riego por aspersión"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   720
      Width           =   6255
   End
   Begin VB.Menu epsc 
      Caption         =   "Parámetros Suelo- Clima"
      Begin VB.Menu mgesu 
         Caption         =   "General suelo"
      End
      Begin VB.Menu mtex 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcon 
         Caption         =   "Conductividad hidráulica"
      End
      Begin VB.Menu meva 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu mht 
      Caption         =   "Hidráulica de Tuberías"
      Begin VB.Menu calat 
         Caption         =   "Cálculo en el lateral"
      End
      Begin VB.Menu mcapro 
         Caption         =   "Cálculo en la principal"
      End
      Begin VB.Menu mbom 
         Caption         =   "Selección bomba"
      End
      Begin VB.Menu mcomdia 
         Caption         =   "Combinación de diámetros"
      End
   End
   Begin VB.Menu gddrrr 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmDAaspersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BC_Click()
On Error GoTo mensaje:

ib = Val(txtIb.text)
If ib = 0 Then
    MsgBox "Ingrese el valor del caudal de infiltración básica", 64, "Riego por aspersión"
    txtIb.SetFocus
    Exit Sub
End If
ff = CS.ListIndex
If ff = -1 Then
    MsgBox "Seleccione el valor de la pendiente", 64, "Riego por aspersión"
    CS.SetFocus
    Exit Sub
End If
Select Case ff
    Case 0
    f = 1
    Case 1
    f = 0.8
    Case 2
    f = 0.6
End Select
vv = cVV.ListIndex
If vv = -1 Then
    MsgBox "Seleccione el valor aprox. de la velocidad del viento", 64, "Riego por aspersión"
    cVV.SetFocus
    Exit Sub
End If
Select Case vv
    Case 0
    v1 = 0.65
    v2 = 0.65
    v3 = 0.65
    Case 1
    v1 = 0.6
    v2 = 0.5
    v3 = 0.65
    Case 2
    v1 = 0.5
    v2 = 0.4
    v3 = 0.6
    Case 3
    v1 = 0.4
    v2 = 0.4
    v3 = 0.5
    Case 4
    v1 = 0.3
    v2 = 0.3
    v3 = 0.4
End Select


ibb = ib * f

For j% = 1 To 9
    sq = ibb * grdD.TextMatrix(j%, 0) * grdD.TextMatrix(j%, 1) / 1000
    d1 = 1 / v1 * grdD.TextMatrix(j%, 0)
    grdD.TextMatrix(j%, 2) = Format(sq, "#0.0##")
    grdD.TextMatrix(j%, 3) = Format(d1, "#0.0##")
    grdD.TextMatrix(j%, 4) = Format(d1, "#0.0##")
Next j%

For j% = 2 To 8 Step 2
    d2 = 1 / v2 * grdD.TextMatrix(j%, 0)
    d3 = 1 / v3 * grdD.TextMatrix(j%, 1)
    grdD.TextMatrix(j%, 3) = Format(d2, "#0.0##")
    grdD.TextMatrix(j%, 4) = Format(d3, "#0.0##")
Next j%
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"

End Sub

Private Sub bevaluar_Click()
On Error GoTo mensaje:
q1 = Val(tQ.text)
ln1 = Val(tln.text)
ef1 = Val(tef.text)

a1 = Val(ta.text)
If q1 = 0 Then
    MsgBox "Ingrese el valor del caudal del aspersor", 64, "Riego por aspersión"
    tQ.SetFocus
    Exit Sub
End If
If a1 = 0 Then
    MsgBox "Ingrese el valor del área efectiva", 64, "Riego por aspersión"
    ta.SetFocus
    Exit Sub
End If
'If etr1 = 0 Then
'    MsgBox "Ingrese el valor de la evapotranspiración", 64, "Riego por aspersión"
 '   tetr.SetFocus
  '  Exit Sub
'End If
If ln1 = 0 Then
    MsgBox "Ingrese el valor de la lámina neta", 64, "Riego por aspersión"
    tln.SetFocus
    Exit Sub
End If
If Option1 = False Then
    If ef1 = 0 Then
        MsgBox "Ingrese el valor de la eficiencia", 64, "Riego por aspersión"
        tef.SetFocus
        Exit Sub
    End If
End If

ib = q1 / a1 * 1000
qhas = 10000 / a1 * q1
tfr.text = Format(qhas, "#0.0##")
 
fr1 = Val(tfr.text)
If Option1 = False Then
    lb1 = ln1 / ef1 * 100
Else
    lb1 = ln1
End If
tr = lb1 / ib

tib.text = Format(ib, "#0.0##")
tlb.text = Format(lb1, "#0.0##")
ttr.text = Format(tr, "#0.0##")

Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"

End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
txtIb.text = ""
cVV.text = ""
CS.text = ""
tQ.text = ""
ta.text = ""
tln.text = ""
tef.text = ""
tib.text = ""
tfr.text = ""
ttr.text = ""
tlb.text = ""
With grdD
    .Clear
    .TextMatrix(0, 0) = "E.asp"
    .TextMatrix(0, 1) = "E.lat"
    .TextMatrix(0, 2) = "Caudal (m3/h)"
    .TextMatrix(0, 3) = "Diá.máx. asp.(m)"
    .TextMatrix(0, 4) = "Diá.máx. lat.(m)"
    
    
    .TextMatrix(1, 0) = 6
    .TextMatrix(2, 0) = 6
    .TextMatrix(3, 0) = 12
    .TextMatrix(4, 0) = 12
    .TextMatrix(5, 0) = 15
    .TextMatrix(6, 0) = 12
    .TextMatrix(7, 0) = 18
    .TextMatrix(8, 0) = 18
    .TextMatrix(9, 0) = 24
    
    .TextMatrix(1, 1) = 6
    .TextMatrix(2, 1) = 12
    .TextMatrix(3, 1) = 12
    .TextMatrix(4, 1) = 15
    .TextMatrix(5, 1) = 15
    .TextMatrix(6, 1) = 18
    .TextMatrix(7, 1) = 18
    .TextMatrix(8, 1) = 24
    .TextMatrix(9, 1) = 24
    
End With
    

End Sub



Private Sub calat_Click()
FrmHLaterales.Show
End Sub

Private Sub Option1_Click()
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
tef.BackColor = &H80000016
tef.Enabled = False
End Sub

Private Sub Option2_Click()
Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
tef.Enabled = True
tef.BackColor = &H80000005

End Sub
Private Sub Form_Load()
With grdD
    .ColWidth(0) = 600
    .ColWidth(1) = 600
    .ColWidth(2) = 1600
    .ColWidth(3) = 1600
    .ColWidth(4) = 1600
    .TextMatrix(0, 0) = "E.asp"
    .TextMatrix(0, 1) = "E.lat"
    .TextMatrix(0, 2) = "Caudal (m3/h)"
    .TextMatrix(0, 3) = "Diá. asp.(m)"
    .TextMatrix(0, 4) = "Diá. lat.(m)"
    
    
    .TextMatrix(1, 0) = 6
    .TextMatrix(2, 0) = 6
    .TextMatrix(3, 0) = 12
    .TextMatrix(4, 0) = 12
    .TextMatrix(5, 0) = 15
    .TextMatrix(6, 0) = 12
    .TextMatrix(7, 0) = 18
    .TextMatrix(8, 0) = 18
    .TextMatrix(9, 0) = 24
    
    .TextMatrix(1, 1) = 6
    .TextMatrix(2, 1) = 12
    .TextMatrix(3, 1) = 12
    .TextMatrix(4, 1) = 15
    .TextMatrix(5, 1) = 15
    .TextMatrix(6, 1) = 18
    .TextMatrix(7, 1) = 18
    .TextMatrix(8, 1) = 24
    .TextMatrix(9, 1) = 24
    
End With
StatusBar1.Panels(1).text = "Ingrese los datos de entrada y oprima el botón de Calcular para ver los requerimientos de los aspersores"

With cVV
    .AddItem "Sin viento"
    .AddItem "Hasta 6 km/h"
    .AddItem "Hasta 12 km/h"
    .AddItem "Hasta 15 km/h"
    .AddItem "Mayor 15 km/h"
End With

With CS
    .AddItem "0 a 5 %"
    .AddItem "6 a 8 %"
    .AddItem "9 a 12 %"
End With

End Sub

Private Sub gddrrr_Click()
Unload Me
frmGeneral.Show
End Sub


Private Sub mbom_Click()
frmbomba.Show
End Sub

Private Sub mcapro_Click()
frmHprincipal.Show
End Sub

Private Sub mcomdia_Click()
frmcombDia.Show
End Sub

Private Sub mcon_Click()
frmconductividad.Show
End Sub

Private Sub meva_Click()
frmETO.Show
End Sub

Private Sub mgesu_Click()
frmgenerales.Show
End Sub

Private Sub mtex_Click()
frmtextura.Show
End Sub

Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    Frame1.Visible = True
    Frame4.Visible = False
    StatusBar1.Panels(1).text = "Digite los datos de entrada y oprima el botón de Evaluar"
    tQ.SetFocus
    Case 2
    Frame4.Visible = True
    Frame1.Visible = False
    txtIb.SetFocus
    StatusBar1.Panels(1).text = "Ingrese los datos de entrada y oprima el botón de Calcular para ver los requerimientos de los aspersores"
End Select

End Sub
