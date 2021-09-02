VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmfertirrigacion 
   Caption         =   "Fertirrigación"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmfertirrigacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   3240
      TabIndex        =   23
      Top             =   6240
      Width           =   5895
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   4080
         Picture         =   "frmfertirrigacion.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2160
         Picture         =   "frmfertirrigacion.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   240
         Picture         =   "frmfertirrigacion.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requerimiento de fertilizantes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Compatibilidad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Caudales y tiempos de inyección"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
            TextSave        =   "29/08/2007"
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
   Begin VB.Frame ffer 
      Caption         =   "Fertilización"
      ForeColor       =   &H00800000&
      Height          =   4815
      Left            =   240
      TabIndex        =   26
      Top             =   1560
      Width           =   8175
      Begin VB.TextBox Text3 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6480
         TabIndex        =   48
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Líquido"
         Height          =   195
         Left            =   4680
         TabIndex        =   43
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Granular"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3360
         TabIndex        =   42
         Top             =   2040
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   615
         Left            =   1200
         TabIndex        =   34
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
      End
      Begin VB.TextBox txtAR 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNR 
         Height          =   285
         Left            =   5400
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton bCalFER 
         Caption         =   "Calcular"
         Height          =   615
         Left            =   6240
         Picture         =   "frmfertirrigacion.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   615
         Left            =   600
         TabIndex        =   35
         Top             =   2280
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid3 
         Height          =   1455
         Left            =   120
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   5
         Cols            =   6
         BackColorFixed  =   8438015
         ForeColorFixed  =   8388608
         GridColor       =   8438015
      End
      Begin VB.Label Label4 
         Caption         =   "Kg de Fertilizante"
         Height          =   375
         Left            =   6480
         TabIndex        =   47
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Indicar las características de la fuente"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "has"
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Area de riego"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Requerimientos de elementos en Kg/ha"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label10 
         Caption         =   "Número de riegos"
         Height          =   375
         Left            =   3960
         TabIndex        =   30
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fcau 
      Caption         =   "Caudales y tiempos de inyección"
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   840
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtCI 
         Height          =   285
         Left            =   2400
         TabIndex        =   44
         Text            =   "100"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "¿Calcular...?"
         Height          =   615
         Left            =   4080
         TabIndex        =   19
         Top             =   120
         Width           =   2535
         Begin VB.OptionButton Option2 
            Caption         =   "Caudal"
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Tiempo"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4440
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtVr 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCF 
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calcular"
         Height          =   615
         Left            =   1080
         Picture         =   "frmfertirrigacion.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label43 
         Caption         =   "Concentración inicial"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label39 
         Caption         =   "%"
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label LU1 
         Caption         =   "Hrs"
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label LU 
         Caption         =   "lts/h"
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label LCT 
         Caption         =   "Caudal de inyección"
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "lts"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label37 
         Caption         =   "Volumen del recipiente"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label42 
         Caption         =   "Concentración final"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   -720
         X2              =   6840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label22 
         Caption         =   "%"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.Label LCT1 
         Caption         =   "Tiempo de inyección"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Frame fcom 
      Caption         =   "Compatibilidad química "
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "Compatibilidad"
         Height          =   615
         Left            =   720
         Picture         =   "frmfertirrigacion.frx":315C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TC 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "VRS"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2325
      Left            =   8640
      Picture         =   "frmfertirrigacion.frx":38C6
      Top             =   2160
      Width           =   2940
   End
   Begin VB.Label Label17 
      Caption         =   "Cálculos en fertirrigación"
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
      Left            =   840
      TabIndex        =   24
      Top             =   480
      Width           =   4575
   End
   Begin VB.Menu mag 
      Caption         =   "Diseño agronómico"
      Begin VB.Menu mgo 
         Caption         =   "Goteo"
      End
      Begin VB.Menu mas 
         Caption         =   "Aspersión"
      End
      Begin VB.Menu mm 
         Caption         =   "Micro- aspersión"
      End
   End
   Begin VB.Menu mcal 
      Caption         =   "Calendario de riego"
   End
   Begin VB.Menu mmp 
      Caption         =   "Menú principal"
   End
End
Attribute VB_Name = "frmfertirrigacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim row, col, numero, num(0 To 1000)
Dim n As Integer, i, ii, punto

Private Sub bCalFER_Click()
On Error GoTo mensaje:
ar = Val(txtAR.text)
nr = Val(txtNR.text)
If ar = 0 Then
MsgBox "Ingrese el área a fertilizar", 64, "Fertirriego"
txtAR.SetFocus
Exit Sub
End If
If nr = 0 Then
MsgBox "Ingrese la cantidad de riegos que se usará para fertilizar", 64, "Fertirriego"
txtNR.SetFocus
Exit Sub
End If
If (Val(grid2.TextMatrix(1, 4))) = 0 Then
MsgBox "Ingrese las características del fertilizante", 64, "Fertirriego"
grid2.SetFocus
Exit Sub
End If
'requerimieto de fertilizantes
NN = Val(Grid1.TextMatrix(1, 0))
pp = Val(Grid1.TextMatrix(1, 1))
KK = Val(Grid1.TextMatrix(1, 2))
oo = Val(Grid1.TextMatrix(1, 3))

'Caracteristicas de la fuentes
nn1 = Val(grid2.TextMatrix(1, 0))
pp1 = Val(grid2.TextMatrix(1, 1))
kk1 = Val(grid2.TextMatrix(1, 2))
oo1 = Val(grid2.TextMatrix(1, 3))
s1 = Val(grid2.TextMatrix(1, 4))

'Calculo del peso por riego

rnn = NN / nr * ar
rpp = pp / nr * ar
rkk = KK / nr * ar
roo = oo / nr * ar

'Calculo de la fuente demandante Nitrogeno
If rnn >= rpp And rnn >= rkk And rnn >= roo Then
    kgnn = rnn * 100 / nn1
    kgnnn = kgnn * nn1 / 100
    kgKK = kgnn * kk1 / 100
    kgPP = kgnn * pp1 / 100
    kgoo = kgnn * oo1 / 100
    agua = kgnn / s1
    If Option3.Value = True Then
    Text3.text = Format(kgnn, "##0.0##")
    Else
    Text3.text = Format(agua, "##0.0##")
    End If
    'defict kg
    dnn = rnn - kgnnn
    dpp = rpp - kgPP
    dkk = rkk - kgKK
    doo = roo - kgoo
    'defict kg/ha
    ddnn = (rnn - kgnnn) / ar * nr
    ddpp = (rpp - kgPP) / ar * nr
    ddkk = (rkk - kgKK) / ar * nr
    ddoo = (roo - kgoo) / ar * nr
    kgnn = kgnnn
End If

'Calculo de la fuente demandante fosoforo
If rpp > rnn And rpp > rkk And rpp > roo Then
    
    kgPP = rpp * 100 / pp1
    kgppp = kgPP * pp1 / 100
    kgnn = kgPP * nn1 / 100
    kgKK = kgPP * kk1 / 100
    kgoo = kgPP * oo1 / 100
    agua = kgPP / s1
    If Option3.Value = True Then
    Text3.text = Format(kgPP, "##0.0##")
    Else
    Text3.text = Format(agua, "##0.0##")
    End If
    'defict kg
    dnn = rnn - kgnn
    dpp = rpp - kgppp
    dkk = rkk - kgKK
    doo = roo - kgoo
    'defict kg/ha
    ddnn = (rnn - kgnn) / ar * nr
    ddpp = (rpp - kgppp) / ar * nr
    ddkk = (rkk - kgKK) / ar * nr
    ddoo = (roo - kgoo) / ar * nr
    kgPP = kgppp
End If



'Calculo de la fuente demandante potasio
If rkk > rnn And rkk > rpp And rkk > roo Then

    kgKK = rkk * 100 / kk1
    kgKKK = kgKK * kk1 / 100
    kgnn = kgKK * nn1 / 100
    kgPP = kgKK * pp1 / 100
    kgoo = kgKK * oo1 / 100
    agua = kgKK / s1
    If Option3.Value = True Then
    Text3.text = Format(kgKK, "##0.0##")
    Else
    Text3.text = Format(agua, "##0.0##")
    End If
    'defict kg
    dnn = rnn - kgnn
    dpp = rpp - kgPP
    dkk = rkk - kgKKK
    doo = roo - kgoo
    'defict kg/ha
    ddnn = (rnn - kgnn) / ar * nr
    ddpp = (rpp - kgPP) / ar * nr
    ddkk = (rkk - kgKKK) / ar * nr
    ddoo = (roo - kgoo) / ar * nr
    kgKK = kgKKK
End If

'Calculo de la fuente demandante otro
If roo > rnn And roo > rpp And roo > rkk Then

    kgoo = roo * 100 / oo1
    kgooo = kgoo * oo1 / 100
    kgnn = kgoo * nn1 / 100
    kgPP = kgoo * pp1 / 100
    kgKK = kgoo * kk1 / 100
    agua = kgoo / s1
    If Option3.Value = True Then
    Text3.text = Format(kgoo, "##0.0##")
    Else
    Text3.text = Format(agua, "##0.0##")
    End If
    'defict kg
    dnn = rnn - kgnn
    dpp = rpp - kgPP
    dkk = rkk - kgKK
    doo = roo - kgooo
    'defict kg/ha
    ddnn = (rnn - kgnn) / ar * nr
    ddpp = (rpp - kgPP) / ar * nr
    ddkk = (rkk - kgKK) / ar * nr
    ddoo = (roo - kgooo) / ar * nr
    kgoo = kgooo
End If
Label4.Visible = True
Text3.Visible = True
With grid3
        .TextMatrix(1, 1) = Format(rnn, "##0.0##")
        .TextMatrix(1, 2) = Format(rpp, "##0.0##")
        .TextMatrix(1, 3) = Format(rkk, "##0.0##")
        .TextMatrix(1, 4) = Format(roo, "##0.0##")
        .TextMatrix(2, 1) = Format(kgnn, "##0.0##")
        .TextMatrix(2, 2) = Format(kgPP, "##0.0##")
        .TextMatrix(2, 3) = Format(kgKK, "##0.0##")
        .TextMatrix(2, 4) = Format(kgoo, "##0.0##")
        .TextMatrix(2, 5) = Format(agua, "##0.0##")
        .TextMatrix(3, 1) = Format(dnn, "##0.0##")
        .TextMatrix(3, 2) = Format(dpp, "##0.0##")
        .TextMatrix(3, 3) = Format(dkk, "##0.0##")
        .TextMatrix(3, 4) = Format(doo, "##0.0##")
        .TextMatrix(4, 1) = Format(ddnn, "##0.0##")
        .TextMatrix(4, 2) = Format(ddpp, "##0.0##")
        .TextMatrix(4, 3) = Format(ddkk, "##0.0##")
        .TextMatrix(4, 4) = Format(ddoo, "##0.0##")
        grid3.Visible = True
    End With
Exit Sub
mensaje:
MsgBox "Ingrese valores adecuados", 64, "Fertirriego"
End Sub

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
TC.text = ""
txtVr.text = ""
txtCI.text = ""
txtCF.text = ""
Text1.text = ""
Text2.text = ""
txtAR.text = ""
txtNR.text = ""
Label4.Visible = False
Text3.text = ""
Text3.Visible = False
For j% = 0 To 4
grid2.TextMatrix(1, j%) = ""
Next j%

For j% = 0 To 3
Grid1.TextMatrix(1, j%) = ""
Next j%

For j% = 1 To 5
For k% = 1 To 4
grid3.TextMatrix(k%, j%) = ""
Next k%
Next j%
grid3.Visible = False
End Sub

Private Sub Command1_Click()
On Error GoTo mensaje

aa = Val(Combo1.ListIndex)
bb = Val(Combo2.ListIndex)
If aa = -1 Then
    MsgBox "Ingrese el producto a X", 64, "fertirrigación"
    Combo1.SetFocus
    Exit Sub
End If
If bb = -1 Then
    MsgBox "Ingrese el producto a Y", 64, "fertirrigación"
    Combo2.SetFocus
    Exit Sub
End If

Select Case aa
Case 0
    Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "incompatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "compatible"
    End Select
Case 1
    Select Case bb
    Case 0
        TC = "incompatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compabilidad limitada"
    Case 4
        TC = "compabilidad limitada"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "compatible"
    End Select
Case 2
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "incompatible"
    End Select
Case 3
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compabilidad limitada"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compabilidad limitada"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "incompatible"
    End Select
Case 4
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compabilidad limitada"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compabilidad limitada"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "incompatible"
    End Select
Case 5
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compabilidad limitada"
    Case 4
        TC = "compabilidad limitada"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "incompatible"
    End Select
Case 6
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "incompatible"
    End Select
Case 7
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "compatible"
    End Select
Case 8
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "incompatible"
    End Select
Case 9
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "compatible"
    Case 3
        TC = "compatible"
    Case 4
        TC = "compatible"
    Case 5
        TC = "compatible"
    Case 6
        TC = "compatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "compatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "compatible"
    End Select
Case 10
Select Case bb
    Case 0
        TC = "compatible"
    Case 1
        TC = "compatible"
    Case 2
        TC = "incompatible"
    Case 3
        TC = "incompatible"
    Case 4
        TC = "incompatible"
    Case 5
        TC = "incompatible"
    Case 6
        TC = "incompatible"
    Case 7
        TC = "compatible"
    Case 8
        TC = "incompatible"
    Case 9
        TC = "compatible"
    Case 10
        TC = "compatible"
    End Select
End Select
Exit Sub
mensaje:
    MsgBox "Ingrese adecuadamente los datos de Entrada", 64, " Fertirrigación"

End Sub

Private Sub Command2_Click()
On Error GoTo mensaje
vr = Val(txtVr.text)
ci = Val(txtCI.text)
cf = Val(txtCF.text)
tc1 = Val(Text2.text)
If vr = 0 Then
    MsgBox "Ingrese el volumen del recipiente", 64, "fertirrigación"
    txtVr.SetFocus
    Exit Sub
End If

If tc1 = 0 Then
    If Option1.Value = True Then
    MsgBox "Ingrese el valor del caudal de inyección", 64, "fertirrigación"
    Else
    MsgBox "Ingrese el valor del tiempo de inyección", 64, "fertirrigación"
    End If
    Text2.SetFocus
    Exit Sub
    
End If
If cf = 0 Then
    tc2 = 4 * vr / tc1
Else
    If ci = 0 Then
        MsgBox "Ingrese el valor de la concentración inicial", 64, "fertirrigación"
        txtCI.SetFocus
        Exit Sub
    End If
    tc2 = -vr * (Log(cf / ci)) / tc1
End If
Text1.text = Format(tc2, "##0.00##")
Exit Sub
mensaje:
    MsgBox "Ingrese adecuadamente los datos ", 64, " Fertirrigación"


End Sub

Private Sub Form_Load()
With Combo1
    .AddItem "Nitrato amoniaco"
    .AddItem "Urea"
    .AddItem "Sulfato amoniaco"
    .AddItem "Superfosfato triple"
    .AddItem "Superfosfato simple"
    .AddItem "Fosfato diamoniaco"
    .AddItem "Fosfato monoamoniaco"
    .AddItem "Cloruro potásico"
    .AddItem "Sulfato potásico"
    .AddItem "nitrato potásico"
    .AddItem "Nitrato cálcico"
End With
With Combo2
    .AddItem "Nitrato amoniaco"
    .AddItem "Urea"
    .AddItem "Sulfato amoniaco"
    .AddItem "Superfosfato triple"
    .AddItem "Superfosfato simple"
    .AddItem "Fosfato diamoniaco"
    .AddItem "Fosfato monoamoniaco"
    .AddItem "Cloruro potásico"
    .AddItem "Sulfato potásico"
    .AddItem "nitrato potásico"
    .AddItem "Nitrato cálcico"
End With
With Grid1
    .TextMatrix(0, 0) = "Nitrógeno"
    .TextMatrix(0, 1) = "Fósforo"
    .TextMatrix(0, 2) = "Potasio"
    .TextMatrix(0, 3) = "Otro"
End With
With grid2
    .TextMatrix(0, 0) = "Nitrógeno %"
    .TextMatrix(0, 1) = "Fósforo %"
    .TextMatrix(0, 2) = "Potasio %"
    .TextMatrix(0, 3) = "Otro %"
    .ColWidth(4) = 1500
    .TextMatrix(0, 4) = "Solubilidad gr/cm3"
End With
With grid3
    .ColWidth(0) = 2700
    .ColWidth(5) = 1200
    
    .TextMatrix(0, 1) = "Nitrógeno "
    .TextMatrix(0, 2) = "Fósforo "
    .TextMatrix(0, 3) = "Potasio "
    .TextMatrix(0, 4) = "Otro "
    .TextMatrix(0, 5) = "Agua lts"
    .TextMatrix(1, 0) = "Requerimos nutrimentos kg/riego"
    .TextMatrix(2, 0) = "Aportamos con la fuente kg/riego"
    .TextMatrix(3, 0) = "Déficit elementos kg/riego"
    .TextMatrix(4, 0) = "Déficit elementos kg/has"
End With
StatusBar1.Panels(1).text = "Ingrese los datos de requerimiento y del tipo de fertilizante"
End Sub




Private Sub Grid1_Click()
i = ""
punto = 0
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)

If Grid1.col <> col Or Grid1.row <> row Then
i = ""
punto = 0
For s = 1 To numero
num(s) = ""
Next s
numero = 0
End If
num(0) = ""
If KeyAscii = 45 And i = "" Then
i = i + "-"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 48 Then
    i = i + "0"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If


If punto <> 1 Then
If KeyAscii = 44 Or KeyAscii = 46 Then
    numero = numero + 1
    If i = "" Then
    i = i + "0."
    Grid1.text = i
    num(numero) = i
    punto = 1
Else
    i = i + "."
    Grid1.text = i
    num(numero) = i
    punto = 1
End If
End If
End If


If KeyAscii = 49 Then
    i = i + "1"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 50 Then
    i = i + "2"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 51 Then
    i = i + "3"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 52 Then
    i = i + "4"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 53 Then
    i = i + "5"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 54 Then
    i = i + "6"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 55 Then
    i = i + "7"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 56 Then
    i = i + "8"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 57 Then
    i = i + "9"
    Grid1.text = i
    numero = numero + 1
    num(numero) = i
End If
Rem tecla de borrado
If numero >= 1 Then
If KeyAscii = 8 Then
i = num(numero - 1)
numero = numero - 1
Grid1.text = i
End If
Else
Grid1.text = ""
End If

Rem pruebas grid1.TextMatrix(numero, 6) = num(numero)

Rem grid1.Text = KeyAscii
col = Grid1.col
row = Grid1.row

End Sub





Private Sub mas_Click()
frmDAaspersion.Show
End Sub

Private Sub mcal_Click()
frmCalendario.Show
End Sub

Private Sub mgo_Click()
frmDAgoteo.Show
End Sub

Private Sub mm_Click()
frmDAMicro.Show
End Sub

Private Sub mmp_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub Option1_Click()
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
LCT.Caption = "Caudal de inyección"
LU.Caption = "lts/hr"
LCT1.Caption = "Tiempo de inyección"
LU1.Caption = "hrs"
End Sub
Private Sub Option2_Click()
Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
LCT1.Caption = "Caudal de inyección"
LU1.Caption = "lts/hr"
LCT.Caption = "Tiempo de inyección"
LU.Caption = "hrs"
End Sub
Private Sub Grid2_Click()
i = ""
punto = 0
End Sub
Private Sub grid2_KeyPress(KeyAscii As Integer)

If grid2.col <> col Or grid2.row <> row Then
i = ""
punto = 0
For s = 1 To numero
num(s) = ""
Next s
numero = 0
End If
num(0) = ""
If KeyAscii = 45 And i = "" Then
i = i + "-"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 48 Then
    i = i + "0"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If


If punto <> 1 Then
If KeyAscii = 44 Or KeyAscii = 46 Then
    numero = numero + 1
    If i = "" Then
    i = i + "0."
    grid2.text = i
    num(numero) = i
    punto = 1
Else
    i = i + "."
    grid2.text = i
    num(numero) = i
    punto = 1
End If
End If
End If


If KeyAscii = 49 Then
    i = i + "1"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 50 Then
    i = i + "2"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 51 Then
    i = i + "3"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 52 Then
    i = i + "4"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 53 Then
    i = i + "5"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 54 Then
    i = i + "6"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 55 Then
    i = i + "7"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 56 Then
    i = i + "8"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 57 Then
    i = i + "9"
    grid2.text = i
    numero = numero + 1
    num(numero) = i
End If
Rem tecla de borrado
If numero >= 1 Then
If KeyAscii = 8 Then
i = num(numero - 1)
numero = numero - 1
grid2.text = i
End If
Else
grid2.text = ""
End If

Rem pruebas grid1.TextMatrix(numero, 6) = num(numero)

Rem grid1.Text = KeyAscii
col = grid2.col
row = grid2.row

End Sub

Private Sub Option3_Click()
Label4.Caption = "Kg de fertilizante"
grid2.TextMatrix(0, 4) = "Solubilidad gr/cm3"
grid3.TextMatrix(0, 5) = "Agua lts"
Option3.ForeColor = &HC0&
Option4.ForeColor = &H80000012
End Sub

Private Sub Option4_Click()
Label4.Caption = "litros de fertilizante"
grid2.TextMatrix(0, 4) = "Peso Esp. gr/cm3"
grid3.ColWidth(5) = 1200
grid3.TextMatrix(0, 5) = "Fertilizante lts"
Option4.ForeColor = &HC0&
Option3.ForeColor = &H80000012
End Sub

Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    ffer.Visible = True
    fcom.Visible = False
    fcau.Visible = False
    StatusBar1.Panels(1).text = "Ingrese los datos de requerimiento y del tipo de fertilizante"
    txtAR.SetFocus
    Case 2
    ffer.Visible = False
    fcom.Visible = True
    fcau.Visible = False
    
    StatusBar1.Panels(1).text = "Seleccione  el nombre de los fertilizantes a mezclar "
    Case 3
    ffer.Visible = False
    fcom.Visible = False
    fcau.Visible = True
    txtVr.SetFocus
    StatusBar1.Panels(1).text = "Ingrese los datos y oprima el botón Calcular"
End Select

End Sub
