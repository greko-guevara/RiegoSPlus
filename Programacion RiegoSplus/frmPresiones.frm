VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPresiones 
   Caption         =   "Cálculo de presiones"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   Icon            =   "frmPresiones.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11850
   Begin VB.ComboBox CTG 
      Height          =   315
      Left            =   6960
      TabIndex        =   50
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   3293
      TabIndex        =   46
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "Solo lateral"
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sólido"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Tipo de sistema"
         Height          =   255
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   1733
      TabIndex        =   43
      Top             =   6360
      Width           =   8175
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular"
         Height          =   735
         Left            =   405
         MaskColor       =   &H008080FF&
         Picture         =   "frmPresiones.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   6120
         Picture         =   "frmPresiones.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   4200
         Picture         =   "frmPresiones.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2280
         Picture         =   "frmPresiones.frx":2308
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdd 
      Height          =   1815
      Left            =   353
      TabIndex        =   42
      Top             =   4320
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      BackColorFixed  =   8438015
      ForeColorFixed  =   8388608
      GridColor       =   8438015
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   8813
      TabIndex        =   38
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton OCMS 
         Caption         =   "Si"
         Height          =   255
         Left            =   1320
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OCMN 
         Caption         =   "No"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Combinó diámetros en la multiple"
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   6053
      TabIndex        =   34
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton OCLS 
         Caption         =   "Si"
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OCLN 
         Caption         =   "No"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Combinó diámetros en la lateral"
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   533
      TabIndex        =   20
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton OA 
         Caption         =   "Aspersión"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OG 
         Caption         =   "Goteo "
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo de riego"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para el diseño"
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   713
      TabIndex        =   14
      Top             =   2040
      Width           =   10455
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   9000
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txthM 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   7800
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtZM 
         Height          =   285
         Left            =   7800
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   9000
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtHl 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   7800
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtZL 
         Height          =   285
         Left            =   7800
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtQO 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtPO 
         Height          =   285
         Left            =   3000
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "m"
         Height          =   255
         Left            =   9120
         TabIndex        =   52
         Top             =   1200
         Width           =   375
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
         Caption         =   "Pérdidas en la multiple"
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Delta de Z en la multiple"
         Height          =   255
         Left            =   5520
         TabIndex        =   31
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "m"
         Height          =   255
         Left            =   8760
         TabIndex        =   30
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "m"
         Height          =   255
         Left            =   9120
         TabIndex        =   29
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Pérdidas en el lateral"
         Height          =   255
         Left            =   5520
         TabIndex        =   28
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Delta de Z en el lateral"
         Height          =   255
         Left            =   5520
         TabIndex        =   27
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Area Efectiva de riego por salida"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "m"
         Height          =   255
         Left            =   8760
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "m2"
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "m"
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "Caudal de Operación por salida"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Altura del elevador"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Presión de Operación"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "m"
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "m3/h"
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   45
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
            TextSave        =   "10/10/2007"
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
   Begin VB.Label Label21 
      Caption         =   "Tipo de gotero"
      Height          =   255
      Left            =   5640
      TabIndex        =   51
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   8633
      Picture         =   "frmPresiones.frx":29F2
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label17 
      Caption         =   "Cálculo de presiones, caudales e intensidades"
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
      TabIndex        =   44
      Top             =   360
      Width           =   6615
   End
   Begin VB.Menu bpsc 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu mgensu 
         Caption         =   "Cálculos en laterales"
      End
      Begin VB.Menu mcp 
         Caption         =   "Cálculos en principales"
      End
      Begin VB.Menu msb 
         Caption         =   "Selección de bombas"
      End
      Begin VB.Menu fff 
         Caption         =   "Combinación de díametros"
      End
   End
   Begin VB.Menu dflkdl 
      Caption         =   "Menú principal"
   End
End
Attribute VB_Name = "frmPresiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim po As Double
Dim qo As Double
Dim h As Double
Dim a As Double
Dim hl As Double
Dim zl1 As Double
Dim zm1 As Double
Dim hm As Double
Dim pel As Double
Dim EL As Double
Dim em As Double
Dim pem As Double
Dim d As Double
Dim pd As Double
Dim k As Double

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
On Error GoTo mensaje:
txtPO.text = ""
txtQo.text = ""
txtH.text = ""
txtA.text = ""
txtHl.text = ""
txtZL.text = ""
txthM.text = ""
txtZM.text = ""
With grdD
    .Cols = 5
    .Clear
    .TextMatrix(0, 1) = "Presión Piso  (m)"
    .TextMatrix(0, 2) = "Presión Aspersor (m)"
    .TextMatrix(0, 3) = "Caudal (m3/h)"
    .TextMatrix(0, 4) = "Intensidad  (mm/h)"
    .ColWidth(3) = 1600
    .ColWidth(4) = 1600
End With
Exit Sub
mensaje:
MsgBox "Error al borrado"
End Sub



Private Sub Command1_Click()
On Error GoTo mensaje:
If Option1.Value = True Then
If OA.Value = True Then
    po = Val(txtPO.text)
    qo = Val(txtQo.text)
    h = Val(txtH.text)
    a = Val(txtA.text)
    hl = Val(txtHl.text)
    zl1 = Val(txtZL.text)
    hm = Val(txthM.text)
    zm1 = Val(txtZM.text)
    If po = 0 Then
        MsgBox "Ingrese el valor de la presión de operación", 64, "Cálculo de presiones, presiones e intensidades"
        txtPO.SetFocus
        Exit Sub
    End If
    If qo = 0 Then
        MsgBox "Ingrese el valor del caudal de operación", 64, "Cálculo de presiones, presiones e intensidades"
        txtQo.SetFocus
        Exit Sub
    End If
        If a = 0 Then
        MsgBox "Ingrese el valor del área efectiva del aspersor", 64, "Cálculo de presiones, presiones e intensidades"
        txtA.SetFocus
        Exit Sub
    End If
    If hl = 0 Then
        MsgBox "Ingrese el valor de las pérdidas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
        txtHl.SetFocus
        Exit Sub
    End If
     'If zl1 = 0 Then
      '  MsgBox "Ingrese la diferencia de cotas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
       ' txtZL.SetFocus
        'Exit Sub
    'End If
    If hm = 0 Then
        MsgBox "Ingrese el valor de las pérdidas en el multiple", 64, "Cálculo de presiones, presiones e intensidades"
        txthM.SetFocus
        Exit Sub
    End If
   '   If zm1 = 0 Then
     '     MsgBox "Ingrese la diferencia de cotas en la múltiple", 64, "Cálculo de presiones, presiones e intensidades"
       '   txtZM.SetFocus
      '    Exit Sub
 '     End If
    z1 = Combo1.ListIndex
    If z1 = -1 Then
        MsgBox "Defina la condición topográfica del lateral", 64, "Cálculo de presiones, presiones e intensidades"
        Combo1.SetFocus
        Exit Sub
    End If
    Select Case z1
    Case 0
        zl = -(zl1)
    Case 1
        zl = (zl1)
    Case 2
        zl = 0
    End Select
    
    z2 = Combo2.ListIndex
    If z2 = -1 Then
        MsgBox "Defina la condición topográfica del múltiple", 64, "Cálculo de presiones, presiones e intensidades"
        Combo2.SetFocus
        Exit Sub
    End If
    Select Case z1
    Case 0
        zm = -(zm1)
    Case 1
        zm = (zm1)
    Case 2
        zm = 0
    End Select
    
    If OCLN.Value = True Then
        EL = po + 3 / 4 * hl + h - 0.38 * zl
    Else
        EL = po + 5 / 8 * hl + h - 0.38 * zl
    End If
    
    If OCMN.Value = True Then
        em = EL + 3 / 4 * hm - 0.38 * zm
    Else
        em = EL + 5 / 8 * hm - 0.38 * zm
    End If
    
    k = qo / Sqr(po)
    io = qo * 1000 / a
    O = po + h
    
    pel = EL - h
    qel = k * pel ^ 0.5
    iel = qel * 1000 / a
    
    pd = EL - hl - h + zl
    qd = k * pd ^ 0.5
    id = qd * 1000 / a
    d = pd + h
    
    
    pem = em - h
    qem = k * pem ^ 0.5
    iem = qem * 1000 / a
    
    
    With grdD
    
        .Visible = True
        .Cols = 5
        .Rows = 5
        .ColWidth(3) = 1600
        .ColWidth(4) = 1600
        .TextMatrix(0, 1) = "Pr. Piso (m)"
        .TextMatrix(0, 2) = "Pr. Aspersor (m)"
        .TextMatrix(0, 3) = "Caudal(m3/h)"
        .TextMatrix(0, 4) = "Intensidad (mm/h)"
        
        .TextMatrix(1, 0) = "Aspersor operación"
        .TextMatrix(2, 0) = "Entrada lateral"
        .TextMatrix(3, 0) = "Aspersor distal"
        .TextMatrix(4, 0) = "Entrada multiple"
    
        .TextMatrix(1, 1) = Format(O, "##0.0##")
        .TextMatrix(2, 1) = Format(EL, "##0.0##")
        .TextMatrix(3, 1) = Format(d, "##0.0##")
        .TextMatrix(4, 1) = Format(em, "##0.0##")
    
        .TextMatrix(1, 2) = Format(po, "##0.0##")
        .TextMatrix(2, 2) = Format(pel, "##0.0##")
        .TextMatrix(3, 2) = Format(pd, "##0.0##")
        .TextMatrix(4, 2) = Format(pem, "##0.0##")
    
        .TextMatrix(1, 3) = Format(qo, "##0.0##")
        .TextMatrix(2, 3) = Format(qel, "##0.0##")
        .TextMatrix(3, 3) = Format(qd, "##0.0##")
        .TextMatrix(4, 3) = Format(qem, "##0.0##")
    
        .TextMatrix(1, 4) = Format(io, "##0.0##")
        .TextMatrix(2, 4) = Format(iel, "##0.0##")
        .TextMatrix(3, 4) = Format(id, "##0.0##")
        .TextMatrix(4, 4) = Format(iem, "##0.0##")
    
    End With
    
    
Else


    po = Val(txtPO.text)
    qo = Val(txtQo.text)
    hl = Val(txtHl.text)
    zl1 = Val(txtZL.text)
    hm = Val(txthM.text)
    zm1 = Val(txtZM.text)
    If po = 0 Then
        MsgBox "Ingrese el valor de la presión de operación", 64, "Cálculo de presiones, presiones e intensidades"
        txtPO.SetFocus
        Exit Sub
    End If
    If qo = 0 Then
        MsgBox "Ingrese el valor del caudal de operación", 64, "Cálculo de presiones, presiones e intensidades"
        txtQo.SetFocus
        Exit Sub
    End If
     If hl = 0 Then
        MsgBox "Ingrese el valor de las pérdidas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
        txtHl.SetFocus
        Exit Sub
    End If
   ' If zl1 = 0 Then
    '    MsgBox "Ingrese la diferencia de cotas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
     '   txtZL.SetFocus
      '  Exit Sub
   ' End If
    If hm = 0 Then
        MsgBox "Ingrese el valor de las pérdidas en el multiple", 64, "Cálculo de presiones, presiones e intensidades"
        txthM.SetFocus
        Exit Sub
    End If
    If zm1 = 0 Then
        MsgBox "Ingrese la diferencia de cotas en la múltiple", 64, "Cálculo de presiones, presiones e intensidades"
        txtZM.SetFocus
        Exit Sub
    End If
    txtH.Enabled = True
    txtA.Enabled = True
    xx = Val(CTG.ListIndex)
    If xx = -1 Then
    MsgBox "Seleccione el tipo de gotero", 64, "Cálculo de presiones, presiones e intensidades"
        CTG.SetFocus
        Exit Sub
    End If
    h = 0
    z1 = Combo1.ListIndex
    If z1 = -1 Then
        MsgBox "Defina la condición topográfica del lateral", 64, "Cálculo de presiones, presiones e intensidades"
        Combo1.SetFocus
        Exit Sub
    End If
    Select Case z1
    Case 0
        zl = -(zl1)
    Case 1
        zm = (zm1)
    Case 2
        zl = 0
    End Select
    
    z2 = Combo1.ListIndex
    If z2 = -1 Then
        MsgBox "Defina la condición topográfica del múltiple", 64, "Cálculo de presiones, presiones e intensidades"
        Combo2.SetFocus
        Exit Sub
    End If
    Select Case z1
    Case 0
        zm = -(zm1)
    Case 1
        zm = (zm1)
    Case 2
        zm = 0
    End Select
    
    If OCLN.Value = True Then
        EL = po + 0.77 * hl - 0.23 * zl
    Else
        EL = po + 0.63 * hl - 0.39 * zl
    End If
    
    If OCMN.Value = True Then
        em = EL + 0.77 * hm - 0.23 * zm
    Else
        em = EL + 0.63 * hm - 0.39 * zm
    End If
    
    Select Case xx
    Case 0
        X = 1
    Case 1
        X = 0.875
    Case 2
        X = 0.7
    Case 3
        X = 0.5
    Case 4
        X = 0.4
    Case 5
        X = 0.1
    Case 6
        X = 0
    End Select
    
    k = qo / (po) ^ X
    
    pel = EL - h
    qel = k * pel ^ X
    
    pd = EL - hl - h + zl
    qd = k * pd ^ X
    
    d = pd + h
    
    
    pem = em - h
    qem = k * pem ^ X
    
    
    With grdD
        .Cols = 3
        .Rows = 5
        .TextMatrix(0, 1) = "Pr. gotero(m)"
        .TextMatrix(0, 2) = "Caudal(m3/h)"
       
        
        .TextMatrix(1, 0) = "gotero operación"
        .TextMatrix(2, 0) = "Entrada lateral"
        .TextMatrix(3, 0) = "gotero distal"
        .TextMatrix(4, 0) = "Entrada multiple"
    
       
        .TextMatrix(1, 1) = Format(po, "##0.0##")
        .TextMatrix(2, 1) = Format(pel, "##0.0##")
        .TextMatrix(3, 1) = Format(pd, "##0.0##")
        .TextMatrix(4, 1) = Format(pem, "##0.0##")
    
        .TextMatrix(1, 2) = Format(qo, "##0.0####")
        .TextMatrix(2, 2) = Format(qel, "##0.0####")
        .TextMatrix(3, 2) = Format(qd, "##0.0####")
        .TextMatrix(4, 2) = Format(qem, "##0.0####")
    End With
End If
    Else
    If OA.Value = True Then
        po = Val(txtPO.text)
        qo = Val(txtQo.text)
        h = Val(txtH.text)
        a = Val(txtA.text)
        hl = Val(txtHl.text)
        zl1 = Val(txtZL.text)
        
        If po = 0 Then
            MsgBox "Ingrese el valor de la presión de operación", 64, "Cálculo de presiones, presiones e intensidades"
            txtPO.SetFocus
            Exit Sub
        End If
        If qo = 0 Then
            MsgBox "Ingrese el valor del caudal de operación", 64, "Cálculo de presiones, presiones e intensidades"
            txtQo.SetFocus
            Exit Sub
        End If
        If h = 0 Then
            MsgBox "Ingrese la altura del elevador", 64, "Cálculo de presiones, presiones e intensidades"
            txtH.SetFocus
            Exit Sub
        End If
        If a = 0 Then
            MsgBox "Ingrese el valor del área efectiva del aspersor", 64, "Cálculo de presiones, presiones e intensidades"
            txtA.SetFocus
            Exit Sub
        End If
        If hl = 0 Then
            MsgBox "Ingrese el valor de las pérdidas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
            txtHl.SetFocus
            Exit Sub
        End If
        ' If zl1 = 0 Then
         '   MsgBox "Ingrese la diferencia de cotas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
          '  txtZL.SetFocus
           ' Exit Sub
       ' End If
        z1 = Combo1.ListIndex
        If z1 = -1 Then
            MsgBox "Defina la condición topográfica del lateral", 64, "Cálculo de presiones, presiones e intensidades"
            Combo1.SetFocus
            Exit Sub
        End If
        Select Case z1
        Case 0
            zl = -(zl1)
        Case 1
            zl = (zl1)
        Case 2
            zl = 0
        End Select
        
        
        If OCLN.Value = True Then
            EL = po + 3 / 4 * hl + h - 0.38 * zl
        Else
            EL = po + 5 / 8 * hl + h - 0.38 * zl
        End If
        
        
        k = qo / Sqr(po)
        io = qo * 1000 / a
        O = po + h
        
        pel = EL - h
        qel = k * pel ^ 0.5
        iel = qel * 1000 / a
        
        pd = EL - hl - h + zl
        qd = k * pd ^ 0.5
        id = qd * 1000 / a
        d = pd + h
        
        
        
        
        With grdD
        
            .Visible = True
            .Cols = 5
            .Rows = 4
            .ColWidth(3) = 1600
            .ColWidth(4) = 1600
            .TextMatrix(0, 1) = "Pr. Piso (m)"
            .TextMatrix(0, 2) = "Pr. Aspersor (m)"
            .TextMatrix(0, 3) = "Caudal(m3/h)"
            .TextMatrix(0, 4) = "Intensidad (mm/h)"
            
            .TextMatrix(1, 0) = "Aspersor operación"
            .TextMatrix(2, 0) = "Entrada lateral"
            .TextMatrix(3, 0) = "Aspersor distal"
        
        
            .TextMatrix(1, 1) = Format(O, "##0.0##")
            .TextMatrix(2, 1) = Format(EL, "##0.0##")
            .TextMatrix(3, 1) = Format(d, "##0.0##")
        
        
            .TextMatrix(1, 2) = Format(po, "##0.0##")
            .TextMatrix(2, 2) = Format(pel, "##0.0##")
            .TextMatrix(3, 2) = Format(pd, "##0.0##")
            
        
            .TextMatrix(1, 3) = Format(qo, "##0.0##")
            .TextMatrix(2, 3) = Format(qel, "##0.0##")
            .TextMatrix(3, 3) = Format(qd, "##0.0##")
            
        
            .TextMatrix(1, 4) = Format(io, "##0.0##")
            .TextMatrix(2, 4) = Format(iel, "##0.0##")
            .TextMatrix(3, 4) = Format(id, "##0.0##")
            
        
        End With
        
        
    Else
    
    
        po = Val(txtPO.text)
        qo = Val(txtQo.text)
        hl = Val(txtHl.text)
        zl1 = Val(txtZL.text)
        If po = 0 Then
            MsgBox "Ingrese el valor de la presión de operación", 64, "Cálculo de presiones, presiones e intensidades"
            txtPO.SetFocus
            Exit Sub
        End If
        If qo = 0 Then
            MsgBox "Ingrese el valor del caudal de operación", 64, "Cálculo de presiones, presiones e intensidades"
            txtQo.SetFocus
            Exit Sub
        End If
         If hl = 0 Then
            MsgBox "Ingrese el valor de las pérdidas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
            txtHl.SetFocus
            Exit Sub
        End If
        'If zl1 = 0 Then
         '   MsgBox "Ingrese la diferencia de cotas en el lateral", 64, "Cálculo de presiones, presiones e intensidades"
          '  txtZL.SetFocus
           ' Exit Sub
        'End If
        xx = Val(CTG.ListIndex)
        If xx = -1 Then
        MsgBox "Seleccione el tipo de gotero", 64, "Cálculo de presiones, presiones e intensidades"
            CTG.SetFocus
            Exit Sub
        End If
        txtH.Enabled = True
        txtA.Enabled = True
        
        h = 0
        z1 = Combo1.ListIndex
        If z1 = -1 Then
            MsgBox "Defina la condición topográfica del lateral", 64, "Cálculo de presiones, presiones e intensidades"
            Combo1.SetFocus
            Exit Sub
        End If
        Select Case z1
        Case 0
            zl = -(zl1)
        Case 1
            zm = (zm1)
        Case 2
            zl = 0
        End Select
        
        If OCLN.Value = True Then
            EL = po + 0.77 * hl - 0.23 * zl
        Else
            EL = po + 0.63 * hl - 0.39 * zl
        End If
        
        Select Case xx
        Case 0
            X = 1
        Case 1
            X = 0.875
        Case 2
            X = 0.7
        Case 3
            X = 0.5
        Case 4
            X = 0.4
        Case 5
            X = 0.1
        Case 6
            X = 0
        End Select
        
        k = qo / (po) ^ X
        
        pel = EL - h
        qel = k * pel ^ X
        
        pd = EL - hl - h + zl
        qd = k * pd ^ 0.5
        
        d = pd + h
        
        
        
        With grdD
            .Cols = 3
            .Rows = 4
            .TextMatrix(0, 1) = "Pr. Aspersor (m)"
            .TextMatrix(0, 2) = "Caudal(m3/h)"
           
            
            .TextMatrix(1, 0) = "gotero operación"
            .TextMatrix(2, 0) = "Entrada lateral"
            .TextMatrix(3, 0) = "gotero distal"
        
        
           
            .TextMatrix(1, 1) = Format(po, "##0.0##")
            .TextMatrix(2, 1) = Format(pel, "##0.0##")
            .TextMatrix(3, 1) = Format(pd, "##0.0##")
        
        
            .TextMatrix(1, 2) = Format(qo, "##0.0####")
            .TextMatrix(2, 2) = Format(qel, "##0.0####")
            .TextMatrix(3, 2) = Format(qd, "##0.0####")
        
        End With
    End If
End If
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
    End Sub



Private Sub dflkdl_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub fff_Click()
frmcombDia.Show
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
StatusBar1.Panels(1).text = "Digite los datos básicos para el cálculo y oprima el botón de Calcular"
With grdD
    .ColWidth(0) = 1600
    .ColWidth(1) = 1600
    .ColWidth(2) = 1600
    .ColWidth(3) = 1600
    .ColWidth(4) = 1600
    
    .TextMatrix(0, 1) = "Presión Piso  (m)"
    .TextMatrix(0, 2) = "Presión Aspersor (m)"
    .TextMatrix(0, 3) = "Caudal (m3/h)"
    .TextMatrix(0, 4) = "Intensidad  (mm/h)"
End With
With CTG
    .AddItem "Laminar (x=1)"
    .AddItem "Microtubos (x=0.875)"
    .AddItem "Helicoidales (x=0.7)"
    .AddItem "Turbulentos, orificio, laberinto (x=0.5)"
    .AddItem "Vortex(X=0.4)"
    .AddItem "Autocompensado (x=0.1)"
    .AddItem " Autocompensado perfecto (x=0)"
End With

End Sub






Private Sub mcp_Click()
frmHprincipal.Show
End Sub

Private Sub mgensu_Click()
FrmHLaterales.Show
End Sub

Private Sub msb_Click()
frmbomba.Show
End Sub

Private Sub OA_Click()
Label21.Visible = False
CTG.Visible = False
OA.ForeColor = &HC0&
OG.ForeColor = &H80000012
OG.Value = False
txtPO.SetFocus
txtH.BackColor = &H80000005
txtA.BackColor = &H80000005
txtH.Enabled = True
txtA.Enabled = True
End Sub

Private Sub OCLN_Click()
OCLN.ForeColor = &HC0&
OCLS.ForeColor = &H80000012
OCLS.Value = False
txtPO.SetFocus
End Sub

Private Sub OCLS_Click()
OCLS.ForeColor = &HC0&
OCLN.ForeColor = &H80000012
OCLN.Value = False
txtPO.SetFocus
End Sub

Private Sub OCMN_Click()
OCMN.ForeColor = &HC0&
OCMS.ForeColor = &H80000012
OCMS.Value = False
txtPO.SetFocus
End Sub

Private Sub OCMS_Click()
OCMS.ForeColor = &HC0&
OCMN.ForeColor = &H80000012
OCMN.Value = False
txtPO.SetFocus
End Sub

Private Sub OG_Click()
Label21.Visible = True
CTG.Visible = True
OG.ForeColor = &HC0&
OA.ForeColor = &H80000012
OA.Value = False
txtPO.SetFocus
txtH.BackColor = &H80000016
txtA.BackColor = &H80000016
txtH.Enabled = False
txtA.Enabled = False
End Sub

Private Sub Option1_Click()
Frame4.Enabled = True
txthM.Enabled = True
txtZM.Enabled = True
Combo2.Enabled = True
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
Option2.Value = False
txthM.BackColor = &H80000005
txtZM.BackColor = &H80000005
End Sub

Private Sub Option2_Click()
Frame4.Enabled = False
txthM.Enabled = False
txtZM.Enabled = False
Combo2.Enabled = False
Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
Option1.Value = False
txthM.BackColor = &H80000016
txtZM.BackColor = &H80000016

End Sub
