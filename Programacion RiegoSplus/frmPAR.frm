VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPAR 
   Caption         =   "Porcentaje de área regada (PAR) en goteo"
   ClientHeight    =   7050
   ClientLeft      =   1290
   ClientTop       =   990
   ClientWidth     =   8385
   Icon            =   "frmPAR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   8385
   Begin VB.CommandButton Command1 
      Caption         =   "Cuadro del PAR"
      Height          =   375
      Left            =   2040
      TabIndex        =   35
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   757
      TabIndex        =   28
      Top             =   3240
      Width           =   3975
      Begin VB.TextBox txteg 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtpar 
         BackColor       =   &H80000016&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "mts"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "%"
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Espacimiento de goteros"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "PAR"
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos básicos"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   465
      TabIndex        =   13
      Top             =   1080
      Width           =   7455
      Begin VB.ComboBox CCaudal 
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox CEspaciamiento 
         Height          =   315
         Left            =   5040
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox CTextura 
         Height          =   315
         Left            =   3000
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox CSituacion 
         Height          =   315
         Left            =   360
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtAA 
         Height          =   285
         Left            =   3000
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtbb 
         Height          =   285
         Left            =   5040
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "mts"
         Height          =   255
         Left            =   6240
         TabIndex        =   27
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "mts"
         Height          =   255
         Left            =   6360
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Caudal del gotero"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Textura"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Espaciamiento entre hileras "
         Height          =   495
         Left            =   5040
         TabIndex        =   23
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Situación"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label37 
         Caption         =   "# de goteros"
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Espaciamiento de plantas"
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   345
      TabIndex        =   6
      Top             =   5160
      Width           =   7695
      Begin VB.CommandButton bcalcular 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   120
         Picture         =   "frmPAR.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2040
         Picture         =   "frmPAR.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   3960
         Picture         =   "frmPAR.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Regresar"
         Height          =   735
         Left            =   5880
         Picture         =   "frmPAR.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
      Begin MSFlexGridLib.MSFlexGrid gridlista 
         Height          =   975
         Left            =   360
         TabIndex        =   11
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   12
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid gridmenos15 
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   12
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   3015
         Left            =   2040
         TabIndex        =   2
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   12
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid4 
         Height          =   3015
         Left            =   4080
         TabIndex        =   3
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   12
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid8 
         Height          =   3015
         Left            =   6000
         TabIndex        =   4
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   12
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid gridmas10 
         Height          =   3015
         Left            =   7560
         TabIndex        =   5
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   12
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   6675
      Width           =   8385
      _ExtentX        =   14790
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   5197
      Picture         =   "frmPAR.frx":29F2
      Top             =   3120
      Width           =   2430
   End
   Begin VB.Label Label17 
      Caption         =   "Estimación del porcentaje de área regada"
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
      Left            =   480
      TabIndex        =   12
      Top             =   360
      Width           =   6255
   End
   Begin VB.Menu m1 
      Caption         =   "Suelo- clima"
      Begin VB.Menu mgs 
         Caption         =   "Generales suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mch 
         Caption         =   "Conductividad hidráulica"
      End
      Begin VB.Menu mev 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu mm 
      Caption         =   "Menú principal"
   End
End
Attribute VB_Name = "frmPAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Single
Dim T As Single
Dim EL As Integer
Dim s As Integer

Private Sub Bcalcular_Click()
On Error GoTo mensaje:

s = Val(CSituacion.ListIndex)
q = Val(CCaudal.ListIndex)
T = Val(CTextura.ListIndex)
EL = Val(CEspaciamiento.ListIndex)
ell = Val(CEspaciamiento.text)
If q = -1 Then
    MsgBox "Ingrese el valor del caudal ", 64, "PAR"
    CCaudal.SetFocus
    Exit Sub
End If
If s = -1 Then
    MsgBox "Ingrese la condición de goteo", 64, "PAR"
    CSituacion.SetFocus
    Exit Sub
End If
If T = -1 Then
    MsgBox "Ingrese la textura", 64, "PAR"
    CTextura.SetFocus
    Exit Sub
End If
If EL = -1 Then
    MsgBox "Ingrese el espaciamiento entre hileras", 64, "PAR"
    CEspaciamiento.SetFocus
    Exit Sub
End If

Select Case s
Case 0
    Call PARRr
    Call ESPGOTEROS
    txteg = eggg
    txtpar = PARRa
Case 1
    Call ESPGOTEROS
    Call buscarS1
    s1 = Val(gridlista.TextMatrix(sss11, 0))
    s2 = ell - s1
    For j% = 0 To 11
        If s2 = Val(gridlista.TextMatrix(j%, 0)) Then
        EL = j%
        Call PARRr
        parr = (100 * s1 + PARRa * s2) / ell
        txteg = eggg
        txtpar = parr
        Exit Sub
        End If
    Next j%
    s3 = CInt(s2)
    For j% = 0 To 11
        If s3 = Val(gridlista.TextMatrix(j%, 0)) Then
        EL = j%
        Call PARRr
        parr = (100 * s1 + PARRa * s2) / ell
        txteg = eggg
        txtpar = parr
        Exit Sub
        End If
    Next j%
Case 2
    Call ESPGOTEROS
    nnn = Val(txtaa.text)
    st = Val(txtbb.text)
    If nnn = 0 Then
        MsgBox "Ingrese el número de goteros por árbol ", 64, "PAR"
        txtaa.SetFocus
        Exit Sub
    End If
    If st = 0 Then
        MsgBox "Ingrese la distancia entre árboles", 64, "PAR"
        txtbb.SetFocus
        Exit Sub
    End If
    Call buscarS1
    s1 = Val(gridlista.TextMatrix(sss11, 0))
    parr = (100 * nnn * eggg * s1) / (ell * st)
    txteg = eggg
    txtpar = parr
End Select

Exit Sub
mensaje:
MsgBox "Ingrese adecuadamente los datos que se le solocitan"

End Sub

Public Sub PARRr()
Select Case q
Case 0
    If T = 0 Then
        PAR = Val(gridmenos15.TextMatrix(EL, 0))
    Else
        If T = 1 Then
            PAR = Val(gridmenos15.TextMatrix(EL, 1))
        Else
            PAR = Val(gridmenos15.TextMatrix(EL, 2))
        End If
    End If
Case 1
    If T = 0 Then
        PAR = Val(grid2.TextMatrix(EL, 0))
    Else
        If T = 1 Then
            PAR = Val(grid2.TextMatrix(EL, 1))
        Else
            PAR = Val(grid2.TextMatrix(EL, 2))
        End If
    End If
Case 2
If T = 0 Then
        PAR = Val(grid4.TextMatrix(EL, 0))
    Else
        If T = 1 Then
            PAR = Val(grid4.TextMatrix(EL, 1))
        Else
            PAR = Val(grid4.TextMatrix(EL, 2))
        End If
    End If
Case 3
If T = 0 Then
        PAR = Val(grid8.TextMatrix(EL, 0))
    Else
        If T = 1 Then
            PAR = Val(grid8.TextMatrix(EL, 1))
        Else
            PAR = Val(grid8.TextMatrix(EL, 2))
        End If
    End If
Case 4
If T = 0 Then
        PAR = Val(gridmas10.TextMatrix(EL, 0))
    Else
        If T = 1 Then
            PAR = Val(gridmas10.TextMatrix(EL, 1))
        Else
            PAR = Val(gridmas10.TextMatrix(EL, 2))
        End If
    End If
End Select
PARRa = PAR
End Sub
Public Sub ESPGOTEROS()
Select Case q
Case 0
    If T = 0 Then
        eg = Val(0.2)
    Else
        If T = 1 Then
            eg = Val(0.5)
        Else
            eg = Val(0.9)
        End If
    End If
Case 1
    If T = 0 Then
        eg = Val(0.3)
    Else
        If T = 1 Then
            eg = Val(0.7)
        Else
            eg = Val(1)
        End If
    End If
Case 2
If T = 0 Then
        eg = Val(0.6)
    Else
        If T = 1 Then
            eg = Val(1)
        Else
            eg = Val(1.3)
        End If
    End If
Case 3
If T = 0 Then
        eg = Val(1)
    Else
        If T = 1 Then
            eg = Val(1.3)
        Else
            eg = Val(1.7)
        End If
    End If
Case 4
If T = 0 Then
        eg = Val(1.3)
    Else
        If T = 1 Then
            eg = Val(1.6)
        Else
            eg = Val(2)
        End If
    End If
End Select
eggg = eg
End Sub

Public Sub buscarS1()
Select Case q
Case 0
    Select Case T
    Case 0
        For j% = 0 To 11
        If gridmenos15.TextMatrix(j%, 0) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 1
        For j% = 0 To 11
        If gridmenos15.TextMatrix(j%, 1) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 2
        For j% = 0 To 11
        If gridmenos15.TextMatrix(j%, 2) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    End Select
Case 1
    Select Case T
    Case 0
        For j% = 0 To 11
        If grid2.TextMatrix(j%, 0) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 1
        For j% = 0 To 11
        If grid2.TextMatrix(j%, 1) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 2
        For j% = 0 To 11
        If grid2.TextMatrix(j%, 2) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    End Select
Case 2
    Select Case T
    Case 0
        For j% = 0 To 11
        If grid4.TextMatrix(j%, 0) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 1
        For j% = 0 To 11
        If grid4.TextMatrix(j%, 1) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 2
        For j% = 0 To 11
        If grid4.TextMatrix(j%, 2) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    End Select
Case 3
    Select Case T
    Case 0
        For j% = 0 To 11
        If grid8.TextMatrix(j%, 0) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 1
        For j% = 0 To 11
        If grid8.TextMatrix(j%, 1) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 2
        For j% = 0 To 11
        If grid8.TextMatrix(j%, 2) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    End Select
Case 4
 Select Case T
    Case 0
        For j% = 0 To 11
        If gridmas10.TextMatrix(j%, 0) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 1
        For j% = 0 To 11
        If gridmas10.TextMatrix(j%, 1) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    Case 2
        For j% = 0 To 11
        If gridmas10.TextMatrix(j%, 2) <> 100 Then
        sss11 = j% - 1
        Exit Sub
        End If
        Next j%
    End Select
End Select
End Sub

Private Sub bfinailizar_Click()
frmDAgoteo.txteg = frmPAR.txteg.text
frmDAgoteo.txtpar = frmPAR.txtpar.text
frmDAgoteo.Show
Unload Me
End Sub

Private Sub blimpiar_Click()
txtpar.text = ""
txteg.text = ""
txtaa.text = ""
txtbb.text = ""
CCaudal.text = ""
CSituacion.text = ""
CTextura.text = ""
CEspaciamiento.text = ""
Label37.Visible = False
Label9.Visible = False
Label10.Visible = False
txtaa.Visible = False
txtbb.Visible = False
End Sub

Private Sub Command1_Click()
Dialog.Show
End Sub

Private Sub CSituacion_click()
THC = Val(CSituacion.ListIndex)
Select Case THC
Case 0
Label37.Visible = False
Label9.Visible = False
Label10.Visible = False
txtaa.Visible = False
txtbb.Visible = False
Case 1
Label37.Visible = False
Label9.Visible = False
Label10.Visible = False
txtaa.Visible = False
txtbb.Visible = False
Case 2
Label37.Visible = True
Label9.Visible = True
Label10.Visible = True
txtaa.Visible = True
txtbb.Visible = True
End Select



End Sub

Private Sub Form_Load()
With CCaudal
    .AddItem "Menos de 1.5 lts/s"
    .AddItem "2 lts/s"
    .AddItem "4 lts/s"
    .AddItem "8 lts/s"
    .AddItem "Más de 10 lts/s"
End With
With CTextura
    .AddItem "Grueso"
    .AddItem "Medio"
    .AddItem "Fino"
End With
With CEspaciamiento
    .AddItem "0.8"
    .AddItem "1.0"
    .AddItem "1.2"
    .AddItem "1.5"
    .AddItem "2.0"
    .AddItem "2.5"
    .AddItem "3.0"
    .AddItem "3.5"
    .AddItem "4.0"
    .AddItem "4.5"
    .AddItem "5.0"
    .AddItem "6.0"
End With
With CSituacion
    .AddItem "Un línea de goteros"
    .AddItem "Doble línea"
    .AddItem "Goteros al rededor de un árbol"
End With
StatusBar1.Panels(1).text = "Seleccione los datos básicos y oprima el botón Calcular"
With gridmenos15
    .TextMatrix(0, 0) = Val(38)
    .TextMatrix(1, 0) = Val(33)
    .TextMatrix(2, 0) = Val(25)
    .TextMatrix(3, 0) = Val(20)
    .TextMatrix(4, 0) = Val(15)
    .TextMatrix(5, 0) = Val(12)
    .TextMatrix(6, 0) = Val(10)
    .TextMatrix(7, 0) = Val(9)
    .TextMatrix(8, 0) = Val(8)
    .TextMatrix(9, 0) = Val(7)
    .TextMatrix(10, 0) = Val(6)
    .TextMatrix(11, 0) = Val(5)
    .TextMatrix(0, 1) = Val(88)
    .TextMatrix(1, 1) = Val(70)
    .TextMatrix(2, 1) = Val(58)
    .TextMatrix(3, 1) = Val(47)
    .TextMatrix(4, 1) = Val(35)
    .TextMatrix(5, 1) = Val(28)
    .TextMatrix(6, 1) = Val(23)
    .TextMatrix(7, 1) = Val(20)
    .TextMatrix(8, 1) = Val(18)
    .TextMatrix(9, 1) = Val(16)
    .TextMatrix(10, 1) = Val(14)
    .TextMatrix(11, 1) = Val(12)
    .TextMatrix(0, 2) = Val(100)
    .TextMatrix(1, 2) = Val(100)
    .TextMatrix(2, 2) = Val(92)
    .TextMatrix(3, 2) = Val(73)
    .TextMatrix(4, 2) = Val(55)
    .TextMatrix(5, 2) = Val(44)
    .TextMatrix(6, 2) = Val(37)
    .TextMatrix(7, 2) = Val(31)
    .TextMatrix(8, 2) = Val(28)
    .TextMatrix(9, 2) = Val(24)
    .TextMatrix(10, 2) = Val(22)
    .TextMatrix(11, 2) = Val(20)
End With
With grid2
    .TextMatrix(0, 0) = Val(50)
    .TextMatrix(1, 0) = Val(40)
    .TextMatrix(2, 0) = Val(33)
    .TextMatrix(3, 0) = Val(26)
    .TextMatrix(4, 0) = Val(20)
    .TextMatrix(5, 0) = Val(16)
    .TextMatrix(6, 0) = Val(13)
    .TextMatrix(7, 0) = Val(11)
    .TextMatrix(8, 0) = Val(10)
    .TextMatrix(9, 0) = Val(9)
    .TextMatrix(10, 0) = Val(8)
    .TextMatrix(11, 0) = Val(7)
    .TextMatrix(0, 1) = Val(100)
    .TextMatrix(1, 1) = Val(80)
    .TextMatrix(2, 1) = Val(67)
    .TextMatrix(3, 1) = Val(53)
    .TextMatrix(4, 1) = Val(40)
    .TextMatrix(5, 1) = Val(32)
    .TextMatrix(6, 1) = Val(26)
    .TextMatrix(7, 1) = Val(23)
    .TextMatrix(8, 1) = Val(20)
    .TextMatrix(9, 1) = Val(18)
    .TextMatrix(10, 1) = Val(16)
    .TextMatrix(11, 1) = Val(14)
    .TextMatrix(0, 2) = Val(100)
    .TextMatrix(1, 2) = Val(100)
    .TextMatrix(2, 2) = Val(100)
    .TextMatrix(3, 2) = Val(80)
    .TextMatrix(4, 2) = Val(60)
    .TextMatrix(5, 2) = Val(48)
    .TextMatrix(6, 2) = Val(40)
    .TextMatrix(7, 2) = Val(34)
    .TextMatrix(8, 2) = Val(30)
    .TextMatrix(9, 2) = Val(26)
    .TextMatrix(10, 2) = Val(24)
    .TextMatrix(11, 2) = Val(20)
End With
With grid4
    .TextMatrix(0, 0) = Val(100)
    .TextMatrix(1, 0) = Val(80)
    .TextMatrix(2, 0) = Val(67)
    .TextMatrix(3, 0) = Val(53)
    .TextMatrix(4, 0) = Val(40)
    .TextMatrix(5, 0) = Val(32)
    .TextMatrix(6, 0) = Val(26)
    .TextMatrix(7, 0) = Val(23)
    .TextMatrix(8, 0) = Val(20)
    .TextMatrix(9, 0) = Val(18)
    .TextMatrix(10, 0) = Val(16)
    .TextMatrix(11, 0) = Val(14)
    .TextMatrix(0, 1) = Val(100)
    .TextMatrix(1, 1) = Val(100)
    .TextMatrix(2, 1) = Val(100)
    .TextMatrix(3, 1) = Val(80)
    .TextMatrix(4, 1) = Val(60)
    .TextMatrix(5, 1) = Val(48)
    .TextMatrix(6, 1) = Val(40)
    .TextMatrix(7, 1) = Val(34)
    .TextMatrix(8, 1) = Val(30)
    .TextMatrix(9, 1) = Val(26)
    .TextMatrix(10, 1) = Val(24)
    .TextMatrix(11, 1) = Val(20)
    .TextMatrix(0, 2) = Val(100)
    .TextMatrix(1, 2) = Val(100)
    .TextMatrix(2, 2) = Val(100)
    .TextMatrix(3, 2) = Val(100)
    .TextMatrix(4, 2) = Val(80)
    .TextMatrix(5, 2) = Val(64)
    .TextMatrix(6, 2) = Val(53)
    .TextMatrix(7, 2) = Val(46)
    .TextMatrix(8, 2) = Val(40)
    .TextMatrix(9, 2) = Val(36)
    .TextMatrix(10, 2) = Val(32)
    .TextMatrix(11, 2) = Val(27)
End With
With grid8
    .TextMatrix(0, 0) = Val(100)
    .TextMatrix(1, 0) = Val(100)
    .TextMatrix(2, 0) = Val(100)
    .TextMatrix(3, 0) = Val(80)
    .TextMatrix(4, 0) = Val(60)
    .TextMatrix(5, 0) = Val(48)
    .TextMatrix(6, 0) = Val(40)
    .TextMatrix(7, 0) = Val(34)
    .TextMatrix(8, 0) = Val(30)
    .TextMatrix(9, 0) = Val(26)
    .TextMatrix(10, 0) = Val(24)
    .TextMatrix(11, 0) = Val(20)
    .TextMatrix(0, 1) = Val(100)
    .TextMatrix(1, 1) = Val(100)
    .TextMatrix(2, 1) = Val(100)
    .TextMatrix(3, 1) = Val(100)
    .TextMatrix(4, 1) = Val(80)
    .TextMatrix(5, 1) = Val(64)
    .TextMatrix(6, 1) = Val(53)
    .TextMatrix(7, 1) = Val(46)
    .TextMatrix(8, 1) = Val(40)
    .TextMatrix(9, 1) = Val(36)
    .TextMatrix(10, 1) = Val(32)
    .TextMatrix(11, 1) = Val(27)
    .TextMatrix(0, 2) = Val(100)
    .TextMatrix(1, 2) = Val(100)
    .TextMatrix(2, 2) = Val(100)
    .TextMatrix(3, 2) = Val(100)
    .TextMatrix(4, 2) = Val(100)
    .TextMatrix(5, 2) = Val(80)
    .TextMatrix(6, 2) = Val(67)
    .TextMatrix(7, 2) = Val(57)
    .TextMatrix(8, 2) = Val(50)
    .TextMatrix(9, 2) = Val(44)
    .TextMatrix(10, 2) = Val(40)
    .TextMatrix(11, 2) = Val(34)
End With
With gridmas10
    .TextMatrix(0, 0) = Val(100)
    .TextMatrix(1, 0) = Val(100)
    .TextMatrix(2, 0) = Val(100)
    .TextMatrix(3, 0) = Val(100)
    .TextMatrix(4, 0) = Val(80)
    .TextMatrix(5, 0) = Val(64)
    .TextMatrix(6, 0) = Val(53)
    .TextMatrix(7, 0) = Val(46)
    .TextMatrix(8, 0) = Val(40)
    .TextMatrix(9, 0) = Val(36)
    .TextMatrix(10, 0) = Val(32)
    .TextMatrix(11, 0) = Val(27)
    .TextMatrix(0, 1) = Val(100)
    .TextMatrix(1, 1) = Val(100)
    .TextMatrix(2, 1) = Val(100)
    .TextMatrix(3, 1) = Val(100)
    .TextMatrix(4, 1) = Val(100)
    .TextMatrix(5, 1) = Val(80)
    .TextMatrix(6, 1) = Val(67)
    .TextMatrix(7, 1) = Val(57)
    .TextMatrix(8, 1) = Val(50)
    .TextMatrix(9, 1) = Val(44)
    .TextMatrix(10, 1) = Val(40)
    .TextMatrix(11, 1) = Val(34)
    .TextMatrix(0, 2) = Val(100)
    .TextMatrix(1, 2) = Val(100)
    .TextMatrix(2, 2) = Val(100)
    .TextMatrix(3, 2) = Val(100)
    .TextMatrix(4, 2) = Val(100)
    .TextMatrix(5, 2) = Val(100)
    .TextMatrix(6, 2) = Val(80)
    .TextMatrix(7, 2) = Val(68)
    .TextMatrix(8, 2) = Val(60)
    .TextMatrix(9, 2) = Val(53)
    .TextMatrix(10, 2) = Val(48)
    .TextMatrix(11, 2) = Val(40)
End With
With gridlista
    .TextMatrix(0, 0) = 0.8
    .TextMatrix(1, 0) = 1
    .TextMatrix(2, 0) = 1.2
    .TextMatrix(3, 0) = 1.5
    .TextMatrix(4, 0) = 2
    .TextMatrix(5, 0) = 2.5
    .TextMatrix(6, 0) = 3
    .TextMatrix(7, 0) = 3.5
    .TextMatrix(8, 0) = 4
    .TextMatrix(9, 0) = 4.5
    .TextMatrix(10, 0) = 5
    .TextMatrix(11, 0) = 6
End With
Frame3.Caption = ""
End Sub



Private Sub mch_Click()
frmconductividad.Show
End Sub

Private Sub mev_Click()
frmETO.Show
End Sub

Private Sub mgs_Click()
frmgenerales.Show
End Sub

Private Sub mm_Click()
Unload Me

End Sub

Private Sub mt_Click()
frmtextura.Show
End Sub
