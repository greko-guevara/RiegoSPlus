VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmconductividad 
   Caption         =   "Conductividad Hidráulica"
   ClientHeight    =   6495
   ClientLeft      =   1290
   ClientTop       =   1155
   ClientWidth     =   8535
   Icon            =   "frmconductividad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   8535
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   600
      TabIndex        =   30
      Top             =   4680
      Width           =   7335
      Begin VB.CommandButton bevaluar 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   120
         Picture         =   "frmconductividad.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   5520
         Picture         =   "frmconductividad.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   3720
         Picture         =   "frmconductividad.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmconductividad.frx":2308
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Agujero de barrena"
      TabPicture(0)   =   "frmconductividad.frx":29F2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Image1"
      Tab(0).Control(3)=   "Label22"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Barreno Invertido"
      TabPicture(1)   =   "frmconductividad.frx":2A0E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
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
         Height          =   1215
         Left            =   600
         TabIndex        =   45
         Top             =   2520
         Width           =   2535
         Begin VB.TextBox txtkk 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   46
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "m/día"
            Height          =   255
            Left            =   1800
            TabIndex        =   49
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Conductividad Hidráulica"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label15 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   2760
            TabIndex        =   47
            Top             =   -240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de Entrada"
         Height          =   1575
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   7815
         Begin VB.TextBox txtt1 
            Height          =   285
            Left            =   6000
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txty1 
            Height          =   285
            Left            =   2040
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txty2 
            Height          =   285
            Left            =   2040
            TabIndex        =   34
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtr2 
            Height          =   285
            Left            =   2040
            TabIndex        =   33
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Tiempo de la prueba"
            Height          =   255
            Left            =   4080
            TabIndex        =   44
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "seg"
            Height          =   255
            Left            =   7320
            TabIndex        =   43
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "Radio del Agujero (r)"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Depresión Inicial (Y1)"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Depresión Final (Y2)"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "cm"
            Height          =   255
            Left            =   3360
            TabIndex        =   39
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "cm"
            Height          =   255
            Left            =   3360
            TabIndex        =   38
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "cm"
            Height          =   255
            Left            =   3360
            TabIndex        =   37
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Entrada"
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   7815
         Begin VB.TextBox txtr1 
            Height          =   285
            Left            =   2040
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtyn 
            Height          =   285
            Left            =   2040
            TabIndex        =   2
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtyo 
            Height          =   285
            Left            =   2040
            TabIndex        =   1
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txth 
            Height          =   285
            Left            =   6000
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtdt 
            Height          =   285
            Left            =   6000
            TabIndex        =   4
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   5040
            TabIndex        =   5
            Text            =   "H vrs capa impermeable"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "cm"
            Height          =   255
            Left            =   3360
            TabIndex        =   27
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "cm"
            Height          =   255
            Left            =   3360
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "cm"
            Height          =   255
            Left            =   3360
            TabIndex        =   25
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Depresión Final (Yn)"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label36 
            Caption         =   "Depresión Inicial (Yo)"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label37 
            Caption         =   "Radio del Agujero (r)"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label39 
            Caption         =   "cm"
            Height          =   255
            Left            =   7320
            TabIndex        =   21
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label40 
            Caption         =   "seg"
            Height          =   255
            Left            =   7320
            TabIndex        =   20
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label42 
            Caption         =   "Tiempo de la prueba"
            Height          =   255
            Left            =   4080
            TabIndex        =   19
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label43 
            Caption         =   "Profundidad del Pozo (H)"
            Height          =   255
            Left            =   4080
            TabIndex        =   18
            Top             =   360
            Width           =   1815
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
         Height          =   1215
         Left            =   -74640
         TabIndex        =   12
         Top             =   2280
         Width           =   2535
         Begin VB.TextBox txtk1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   2760
            TabIndex        =   16
            Top             =   -240
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "Conductividad Hidráulica"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "m/día"
            Height          =   255
            Left            =   1800
            TabIndex        =   14
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1680
         Left            =   -70680
         Picture         =   "frmconductividad.frx":2A2A
         Top             =   2160
         Width           =   2760
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   4320
         Picture         =   "frmconductividad.frx":10E3C
         Top             =   2160
         Width           =   2760
      End
      Begin VB.Label Label22 
         Caption         =   "Método del Agujero Barreno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   -71040
         TabIndex        =   29
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label21 
         Caption         =   "Método de Barreno Invertido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3960
         TabIndex        =   28
         Top             =   0
         Width           =   3735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   6120
      Width           =   8535
      _ExtentX        =   15055
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
   Begin VB.Label lbltitulo 
      Caption         =   "Determinación de la conductividad Hidráulica K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   6615
   End
   Begin VB.Menu motros 
      Caption         =   "Otros Parámetros"
      Begin VB.Menu Mgern 
         Caption         =   "Generales Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mevapo 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu mm 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmconductividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub bevaluar_Click()
On Error GoTo mensaje

Select Case SSTab1.Tab
Case 0
    Rem barreno********************************************
    R = Val(txtr1.text)
    yo = Val(txtyo.text)
    yn = Val(txtyn.text)
    dt = Val(txtdt.text)
    h1 = Val(txtH.text)
    If R = 0 Then
    MsgBox "Introduzca el valor del radio del agujero", 64, "Determinación de Conductividad Hidráulica"
    txtr1.SetFocus
    Exit Sub
    End If
    If yo = 0 Then
    MsgBox "Introduzca el valor de la depresión inicial", 64, "Determinación de Conductividad Hidráulica"
    txtyo.SetFocus
    Exit Sub
    End If
    If yn = 0 Then
    MsgBox "Introduzca el valor de la depresión final", 64, "Determinación de Conductividad Hidráulica"
    txtyn.SetFocus
    Exit Sub
    End If
    If dt = 0 Then
    MsgBox "Introduzca el valor del tiempo de la prueba", 64, "Determinación de Conductividad Hidráulica"
    txtdt.SetFocus
    Exit Sub
    End If
    If h1 = 0 Then
    MsgBox "Introduzca el valor de la profundidad del pozo", 64, "Determinación de Conductividad Hidráulica"
    txtH.SetFocus
    Exit Sub
    End If
    dy = yo - yn
    If dy <= 0 Then
    MsgBox "Yo debe ser mayor a Yn", 64, "Determinación de Conductividad Hidráulica"
    txtyo.SetFocus
    Exit Sub
    End If
    dyt = dy / dt
    Y = yo - dy / 2
    
    
    If Combo1.text = "H vrs capa impermeable" Then
    MsgBox "Introduzca la relación de H vrs capa impermeable ", 64, "Determinación de Conductividad Hidráulica"
    Combo1.SetFocus
    Exit Sub
    End If
    
    CCC = Combo1.ListIndex
    Select Case CCC
    Case 0
        k = 3600 * R ^ 2 * dyt / ((h1 + 10 * R) * (2 - Y / h1) * Y)
    Case 1
        k = 4000 * R ^ 2 * dyt / ((h1 + 20 * R) * (2 - Y / h1) * Y)
    End Select
    txtk1.text = Format(k, "##0.0##")
Case 1
'barreno invertido***********************************************
    R1 = Val(txtr2.text)
    Y2 = Val(txty2.text)
    Y1 = Val(txty1.text)
    t1 = Val(txtt1.text)
    If R1 = 0 Then
    MsgBox "Introduzca el valor del radio del agujero", 64, "Determinación de Conductividad Hidráulica"
    txtr2.SetFocus
    Exit Sub
    End If
    If Y1 = 0 Then
    MsgBox "Introduzca el valor de la depresión inicial", 64, "Determinación de Conductividad Hidráulica"
    txty1.SetFocus
    Exit Sub
    End If
    If Y2 = 0 Then
    MsgBox "Introduzca el valor de la depresión final", 64, "Determinación de Conductividad Hidráulica"
    txty2.SetFocus
    Exit Sub
    End If
    If t1 = 0 Then
    MsgBox "Introduzca el valor del tiempo de prueba", 64, "Determinación de Conductividad Hidráulica"
    txtt1.SetFocus
    Exit Sub
    End If
    dy = Y1 - Y2
    If dy <= 0 Then
    MsgBox "Y1 debe ser mayor a Y2", 64, "Determinación de Conductividad Hidráulica"
    txty1.SetFocus
    Exit Sub
    End If
    
    k = (R1 / (2 * t1)) * Log((Y1 + R1 / 2) / (Y2 + R1 / 2)) * 864
    txtkk.text = Format(k, "##0.0##")
End Select

Exit Sub
mensaje:
MsgBox "Introduzca adecuadamente los datos", 64, "Determinación de Conductividad Hidráulica"
    
End Sub

Private Sub bfinailizar_Click()
Unload Me
End Sub


Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtyo.text = ""
txtyn.text = ""
txtr1.text = ""
txtH.text = ""
txtdt.text = ""
txtr2.text = ""
txty1.text = ""
txty2.text = ""
txtt1.text = ""
txtk1.text = ""
txtkk.text = ""

End Sub

Private Sub Form_Load()
With Combo1
    .AddItem "S=0"
    .AddItem "s > 0.5 H"
End With
StatusBar1.Panels(1).text = "Para ambos métodos: Digite los Datos de Entrada y oprima Calcular para estimar K"

End Sub



Private Sub mevapo_Click()
frmETO.Show
End Sub

Private Sub Mgern_Click()
frmgenerales.Show
End Sub

Private Sub mm_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub mreffe_Click()
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
