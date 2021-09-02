VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frminfiltracioncanales 
   Caption         =   "Cálculo de Infiltraciones en Canales"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frminfiltracioncanales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   2760
      TabIndex        =   38
      Top             =   6480
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2280
         Picture         =   "frminfiltracioncanales.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bL 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   240
         Picture         =   "frminfiltracioncanales.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bS 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   4320
         Picture         =   "frminfiltracioncanales.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pérdidas en m3/s-km"
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   960
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton bP 
         Caption         =   "&Evaluar"
         Height          =   615
         Left            =   6960
         Picture         =   "frminfiltracioncanales.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox tP 
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Height          =   1695
         Left            =   5400
         TabIndex        =   27
         Top             =   360
         Width           =   3975
         Begin VB.TextBox txtp 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   31
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox tQ2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox tQ1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   28
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "m3/s"
            Height          =   255
            Left            =   3360
            TabIndex        =   47
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "%"
            Height          =   255
            Left            =   3360
            TabIndex        =   46
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "%"
            Height          =   255
            Left            =   3360
            TabIndex        =   45
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "m3/s"
            Height          =   255
            Left            =   3360
            TabIndex        =   44
            Top             =   240
            Width           =   495
         End
         Begin VB.Label tper 
            Caption         =   "% de Pérdidas"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Caudal correg. "
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label11 
            Caption         =   "Caudal de salida"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label13 
            Caption         =   "Eficiencia "
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Sel. los valores a promediar"
         Height          =   2055
         Left            =   3000
         TabIndex        =   26
         Top             =   360
         Width           =   2175
         Begin VB.CheckBox c1 
            Caption         =   "Ingham"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox c7 
            Caption         =   "Moritz"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CheckBox c5 
            Caption         =   "Punjab"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox c4 
            Caption         =   "Davis- Wilson"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox c3 
            Caption         =   "Pavlovski"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox c2 
            Caption         =   "Etcheverry"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox c6 
            Caption         =   "Kostiakov"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid gDatos 
         Height          =   2055
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   8
         BackColorFixed  =   -2147483626
      End
      Begin VB.Label Label7 
         Caption         =   "Promedio"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Entrada"
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   960
      TabIndex        =   21
      Top             =   840
      Width           =   9615
      Begin VB.TextBox tQ 
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox tB 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox tZ 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   2655
         Begin VB.TextBox tT 
            Height          =   285
            Left            =   840
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton bY 
            Caption         =   "C&alcular tirante"
            Height          =   375
            Left            =   480
            TabIndex        =   1
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Tirante "
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "m"
            Height          =   255
            Left            =   2160
            TabIndex        =   49
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ingrese las const. en F(x) al revestimiento "
         Height          =   2055
         Left            =   6000
         TabIndex        =   23
         Top             =   120
         Width           =   3255
         Begin VB.ComboBox Combo1 
            ForeColor       =   &H00004000&
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Text            =   "Constante de Etcheverry"
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox Combo2 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   240
            TabIndex        =   7
            Text            =   "Const. de Pavlovski y Kostiakov"
            Top             =   600
            Width           =   2655
         End
         Begin VB.ComboBox Combo3 
            ForeColor       =   &H00004000&
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Text            =   "Constante de Davis- Wilson"
            Top             =   960
            Width           =   2655
         End
         Begin VB.ComboBox Combo4 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   240
            TabIndex        =   9
            Text            =   "Constante de Punjab"
            Top             =   1320
            Width           =   2655
         End
         Begin VB.ComboBox Combo5 
            ForeColor       =   &H00004000&
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Text            =   "Constante de Moritz"
            Top             =   1680
            Width           =   2655
         End
      End
      Begin VB.CommandButton bC 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   3960
         Picture         =   "frminfiltracioncanales.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox tL 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Caudal  "
         Height          =   255
         Left            =   3000
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Solera "
         Height          =   255
         Left            =   3000
         TabIndex        =   54
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Talud"
         Height          =   255
         Left            =   3000
         TabIndex        =   53
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "m3/s"
         Height          =   255
         Left            =   5400
         TabIndex        =   52
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "m"
         Height          =   255
         Left            =   5400
         TabIndex        =   51
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "km"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Longitud "
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   36
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
   Begin VB.Label Label1 
      Caption         =   "Cálculo de Infiltraciones en canales "
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
      TabIndex        =   20
      Top             =   360
      Width           =   4575
   End
   Begin VB.Menu mpasdsf 
      Caption         =   "Parámetros  Suelo - Clima"
      Begin VB.Menu mgs 
         Caption         =   "Generales Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcond 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu meva 
         Caption         =   "Evapotranspiración"
      End
   End
   Begin VB.Menu h 
      Caption         =   "Hidráulica de Canales"
   End
   Begin VB.Menu mmp 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frminfiltracioncanales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BC_Click()
On Error GoTo mensaje

Y = Val(tT.text)
q = Val(tQ.text)
b = Val(tB.text)
z = Val(tZ.text)
ce = Combo1.text
k = Combo2.text
cd = Combo3.text
cp = Combo4.text
cm = Combo5.text
If Y = 0 Then
MsgBox "Ingrese el valor del tirante", 64, "Infiltración en canales"
tT.SetFocus
Exit Sub
End If
If q = 0 Then
MsgBox "Ingrese el valor del caudal", 64, "Infiltración en canales"
tQ.SetFocus
Exit Sub
End If
If tB = "" Then
MsgBox "Ingrese el valor de la solera (si es triangular éste es igual a cero)", 64, "Infiltración en canales"
tB.SetFocus
Exit Sub
End If
If tZ = "" Then
MsgBox "Ingrese el valor del talud (si es rectangular éste es igual a cero)", 64, "Infiltración en canales"
tZ.SetFocus
Exit Sub
End If

    a = (b + z * Y) * Y
    V = q / a
    Rem------------------------------------------------------------
    Rem seleccion de ce
    Select Case ce
     Case "Arcillosos"
        ce = 0.375
     Case "Franco arcillosos"
        ce = 0.625
     Case "Limosos y francos"
        ce = 0.875
     Case "Franco arenosos"
        ce = 1.25
     Case "Arenas finas"
        ce = 1.625
     Case "Arenas gruesas"
        ce = 2.25
     Case "Gravas"
        ce = 4.55
    End Select
    If Combo1.text = "Constante de Etcheverry" Then
    MsgBox "Seleccione el valor de la constante de Etcheverry", 64, "Infiltración en canales"
    Combo1.SetFocus
    Exit Sub
    End If

    Rem seleccion de k
    Select Case k
     Case "Grava"
        k = 500.5
     Case "Arena gruesa"
        k = 0.505
     Case "Arena fina"
        k = 0.0505
     Case "Tierra arenosa"
        k = 0.00505
     Case "Tierra franco arcillosa"
        k = 0.000050005
     Case "Tierra franca"
        k = 0.0005005
     Case "Limo"
        k = 0.00055
     Case "Arcilla"
        k = 0.00000505
     Case "Arcilla compacta"
        k = 0.0000005005
    End Select
    If Combo2.text = "Const. de Pavlovski y Kostiakov" Then
    MsgBox "Seleccione el valor de la constante de Pavlovski", 64, "Infiltración en canales"
    Combo2.SetFocus
    Exit Sub
    End If

    Rem seleccion de cd
    Select Case cd
     Case "Hormigón 10cm"
        cd = 1
     Case "Arcilla 15cm"
        cd = 4
     Case "Enlucido de cemento 2.5cm"
        cd = 6
     Case "Suelo arcilloso"
        cd = 12
     Case "Suelo franco arcilloso"
        cd = 15
     Case "Suelo franco"
        cd = 20
     Case "Suelo franco arenoso"
        cd = 25
     Case "Suelo arcillo limoso"
        cd = 30
     Case "Arena"
        cd = 45
    End Select
    If Combo3.text = "Constante de Davis- Wilson" Then
    MsgBox "Seleccione el valor de la constante de Davis-Wilson", 64, "Infiltración en canales"
    Combo3.SetFocus
    Exit Sub
    End If

    Rem seleccion de cp
    Select Case cp
     Case "Suelos muy permeables"
        cp = 0.03
     Case "Suelos comunes (medios)"
        cp = 0.02
     Case "Suelos impermeables"
        cp = 0.01
    End Select
    If Combo4.text = "Constante de Punjab" Then
    MsgBox "Seleccione el valor de la constante de Punjab", 64, "Infiltración en canales"
    Combo4.SetFocus
    Exit Sub
    End If

    Rem seleccion de cm
    Select Case cm
     Case "Franco arcilloso impermeable"
        cm = 0.095
     Case "Franco arcilloso semi-impermeable"
        cm = 0.13
     Case "Franco arcilloso ordinario, limo"
        cm = 0.19
     Case "Franco arcilloso con arena o grava"
        cm = 0.265
     Case "Franco arenoso"
        cm = 0.375
     Case "Suelos arenosos sueltos"
        cm = 0.5
     Case "Suelos arenosos con grava"
        cm = 0.65
     Case "Roca desintegrada con grava"
        cm = 0.825
     Case "Suelo con mucha grava"
        cm = 1.4
    End Select
    If Combo5.text = "Constante de Moritz" Then
    MsgBox "Seleccione el valor de la constante de Moritz", 64, "Infiltración en canales"
    Combo4.SetFocus
    Exit Sub
    End If

    Rem--------------------------------------------------
    Rem ingham
    p1 = 0.0025 * (Y) ^ (1 / 2) * (b + 2 * z * Y)
    
    Rem etcheverry
    p2 = 0.0064 * ce * (Y) ^ (1 / 2) * (b + 1.33 * Y * (1 + z ^ 2) ^ (1 / 2))
    
    Rem pavloski
    p3 = 10 * k * (b + 2 * Y * (1 + z))
    
    Rem davis wilson
    
    p4 = (cd * (Y) ^ (1 / 3) * (b + 2 * Y * (1 + z ^ 2) ^ (1 / 2))) / (8861 + 8 * (V) ^ (1 / 2))
    
    Rem Punjab
    p5 = cp * q ^ 0.563
    
    Rem kostiakov
    p6 = 10 * k * (b + 2.4 * Y * (1 + z ^ 2) ^ (1 / 2))
    
    Rem
    p7 = 0.0375 * cm * a ^ (1 / 2)
    
    Rem---------------------------------------------
    p11 = Format(p1, "##0.0000###")
    p12 = Format(p2, "##0.0000###")
    p13 = Format(p3, "##0.0000###")
    p14 = Format(p4, "##0.0000###")
    p15 = Format(p5, "##0.0000###")
    p16 = Format(p6, "##0.0000###")
    p17 = Format(p7, "##0.0000###")
    
    
    With gDatos
        .TextMatrix(1, 1) = p11
        .TextMatrix(2, 1) = p12
        .TextMatrix(3, 1) = p13
        .TextMatrix(4, 1) = p14
        .TextMatrix(5, 1) = p15
        .TextMatrix(6, 1) = p16
        .TextMatrix(7, 1) = p17
    End With

Frame4.Visible = True

StatusBar1.Panels(1).text = "Seleccione los valores de pérdidas que considere promediar, oprima el botón de Evaluar para calcular el caudal final "
Exit Sub
mensaje:
MsgBox "Ingrese adecuadamente los datos", 64, "Infiltración en canales"

End Sub

Private Sub BL_Click()
tQ.text = ""
tL.text = ""
tT.text = ""
tB.text = ""
tZ.text = ""
tQ1.text = ""
tQ2.text = ""
txtp.text = ""
Text1.text = ""

tP.text = ""
c1.Value = 0
c2.Value = 0
c3.Value = 0
c4.Value = 0
c5.Value = 0
c6.Value = 0
c7.Value = 0
With gDatos
    .Clear
    .TextMatrix(0, 0) = "Fórmula"
    .TextMatrix(0, 1) = "Pérdidas (m3/s*km)"
    .TextMatrix(1, 0) = "Ingham"
    .TextMatrix(2, 0) = "Etcheverry"
    .TextMatrix(3, 0) = "Pavlovski"
    .TextMatrix(4, 0) = "Davis-Wilson"
    .TextMatrix(5, 0) = "Punjab"
    .TextMatrix(6, 0) = "Kostiakov"
    .TextMatrix(7, 0) = "Moritz"
End With
tQ.SetFocus
Frame4.Visible = False


End Sub

Private Sub bP_Click()
On Error GoTo mensaje
q = Val(tQ.text)
L = Val(tL.text)
sx = 0
n = 0

    If c1.Value = 1 Then
     sx = sx + Val(gDatos.TextMatrix(1, 1))
     n = n + 1
    End If
    If c2.Value = 1 Then
     sx = sx + Val(gDatos.TextMatrix(2, 1))
     n = n + 1
    End If
    If c3.Value = 1 Then
     sx = sx + Val(gDatos.TextMatrix(3, 1))
     n = n + 1
    End If
    If c4.Value = 1 Then
     sx = sx + Val(gDatos.TextMatrix(4, 1))
     n = n + 1
    End If
    If c5.Value = 1 Then
     sx = sx + Val(gDatos.TextMatrix(5, 1))
     n = n + 1
    End If
    If c6.Value = 1 Then
     sx = sx + Val(gDatos.TextMatrix(6, 1))
     n = n + 1
    End If
    If c7.Value = 1 Then
     sx = sx + Val(gDatos.TextMatrix(7, 1))
     n = n + 1
    End If

    If n = 0 Then
    MsgBox "Seleccione los valores que desea promediar", 64, "Pérdidas en canales"
    Exit Sub
    End If
    
    If L = 0 Then
    MsgBox "Seleccione la longitud en kilómetros del canal", 64, "Pérdidas en canales"
    tL.SetFocus
    Exit Sub
    End If
    
        pr = sx / n
        tP.text = Format(pr, "##0.0000###")
        Rem calculo de caudales--------------------------------------
        R = pr / q
        
        
        
        Rem---------
        cp = Combo4.text
        Select Case cp
         Case "Suelos muy permeables"
            NN = 0.5
         Case "Suelos comunes (medios)"
            NN = 0.4
         Case "Suelos impermeables"
            NN = 0.2
        End Select
        Rem------------------
        aa = R * q ^ NN
        q1 = (1 - R * L) * q
        
        If q1 <= 0 Then
        MsgBox "Se pierde mas agua que la que circula", 13, "Pérdidas en canales"
        tL.SetFocus
        Exit Sub
        End If
        
        If NN = 0.5 And q1 < 0 Then
            MsgBox "No lluega Agua al final del canal", 64, " Infiltración en canales"
        Else
            rr = aa / ((q1) ^ (NN))
            
            rrr = (R + rr) / 2
            
            q2 = q * (1 - rrr * L)
            tQ2.text = Format(q2, "##0.00##")
            
            porcentaje = 100 - ((q - q2) / q * 100)
            por = ((q - q2) / q * 100)
            txtp.text = Format(porcentaje, "##0.00##")
            Text1.text = Format(por, "##0.00##")
            tQ1.text = Format(q1, "##0.00##")
        End If
        
    
Exit Sub
mensaje:
    MsgBox "Seleccione datos adecuadamente", 64, "Pérdidas en canales"
    

End Sub

Private Sub bS_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub bY_Click()
Frmhidraulica.Show
End Sub


Private Sub Command1_Click()
Print Form
End Sub

Private Sub Form_Load()
tT.text = ""

With Combo1
    .AddItem "Arcillosos"
    .AddItem "Franco arcillosos"
    .AddItem "Limosos y francos"
    .AddItem "Franco arenosos"
    .AddItem "Arenas finas"
    .AddItem "Arenas gruesas"
    .AddItem "Gravas"
End With

With Combo2
    .AddItem "Grava"
    .AddItem "Arena gruesa"
    .AddItem "Arena fina"
    .AddItem "Tierra arenosa"
    .AddItem "Tierra franco arcillosa"
    .AddItem "Tierra franca"
    .AddItem "Limo"
    .AddItem "Arcilla"
    .AddItem "Arcilla compacta"
End With

With Combo3
    .AddItem "Hormigón 10cm"
    .AddItem "Arcilla 15cm"
    .AddItem "Enlucido de cemento 2.5cm"
    .AddItem "Suelo arcilloso"
    .AddItem "Suelo franco arcilloso"
    .AddItem "Suelo franco"
    .AddItem "Suelo franco arenoso"
    .AddItem "Suelo arcillo limoso"
    .AddItem "Arena"
End With

With Combo4
    .AddItem "Suelos muy permeables"
    .AddItem "Suelos comunes (medios)"
    .AddItem "Suelos impermeables"
End With

With Combo5
    .AddItem "Franco arcilloso impermeable"
    .AddItem "Franco arcilloso semi-impermeable"
    .AddItem "Franco arcilloso ordinario, limo"
    .AddItem "Franco arcilloso con arena o grava"
    .AddItem "Franco arenoso"
    .AddItem "Suelos arenosos sueltos"
    .AddItem "Suelos arenosos con grava"
    .AddItem "Roca desintegrada con grava"
    .AddItem "Suelo con mucha grava"
End With

With gDatos
    .TextMatrix(0, 0) = "Fórmula"
    .TextMatrix(0, 1) = "Pérd. (m3/s*km)"
    .TextMatrix(1, 0) = "Ingham"
    .TextMatrix(2, 0) = "Etcheverry"
    .TextMatrix(3, 0) = "Pavlovski"
    .TextMatrix(4, 0) = "Davis-Wilson"
    .TextMatrix(5, 0) = "Punjab"
    .TextMatrix(6, 0) = "Kostiakov"
    .TextMatrix(7, 0) = "Moritz"
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
End With

StatusBar1.Panels(1).text = "Digite los valores que caracterizan al canal y seleccione todas las condiciones del revestimiento"
End Sub




Private Sub h_Click()
Frmhidraulica.Show
End Sub

Private Sub mcond_Click()
frmconductividad.Show

End Sub



Private Sub meva_Click()
frmETO.Show
End Sub

Private Sub mgs_Click()
frmgenerales.Show


End Sub

Private Sub mmp_Click()
frmGeneral.Show
Unload Me

End Sub



Private Sub mt_Click()
frmtextura.Show
End Sub
