VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmETO 
   Caption         =   "Evapotranspiración"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "frmETO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   1800
      TabIndex        =   39
      Top             =   6360
      Width           =   8055
      Begin VB.CommandButton bevaluar 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   360
         Picture         =   "frmETO.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   6120
         Picture         =   "frmETO.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   4200
         Picture         =   "frmETO.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   2280
         Picture         =   "frmETO.frx":2308
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   720
      TabIndex        =   12
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Método Hargreaves"
      TabPicture(0)   =   "frmETO.frx":29F2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tanque Vaporímetro"
      TabPicture(1)   =   "frmETO.frx":2A0E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "Frame10"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(5)=   "Label21"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame3 
         Caption         =   "Resultados"
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   4800
         TabIndex        =   60
         Top             =   600
         Width           =   4215
         Begin VB.TextBox txtra 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1680
            TabIndex        =   63
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txteto 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1680
            TabIndex        =   62
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtetr 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1680
            TabIndex        =   61
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "Radiación Extraterrestre"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label25 
            Caption         =   "mm/día"
            Height          =   255
            Left            =   3360
            TabIndex        =   69
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   2760
            TabIndex        =   68
            Top             =   -240
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "Evapotranspiración del Cultivo de Referencia ETO"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Label17 
            Caption         =   "mm/día"
            Height          =   255
            Left            =   3360
            TabIndex        =   66
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "Evapotraspiración Real ETR"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   1560
            Width           =   3495
         End
         Begin VB.Label Label24 
            Caption         =   "mm/día"
            Height          =   255
            Left            =   3360
            TabIndex        =   64
            Top             =   1800
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Tiempo"
         ForeColor       =   &H00800000&
         Height          =   1815
         Left            =   480
         TabIndex        =   54
         Top             =   3480
         Width           =   3735
         Begin VB.CommandButton Command1 
            Caption         =   "&Ver Kc"
            Height          =   255
            Left            =   2280
            TabIndex        =   57
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtkc 
            Height          =   285
            Left            =   960
            TabIndex        =   56
            Text            =   "1"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1440
            TabIndex        =   55
            Text            =   "Enero"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3720
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            Caption         =   "Kc"
            Height          =   255
            Left            =   480
            TabIndex        =   59
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            Caption         =   "Mes"
            Height          =   255
            Left            =   480
            TabIndex        =   58
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Temperaturas"
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   480
         TabIndex        =   47
         Top             =   1800
         Width           =   3735
         Begin VB.TextBox txttmin 
            Height          =   285
            Left            =   1800
            TabIndex        =   49
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txttmax 
            Height          =   285
            Left            =   1800
            TabIndex        =   48
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "°C"
            Height          =   255
            Left            =   3120
            TabIndex        =   53
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "°C"
            Height          =   255
            Left            =   3120
            TabIndex        =   52
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Mínima"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "Máxima"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ubicación"
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   480
         TabIndex        =   42
         Top             =   600
         Width           =   3735
         Begin VB.TextBox txtlat 
            Height          =   285
            Left            =   1080
            TabIndex        =   44
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox txtns 
            Height          =   315
            Left            =   2160
            TabIndex        =   43
            Text            =   "N"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "°"
            Height          =   255
            Left            =   2040
            TabIndex        =   46
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Caption         =   "Latitud"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   -70200
         Picture         =   "frmETO.frx":2A2A
         ScaleHeight     =   1995
         ScaleWidth      =   4155
         TabIndex        =   41
         Top             =   3240
         Width           =   4215
      End
      Begin VB.PictureBox Picture1 
         Height          =   2055
         Left            =   -70200
         Picture         =   "frmETO.frx":239C4
         ScaleHeight     =   1995
         ScaleWidth      =   4155
         TabIndex        =   40
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Frame Frame9 
         Caption         =   "Datos climáticos medios y de cultivo"
         ForeColor       =   &H00800000&
         Height          =   2535
         Left            =   -74520
         TabIndex        =   31
         Top             =   600
         Width           =   3735
         Begin VB.CommandButton Command2 
            Caption         =   "&Ver Kc"
            Height          =   255
            Left            =   2280
            TabIndex        =   6
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cbV 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cbHR 
            Height          =   315
            Left            =   1320
            TabIndex        =   5
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtkcc 
            Height          =   285
            Left            =   960
            TabIndex        =   7
            Text            =   "1"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtetan 
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3720
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3720
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label Label14 
            Caption         =   "Velocidad del viento"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Humedad Relativa"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "km/h"
            Height          =   255
            Left            =   3000
            TabIndex        =   36
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label19 
            Caption         =   "%"
            Height          =   255
            Left            =   3000
            TabIndex        =   35
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label12 
            Caption         =   "Kc"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "Evapotranspiración registrada"
            Height          =   495
            Left            =   240
            TabIndex        =   33
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label30 
            Caption         =   "mm/día"
            Height          =   255
            Left            =   3000
            TabIndex        =   32
            Top             =   2040
            Width           =   615
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Ubicación del Tanque"
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   -74520
         TabIndex        =   28
         Top             =   3240
         Width           =   3735
         Begin VB.OptionButton optbar 
            Caption         =   "Superficie en barbecho Seco"
            Height          =   195
            Left            =   360
            TabIndex        =   9
            Top             =   480
            Value           =   -1  'True
            Width           =   2895
         End
         Begin VB.OptionButton optver 
            Caption         =   "Superficie verde de poca altura"
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   840
            Width           =   2895
         End
         Begin VB.ComboBox cbD 
            Height          =   315
            Left            =   1320
            TabIndex        =   11
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Distancia a barlovento a la cual cambia la cobertura"
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   3720
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label29 
            Caption         =   "m"
            Height          =   255
            Left            =   3000
            TabIndex        =   29
            Top             =   1680
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Resultados"
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   -70200
         TabIndex        =   16
         Top             =   600
         Width           =   4215
         Begin VB.TextBox txtktan 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1680
            TabIndex        =   26
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtetr1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1680
            TabIndex        =   18
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txteto1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1680
            TabIndex        =   17
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "mm/día"
            Height          =   255
            Left            =   3360
            TabIndex        =   25
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Evapotraspiración Real ETR"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "mm/día"
            Height          =   255
            Left            =   3360
            TabIndex        =   23
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Evapotranspiración del Cultivo de Referencia ETO"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Label7 
            Caption         =   "Lámina Bruta"
            Height          =   255
            Left            =   2760
            TabIndex        =   21
            Top             =   -240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "mm/día"
            Height          =   255
            Left            =   3360
            TabIndex        =   20
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Coeficiente del tanque Vaporímetro"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2100
         Left            =   5520
         Picture         =   "frmETO.frx":4495E
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label21 
         Caption         =   "Método del Tanque Vaporímetro"
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
         Left            =   -69960
         TabIndex        =   14
         Top             =   0
         Width           =   3975
      End
      Begin VB.Label Label22 
         Caption         =   "Método de Hargreaves"
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
         Left            =   5640
         TabIndex        =   13
         Top             =   0
         Width           =   2775
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
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
            TextSave        =   "08/07/2005"
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
      Caption         =   "Determinación de la Evapotranspiración"
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
      Left            =   720
      TabIndex        =   15
      Top             =   120
      Width           =   6615
   End
   Begin VB.Menu msuelo 
      Caption         =   "Generalidades del Suelo"
      Begin VB.Menu mpara 
         Caption         =   "Parámetros Generales"
      End
      Begin VB.Menu mtext 
         Caption         =   "Textura"
      End
      Begin VB.Menu mcond 
         Caption         =   "Conductividad Hidráulica"
      End
   End
   Begin VB.Menu massis 
      Caption         =   "Calendarios de riego"
   End
   Begin VB.Menu mm 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmETO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bevaluar_Click()
On Error GoTo mensaje
Select Case SSTab1.Tab
Case 0
'HARGREAVES
    X = Val(txtlat.text)
    ns = (txtns.text)
    kc = Val(txtkc.text)
    Min = Val(txttmin.text)
    Max = Val(txttmax.text)
    mes = (Combo1.text)
    If X = 0 Then
    MsgBox "Introduzca el valor de la latitud", 64, "Determinación de evapotranspiración"
    txtlat.SetFocus
    Exit Sub
    End If
    If txttmax.text = "" Then
    MsgBox "Introduzca el valor de la temperatura máxima diurna", 64, "Determinación de evapotranspiración"
    txttmax.SetFocus
    Exit Sub
    End If
    If txttmin.text = "" Then
    MsgBox "Introduzca el valor de la temperatura mínima diurna", 64, "Determinación de evapotranspiración"
    txttmin.SetFocus
    Exit Sub
    End If
    'definicion de RA=y
    If ns <> "Ecuador" Then
        Select Case mes
        Case "Enero"
            If ns = "N" Then
                Y = -0.001 * X ^ 2 - 0.1784 * X + 15.066
            Else
                Y = -0.0021 * X ^ 2 + 0.1577 * X + 14.961
            End If
        Case "Febrero"
            If ns = "N" Then
                Y = -0.0014 * X ^ 2 - 0.1187 * X + 15.544
            Else
                Y = -0.0022 * X ^ 2 + 0.0927 * X + 15.518
            End If
        Case "Marzo"
            If ns = "N" Then
                Y = -0.0021 * X ^ 2 - 0.0229 * X + 15.651
            Else
                Y = -0.0021 * X ^ 2 + 0.0073 * X + 15.634
            End If
        Case "Abril"
            If ns = "N" Then
                Y = -0.0024 * X ^ 2 + 0.0703 * X + 15.163
            Else
                Y = -0.0016 * X ^ 2 - 0.0845 * X + 15.25
            End If
        Case "Mayo"
            If ns = "N" Then
                Y = -0.0026 * X ^ 2 + 0.1674 * X + 13.883
            Else
                Y = -0.0011 * X ^ 2 - 0.1517 * X + 14.416
            End If
        Case "Junio"
            If ns = "N" Then
                Y = -0.0019 * X ^ 2 + 0.1609 * X + 13.835
            Else
                Y = -0.0007 * X ^ 2 - 0.1867 * X + 13.943
            End If
        Case "Julio"
            If ns = "N" Then
                Y = -0.0022 * X ^ 2 + 0.1558 * X + 13.997
            Else
                Y = -0.0008 * X ^ 2 - 0.1752 * X + 14.164
            End If
        Case "Agosto"
            If ns = "N" Then
                Y = -0.0023 * X ^ 2 + 0.1002 * X + 14.722
            Else
                Y = -0.0015 * X ^ 2 - 0.1117 * X + 14.75
            End If
        Case "Setiembre"
            If ns = "N" Then
                Y = -0.0021 * X ^ 2 + 0.0155 * X + 15.301
            Else
                Y = -0.0021 * X ^ 2 - 0.0237 * X + 15.232
            End If
        Case "Octubre"
            If ns = "N" Then
                Y = -0.0017 * X ^ 2 - 0.0784 * X + 15.533
            Else
                Y = -0.0023 * X ^ 2 + 0.0623 * X + 15.417
            End If
        Case "Noviembre"
            If ns = "N" Then
                Y = -0.0012 * X ^ 2 - 0.1575 * X + 15.207
            Else
                Y = -0.0024 * X ^ 2 + 0.152 * X + 14.734
            End If
        Case "Diciembre"
            If ns = "N" Then
                Y = -0.0007 * X ^ 2 - 0.2007 * X + 14.936
            Else
                Y = -0.002 * X ^ 2 + 0.1695 * X + 14.736
            End If
        End Select
    Else
        Select Case mes
        Case "Enero"
            Y = 15
        Case "Febrero"
            Y = 15.5
        Case "Marzo"
            Y = 15.7
        Case "Abril"
            Y = 15.3
        Case "Mayo"
            Y = 13.9
        Case "Junio"
            Y = 13.9
        Case "Julio"
            Y = 14.1
        Case "Agosto"
            Y = 14.8
        Case "Setiembre"
            Y = 15.3
        Case "Octubre"
            Y = 15.4
        Case "Noviembre"
            Y = 15.1
        Case "Diciembre"
            Y = 14.8
        End Select
    End If
        'calculo de Eto y etr
        
    If Max <= Min Then
    MsgBox "El valor de la temperatura máxima debe ser mayor al valor de la temperatura mínima", 64, "Determinación de evapotranspiración"
    txttmax.SetFocus
    Exit Sub
    End If
    eto = 0.0023 * Y * ((Max + Min) / 2 + 17.8) * (Max - Min) ^ 0.5
    etr = eto * kc
   
    txteto.text = Format(eto, "##0.0#")
    txtetr.text = Format(etr, "##0.0#")
    txtra.text = Format(Y, "##0.0#")
Case 1
    'tanque evaporimetro
    d = cbD.text
    hr1 = cbHR.text
    v1 = cbV.text
    kcc = Val(txtkcc.text)
    etan = Val(txtetan.text)
    If cbV.text = "" Then
    MsgBox "selecione la velocidad del viento", 64, " Estimación de ETo y ETr"
    cbV.SetFocus
    Exit Sub
    End If
    If cbHR.text = "" Then
    MsgBox "Selecione la humedad relativa", 64, " Estimación de ETo y ETr"
    cbHR.SetFocus
    Exit Sub
    End If
    If etan = 0 Then
    MsgBox "selecione la evapotranspiración registrada en el tanque", 64, " Estimación de ETo y ETr"
    txtetan.SetFocus
    Exit Sub
    End If
    If cbD.text = "" Then
    MsgBox "selecione la distancia a que cambia la cobertura", 64, " Estimación de ETo y ETr"
    cbD.SetFocus
    Exit Sub
    End If
 
    Select Case hr1
        Case "< 40"
        hr = 5
        Case "40-70"
        hr = 6
        Case "> 70"
        hr = 7
    End Select
    
    Select Case v1
        Case "< 2"
        V = 1
        Case "2-5"
        V = 2
        Case "5-8"
        V = 3
        Case "> 8"
        V = 4
    End Select
    
    Select Case d
        Case "< 9"
        If optbar.Value = True Then
            ktan = 0.4458 + 0.688 * hr - 0.0683 * V
        Else
            ktan = 0.2667 + 0.075 * hr - 0.0667 * V
        End If
        Case "10-99"
        If optbar.Value = True Then
            ktan = 0.3137 + 0.075 * hr - 0.0667 * V
        Else
            ktan = 0.375 + 0.075 * hr - 0.0733 * V
        End If
        Case "100-99"
        If optbar.Value = True Then
            ktan = 0.2667 + 0.075 * hr - 0.0667 * V
        Else
            ktan = 0.4458 + 0.0688 * hr - 0.0683 * V
        End If
        Case "> 1000"
        If optbar.Value = True Then
            ktan = 0.2167 + 0.075 * hr - 0.0667 * V
        Else
            ktan = 0.6 + 0.05 * hr - 0.0717 * V
        End If
    End Select
    
    
    eto1 = etan * ktan
    etr1 = eto1 * kcc
    txteto1.text = Format(eto1, "##0.0#")
    txtetr1.text = Format(etr1, "##0.0#")
    txtktan.text = Format(ktan, "##0.0#")
    
End Select
Exit Sub
mensaje:
    MsgBox "Ingrese adecuadamente los datos de Entrada", 64, " Estimación de ETo y ETr"

    
    
End Sub



Private Sub bfinailizar_Click()
Unload Me

End Sub

Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
Unload Me
frmETO.Show
End Sub



Private Sub Command1_Click()
Dialog1.Show
End Sub

Private Sub Command2_Click()
Dialog1.Show
End Sub

Private Sub Form_Load()
optbar.ForeColor = &HC0&
With txtns
    .AddItem "N"
    .AddItem "S"
    .AddItem "Ecuador"
End With
With Combo1
    .AddItem "Enero"
    .AddItem "Febrero"
    .AddItem "Marzo"
    .AddItem "Abril"
    .AddItem "Mayo"
    .AddItem "Junio"
    .AddItem "Julio"
    .AddItem "Agosto"
    .AddItem "Setiembre"
    .AddItem "Octubre"
    .AddItem "Noviembre"
    .AddItem "Diciembre"
 End With
 
 With cbD
    .AddItem "< 9"
    .AddItem "10-99"
    .AddItem "100-99"
    .AddItem "> 1000"
End With

With cbV
    .AddItem "< 7.5"
    .AddItem "7.5-19"
    .AddItem "19-30"
    .AddItem "> 30"
End With

With cbHR
    .AddItem "< 40"
    .AddItem "40-70"
    .AddItem "> 70"
End With

   StatusBar1.Panels(1).text = "Digite los datos de Climáticos y de localización, luego Oprima Calcular Para Estimar ETo y ETr"

End Sub








Private Sub massis_Click()
frmCalendario.Show
End Sub

Private Sub mcond_Click()
frmconductividad.Show
End Sub

Private Sub mm_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub mpara_Click()
frmgenerales.Show
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

Private Sub mtext_Click()
frmtextura.Show
End Sub

Private Sub optbar_Click()
Picture1.Visible = False
Picture2.Visible = True
optver.ForeColor = &H80000012
optbar.ForeColor = &HC0&

End Sub

Private Sub optver_Click()
Picture2.Visible = False
Picture1.Visible = True
optver.ForeColor = &HC0&
optbar.ForeColor = &H80000012

End Sub
