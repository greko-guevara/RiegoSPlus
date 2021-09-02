VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmcombDia 
   Caption         =   "Combinación de Diámetros"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   Icon            =   "frmcombDia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11850
   Begin VB.Frame Frame10 
      Height          =   1095
      Left            =   2760
      TabIndex        =   10
      Top             =   6480
      Width           =   6255
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú Principal"
         Height          =   735
         Left            =   4440
         Picture         =   "frmcombDia.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2400
         MaskColor       =   &H000000FF&
         Picture         =   "frmcombDia.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   240
         Picture         =   "frmcombDia.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7785
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
            TextSave        =   "14/05/2007"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tuberias sin salidas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tuberias con salidas"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Combinación de diámetros en tuberías sin salida"
      ForeColor       =   &H00800000&
      Height          =   5295
      Left            =   840
      TabIndex        =   48
      Top             =   960
      Width           =   10095
      Begin VB.TextBox txtp 
         Height          =   285
         Left            =   5880
         TabIndex        =   17
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Height          =   2295
         Left            =   4320
         TabIndex        =   60
         Top             =   2760
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtvv2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   3600
            TabIndex        =   77
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtvv1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   3600
            TabIndex        =   76
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtll2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   65
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txthftt2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   64
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txthftt 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   63
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtll1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   62
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txthftt1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   61
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label40 
            Caption         =   "m/s"
            Height          =   255
            Left            =   4920
            TabIndex        =   81
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label38 
            Caption         =   "Velocidad tramo 2"
            Height          =   255
            Left            =   3600
            TabIndex        =   80
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label27 
            Caption         =   "m/s"
            Height          =   255
            Left            =   4920
            TabIndex        =   79
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label26 
            Caption         =   "Velocidad tramo 1"
            Height          =   255
            Left            =   3600
            TabIndex        =   78
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label43 
            Caption         =   "m"
            Height          =   255
            Left            =   1560
            TabIndex        =   75
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label25 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   74
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label fsf 
            Caption         =   "Longitud tramo 2"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label24 
            Caption         =   "Pérdidas tramo 2"
            Height          =   255
            Left            =   2040
            TabIndex        =   72
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   71
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "m"
            Height          =   255
            Left            =   1560
            TabIndex        =   70
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label21 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   69
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label20 
            Caption         =   "Pérdidas efectivas en todo el tramo"
            Height          =   255
            Left            =   360
            TabIndex        =   68
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label19 
            Caption         =   "Longitud tramo 1"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label16 
            Caption         =   "Pérdidas tramo 1"
            Height          =   255
            Left            =   2040
            TabIndex        =   66
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtll 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Diámetro Inferior"
         Height          =   735
         Left            =   5880
         TabIndex        =   55
         Top             =   840
         Width           =   3615
         Begin VB.TextBox txtdd2 
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "mm"
            Height          =   255
            Left            =   3000
            TabIndex        =   57
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label28 
            Caption         =   "Diámetro"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Combinar diámetros"
         Height          =   735
         Left            =   1440
         MaskColor       =   &H008080FF&
         Picture         =   "frmcombDia.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Diámetro Superior"
         Height          =   735
         Left            =   480
         TabIndex        =   49
         Top             =   840
         Width           =   3615
         Begin VB.TextBox txtdd1 
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label34 
            Caption         =   "Diámetro"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label32 
            Caption         =   "mm"
            Height          =   255
            Left            =   3000
            TabIndex        =   50
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   6360
         TabIndex        =   13
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtqq 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Left            =   360
         Picture         =   "frmcombDia.frx":29F2
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   3465
      End
      Begin VB.Line Line1 
         X1              =   9720
         X2              =   240
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label47 
         Caption         =   "m"
         Height          =   255
         Left            =   7200
         TabIndex        =   83
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label41 
         Caption         =   "Pérdidas admisibles"
         Height          =   255
         Left            =   4200
         TabIndex        =   82
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label37 
         Caption         =   "Longitud del tramo"
         Height          =   255
         Left            =   480
         TabIndex        =   59
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "m"
         Height          =   255
         Left            =   3240
         TabIndex        =   58
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label46 
         Caption         =   "Seleccione el Material a utilizar"
         Height          =   375
         Left            =   4080
         TabIndex        =   54
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label45 
         Caption         =   "Caudal del sistema"
         Height          =   255
         Left            =   480
         TabIndex        =   53
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   "m3/dia"
         Height          =   255
         Left            =   3120
         TabIndex        =   52
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Combinación de diámetros en tuberias con salidas"
      ForeColor       =   &H00800000&
      Height          =   5295
      Left            =   840
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox xxx 
         Height          =   285
         Left            =   8280
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtl 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txth 
         Height          =   285
         Left            =   5280
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Frame Frame14 
         Height          =   2175
         Left            =   4320
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox Text4 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   89
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   88
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtv1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   3720
            TabIndex        =   43
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtv2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   3720
            TabIndex        =   42
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txthft1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   35
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txthft 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   34
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txthft2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   33
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label49 
            Caption         =   "Longitud tramo 1"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label48 
            Caption         =   "m"
            Height          =   255
            Left            =   1560
            TabIndex        =   92
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label18 
            Caption         =   "Longitud tramo 2"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "m"
            Height          =   255
            Left            =   1560
            TabIndex        =   90
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "Velocidad tramo 1"
            Height          =   255
            Left            =   3720
            TabIndex        =   47
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "m/s"
            Height          =   255
            Left            =   5040
            TabIndex        =   46
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "Velocidad tramo 2"
            Height          =   255
            Left            =   3720
            TabIndex        =   45
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "m/s"
            Height          =   255
            Left            =   5040
            TabIndex        =   44
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label33 
            Caption         =   "Pérdidas tramo 1"
            Height          =   255
            Left            =   2040
            TabIndex        =   41
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label35 
            Caption         =   "Pérdidas efectivas en todo el tramo"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label Label36 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   39
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label39 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   38
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label fsfff 
            Caption         =   "Pérdidas tramo 2"
            Height          =   255
            Left            =   2040
            TabIndex        =   37
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   36
            Top             =   1200
            Width           =   255
         End
      End
      Begin VB.TextBox txtQ 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6360
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
      Begin VB.Frame Frame3 
         Caption         =   "Diámetro Superior"
         Height          =   735
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   3615
         Begin VB.TextBox txtd1 
            Height          =   285
            Left            =   1680
            TabIndex        =   2
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "mm"
            Height          =   255
            Left            =   3000
            TabIndex        =   28
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Diámetro"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Diámetro Inferiror"
         Height          =   855
         Left            =   5880
         TabIndex        =   23
         Top             =   840
         Width           =   3615
         Begin VB.TextBox txtd2 
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Diámetro"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "mm"
            Height          =   255
            Left            =   3000
            TabIndex        =   24
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.CommandButton bcombinacion 
         Caption         =   "Combinar diámetros"
         Height          =   735
         Left            =   1440
         MaskColor       =   &H008080FF&
         Picture         =   "frmcombDia.frx":934C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1290
         Left            =   360
         Picture         =   "frmcombDia.frx":9AB6
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   3420
      End
      Begin VB.Label Label50 
         Caption         =   "Espaciamiento"
         Height          =   255
         Left            =   7080
         TabIndex        =   94
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "m"
         Height          =   255
         Left            =   3120
         TabIndex        =   87
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Longitud del tramo"
         Height          =   255
         Left            =   360
         TabIndex        =   86
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Pérdidas admisibles"
         Height          =   255
         Left            =   3720
         TabIndex        =   85
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "m"
         Height          =   255
         Left            =   6600
         TabIndex        =   84
         Top             =   2040
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   9720
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label13 
         Caption         =   "m3/dia"
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Caudal por salida"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label30 
         Caption         =   "Seleccione el Material a utilizar"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Combinación de diámetros"
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
      Left            =   1560
      TabIndex        =   21
      Top             =   360
      Width           =   3855
   End
   Begin VB.Menu vocehdt 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu mcallat 
         Caption         =   "Cálculos en laterales"
      End
      Begin VB.Menu mcalprin 
         Caption         =   "Cálculos en la principal"
      End
      Begin VB.Menu mequibomb 
         Caption         =   "Equipo de bombeo"
      End
   End
   Begin VB.Menu mm 
      Caption         =   "Menu Principal"
   End
End
Attribute VB_Name = "frmcombDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qs As Double
Dim L As Double
Dim ysup As Double
Dim ymin As Double
Dim easp As Double
Dim hff As Double
Dim l1 As Double
Dim l2 As Double
Dim n As Double
Dim n1 As Double
Dim n2 As Double
Dim n3 As Double
Dim n4 As Double
Dim qt As Double

Private Sub bcombinacion_Click()


On Error GoTo mensaje:
qs = Val(txtQ.text)
L = Val(txtL.text)
ysup = Val(txtd1.text)
hff = Val(txtH.text)
ymin = Val(txtd2.text)
easp = Val(xxx.text)
If qs = 0 Then
    MsgBox "Ingrese el valor del caudal", 64, "Combinación de diámetros"
    txtQ.SetFocus
    Exit Sub
End If
If L = 0 Then
    MsgBox "Ingrese la longitud", 64, "Combinación de diámetros"
    txtL.SetFocus
    Exit Sub
End If

If hff = 0 Then
    MsgBox "Ingrese el valor de las pérdidas admisibles", 64, "Combinación de diámetros"
    txtH.SetFocus
    Exit Sub
End If
If ysup = 0 Then
    MsgBox "Ingrese el valor del diámetro superior", 64, "Combinación de diámetros"
    txtd1.SetFocus
    Exit Sub
End If
If ymin = 0 Then
    MsgBox "Ingrese el valor del diámetro inferior", 64, "Combinación de diámetros"
    txtd2.SetFocus
    Exit Sub
End If


n = L / easp
qt = qs * n

c = Combo1.ListIndex
If c = -1 Then
    MsgBox "Seleccione el tipo de tubería", 64, "Combinación de diámetros"
    Combo1.SetFocus
    Exit Sub
End If
Select Case c
    Case Is = 0
    c = 140
    m = 1.852
    Case Is = 1
    c = 140
    m = 1.852
    Case Is = 2
    c = 120
    m = 1.852
    Case Is = 3
    c = 110
    m = 1.852
    Case Is = 4
    c = 120
    m = 1.852
    Case Is = 5
    c = 115
    m = 1.852
    Case Is = 6
    c = 150
    m = 1.76
    Case Is = 7
    c = 140
    m = 1.76
End Select
            
   
    f = 2 * n / (2 * n - 1) * (1 / (m + 1) + (m - 1) ^ 0.5 / (6 * n ^ 2))


hf1 = 1.131 * 10 ^ 9 * (qt / c) ^ 1.852 * ysup ^ -4.872 * L * f

For j% = 1 To n - 1
n3 = n - j%
n2 = n - n3

l1 = L - j% * easp
l2 = L - l1


    f2 = (1 / (m + 1) + 1 / (2 * n2) + (m - 1) ^ 0.5 / (6 * n2 ^ 2))
q2 = qs * n2
hf2 = 1.131 * 10 ^ 9 * (q2 / c) ^ 1.852 * ysup ^ -4.872 * l2 * f2

hft1 = hf1 - hf2
hft2 = 1.131 * 10 ^ 9 * (q2 / c) ^ 1.852 * ymin ^ -4.872 * l2 * f2
hft = hft1 + hft2

If hft > hff Then
If j% = 1 Then
MsgBox "Ambos diámetros superan las pérdidas admisibles. Trabaje con diámetros superiores", 64, "Combinación de diámetros"
Exit Sub
End If

n4 = n - (j% - 1)
n2 = n - n4
l1 = L - (j% - 1) * easp
l2 = L - l1
    f2 = (1 / (m + 1) + 1 / (2 * n2) + (m - 1) ^ 0.5 / (6 * n2 ^ 2))

q2 = qs * n2
hf2 = 1.131 * 10 ^ 9 * (q2 / c) ^ 1.852 * ysup ^ -4.872 * l2 * f2

hft1 = hf1 - hf2
hft2 = 1.131 * 10 ^ 9 * (q2 / c) ^ 1.852 * ymin ^ -4.872 * l2 * f2
hft = hft1 + hft2
j% = n - 1
End If
Next j%



v1 = qt / (3.14159 / 4 * (ysup / 1000) ^ 2) * 1 / 3600
v2 = q2 / (3.14159 / 4 * (ymin / 1000) ^ 2) * 1 / 3600
Text4 = Format(L - l2, "##0.0##")
Text3 = Format(l2, "##0.0##")
txtv1 = Format(v1, "##0.0##")
txtv2 = Format(v2, "##0.0##")
txthft1 = Format(hft1, "##0.0##")
txthft2 = Format(hft2, "##0.0##")
txthft = Format(hft, "##0.0##")
Frame14.Visible = True
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub

Private Sub bfinailizar_Click()
frmGeneral.Show
Unload Me
End Sub

Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtQ.text = ""
txtqq.text = ""


txtLL.text = ""
txtd1.text = ""
txtd2.text = ""
txtdd1.text = ""
txtdd2.text = ""
txtv1.text = ""
txtvv1.text = ""
txtvv2.text = ""
txtv2.text = ""
txthft1.text = ""
txthftt1.text = ""
txthft2.text = ""
txthftt2.text = ""
txthft.text = ""
txthftt.text = ""
Combo1.text = ""


Combo4.text = ""
txtll1.text = ""
txtll2.text = ""


txtp.text = ""

Frame5.Visible = False
Frame14.Visible = False

End Sub

Private Sub Command1_Click()
On Error GoTo mensaje:
qq = Val(txtqq.text)
LL = Val(txtLL.text)
dd1 = Val(txtdd1.text)
dd2 = Val(txtdd2.text)
P = Val(txtp.text)

If qq = 0 Then
    MsgBox "Ingrese el valor del caudal", 64, "Combinación de diámetros"
    txtqq.SetFocus
    Exit Sub
End If
If LL = 0 Then
    MsgBox "Ingrese el valor de la longitud", 64, "Combinación de diámetros"
    txtLL.SetFocus
    Exit Sub
End If
If dd1 = 0 Then
    MsgBox "Ingrese el valor del diámetro superior", 64, "Combinación de diámetros"
    txtdd1.SetFocus
    Exit Sub
End If
If dd2 = 0 Then
    MsgBox "Ingrese el valor del diámetro inferior", 64, "Combinación de diámetros"
    txtdd2.SetFocus
    Exit Sub
End If
If P = 0 Then
    MsgBox "Ingrese el valor de las pérdidas admisibles", 64, "Combinación de diámetros"
    txtp.SetFocus
    Exit Sub
End If

ccc = Combo4.ListIndex
If ccc = -1 Then
    MsgBox "Seleccione el tipo de tubería", 64, "Combinación de diámetros"
    Combo4.SetFocus
    Exit Sub
End If

Select Case ccc
    Case Is = 0
    ccc = 140
    m = 1.852
    Case Is = 1
    ccc = 140
    m = 1.852
    Case Is = 2
    ccc = 120
    m = 1.852
    Case Is = 3
    ccc = 110
    m = 1.852
    Case Is = 4
    ccc = 120
    m = 1.852
    Case Is = 5
    ccc = 115
    m = 1.852
    Case Is = 6
    ccc = 150
    m = 1.76
    Case Is = 7
    ccc = 140
    m = 1.76
End Select


hff1 = 1.131 * 10 ^ 9 * (qq / ccc) ^ 1.852 * dd1 ^ -4.872
   
hff2 = 1.131 * 10 ^ 9 * (qq / ccc) ^ 1.852 * dd2 ^ -4.872

If hff1 > P / LL Then
    MsgBox "En ambos diámetros se superan las pérdidas admisibles. Trabaje con diámetros más grandes.", 64, "Combinación de diámetros"
    Exit Sub
End If

If hff2 < P / LL Then
    MsgBox "En ambos diámetros no se superan las pérdidas admisibles. Trabaje con diámetros más pequeños", 64, "Combinación de diámetros"
    Exit Sub
End If

Ll1 = (P - (hff2 * LL)) / (hff1 - hff2)
ll2 = LL - Ll1
hftt1 = hff1 * Ll1
hftt2 = hff2 * ll2
hftt = hftt1 + hftt2
txtll1 = Format(Ll1, "##0.0#")
txtll2 = Format(ll2, "##0.0#")
txthftt1 = Format(hftt1, "##0.0##")
txthftt2 = Format(hftt2, "##0.0##")
txthftt = Format(hftt, "##0.0##")
vv1 = qq / (3.14159 / 4 * (dd1 / 1000) ^ 2) * 1 / 3600
vv2 = qq / (3.14159 / 4 * (dd2 / 1000) ^ 2) * 1 / 3600
txtvv1 = Format(vv1, "##0.0##")
txtvv2 = Format(vv2, "##0.0##")

Frame5.Visible = True
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub

Private Sub Form_Load()
With Combo4
    .AddItem "Acero Nuevo (C= 140)"
    .AddItem "Aluminio Nuevo (C= 140)"
    .AddItem "Acero viejo 15 años (C= 120)"
    .AddItem "Acero remachado 10 años (C= 110)"
    .AddItem "Aluminio con acoples (C= 120)"
    .AddItem "Galvanizado con uniones (C= 115)"
    .AddItem "P.V.C. (C= 150)"
    .AddItem "Polietileno (C= 140)"
End With
With Combo1
    .AddItem "Acero Nuevo (C= 140)"
    .AddItem "Aluminio Nuevo (C= 140)"
    .AddItem "Acero viejo 15 años (C= 120)"
    .AddItem "Acero remachado 10 años (C= 110)"
    .AddItem "Aluminio con acoples (C= 120)"
    .AddItem "Galvanizado con uniones (C= 115)"
    .AddItem "P.V.C. (C= 150)"
    .AddItem "Polietileno (C= 140)"
End With

StatusBar1.Panels(1).text = "Digite los datos básicos y oprima el botón de Combinación de diámetros para iniciar los cálculos"
End Sub




Private Sub mcallat_Click()
FrmHLaterales.Show
End Sub

Private Sub mcalprin_Click()
frmHprincipal.Show
End Sub

Private Sub mequibomb_Click()
frmbomba.Show
End Sub

Private Sub mm_Click()
frmGeneral.Show
End Sub

Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    Frame4.Visible = True
    txtqq.SetFocus
    Frame2.Visible = False
    Case 2
    Frame2.Visible = True
    Frame4.Visible = False
    txtQ.SetFocus
    End Select
End Sub

