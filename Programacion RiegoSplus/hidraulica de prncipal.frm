VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHprincipal 
   Caption         =   "Diseño Hidráulico de la Principal"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11400
   Icon            =   "hidraulica de prncipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11400
   Begin VB.Frame Frame1 
      Caption         =   "Datos básicos para el diseño"
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   173
      TabIndex        =   45
      Top             =   1200
      Visible         =   0   'False
      Width           =   11535
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtQ 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Bevaluar 
         Caption         =   "&Evaluar"
         Height          =   615
         Left            =   9720
         Picture         =   "hidraulica de prncipal.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   7560
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   7560
         TabIndex        =   46
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbletiqueta 
         Height          =   255
         Left            =   360
         TabIndex        =   97
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblunidades 
         Height          =   255
         Left            =   3960
         TabIndex        =   96
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label30 
         Caption         =   "Seleccione el Material a utilizar"
         Height          =   255
         Left            =   5280
         TabIndex        =   53
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Longitud del tramo"
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Caudal del sistema"
         Height          =   375
         Left            =   360
         TabIndex        =   51
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "m"
         Height          =   255
         Left            =   3960
         TabIndex        =   50
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "m3/h"
         Height          =   255
         Left            =   3960
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbl1 
         Caption         =   "Selec. Polietileno"
         Height          =   255
         Left            =   5280
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbl2 
         Caption         =   "Especifique el tipo de PVC"
         Height          =   255
         Left            =   5280
         TabIndex        =   47
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Qué desea calcular ?"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   6360
      TabIndex        =   42
      Top             =   240
      Width           =   4335
      Begin VB.OptionButton Option2 
         Caption         =   "Pérdidas en el tramo"
         Height          =   375
         Left            =   2160
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Díametro de la tuberia "
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1935
      End
      Begin VB.Line Line3 
         X1              =   2400
         X2              =   2640
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   2573
      TabIndex        =   41
      Top             =   6480
      Width           =   6255
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   240
         Picture         =   "hidraulica de prncipal.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2280
         Picture         =   "hidraulica de prncipal.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   4320
         Picture         =   "hidraulica de prncipal.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
      Begin VB.Frame Frame11 
         Caption         =   "PVC"
         ForeColor       =   &H00000080&
         Height          =   3735
         Left            =   360
         TabIndex        =   32
         Top             =   3960
         Width           =   5175
         Begin MSFlexGridLib.MSFlexGrid gridSDR17 
            Height          =   1335
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   13
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridSDR26 
            Height          =   1335
            Left            =   2640
            TabIndex        =   34
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   13
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridSDR41 
            Height          =   1335
            Left            =   2640
            TabIndex        =   35
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   13
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridSDR325 
            Height          =   1335
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   13
            FixedCols       =   0
         End
         Begin VB.Label Label19 
            Caption         =   "SDR 17"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   480
            TabIndex        =   40
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "SDR 32.5"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   480
            TabIndex        =   39
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "SDR 26"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3120
            TabIndex        =   38
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "SDR 41"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3120
            TabIndex        =   37
            Top             =   1920
            Width           =   1335
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Polietileno blando"
         ForeColor       =   &H00000080&
         Height          =   3735
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   5295
         Begin MSFlexGridLib.MSFlexGrid gridP25 
            Height          =   1335
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   8
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridP40 
            Height          =   1335
            Left            =   2760
            TabIndex        =   25
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   8
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridP60 
            Height          =   1335
            Left            =   120
            TabIndex        =   26
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   8
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridP80 
            Height          =   1335
            Left            =   2760
            TabIndex        =   27
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   8
            FixedCols       =   0
         End
         Begin VB.Label Label24 
            Caption         =   "Blando 2.5 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label25 
            Caption         =   "Blando 6 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Blando 4 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3000
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Blando 8 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3000
            TabIndex        =   28
            Top             =   1920
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Polietileno Duro"
         ForeColor       =   &H00000080&
         Height          =   3735
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   5175
         Begin MSFlexGridLib.MSFlexGrid gridPd25 
            Height          =   1335
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   11
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridPd40 
            Height          =   1335
            Left            =   2640
            TabIndex        =   16
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   11
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridpD60 
            Height          =   1335
            Left            =   2640
            TabIndex        =   17
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   11
            FixedCols       =   0
         End
         Begin MSFlexGridLib.MSFlexGrid gridpD80 
            Height          =   1335
            Left            =   120
            TabIndex        =   18
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   11
            FixedCols       =   0
         End
         Begin VB.Label Label28 
            Caption         =   "Duro 2.5 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label29 
            Caption         =   "Duro 6 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label31 
            Caption         =   "Duro 4 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2880
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label32 
            Caption         =   "Duro 8 kg/cm2"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2880
            TabIndex        =   19
            Top             =   1920
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Aluminio"
         ForeColor       =   &H00000080&
         Height          =   1935
         Left            =   5640
         TabIndex        =   12
         Top             =   4080
         Width           =   2775
         Begin MSFlexGridLib.MSFlexGrid gridAL 
            Height          =   1335
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   9
            FixedCols       =   0
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Regresar"
         Height          =   495
         Left            =   8520
         TabIndex        =   11
         Top             =   6840
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   98
      Top             =   7815
      Width           =   11400
      _ExtentX        =   20108
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
            TextSave        =   "8/21/2008"
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
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   173
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   11535
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   960
         TabIndex        =   92
         Top             =   240
         Width           =   1935
         Begin VB.TextBox txtY 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbletiqueta1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblunidades1 
            Height          =   255
            Left            =   1320
            TabIndex        =   94
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Diámetros Superior"
         Height          =   1815
         Left            =   120
         TabIndex        =   82
         Top             =   1440
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox Text1 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2520
            TabIndex        =   99
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtDC 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1920
            TabIndex        =   85
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtDCC 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtPC 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   600
            TabIndex        =   83
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Vel= m/s"
            Height          =   255
            Left            =   2760
            TabIndex        =   101
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Diámetro interno"
            Height          =   255
            Left            =   1920
            TabIndex        =   91
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Diámetro Comercial"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Pérdidas efectivas en el tramo"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "mm"
            Height          =   255
            Left            =   3240
            TabIndex        =   88
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "mm"
            Height          =   255
            Left            =   1440
            TabIndex        =   87
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "m"
            Height          =   255
            Left            =   2040
            TabIndex        =   86
            Top             =   1320
            Width           =   375
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1815
         Left            =   3840
         TabIndex        =   72
         Top             =   1440
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox Text2 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2520
            TabIndex        =   100
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtPC1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   600
            TabIndex        =   75
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtDCC1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtDC1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1920
            TabIndex        =   73
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Vel= m/s"
            Height          =   255
            Left            =   2640
            TabIndex        =   102
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "m"
            Height          =   255
            Left            =   1920
            TabIndex        =   81
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "mm"
            Height          =   255
            Left            =   1440
            TabIndex        =   80
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "mm"
            Height          =   255
            Left            =   3240
            TabIndex        =   79
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   "Pérdidas efectivas en el tramo"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label15 
            Caption         =   "Diámetro Comercial"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label16 
            Caption         =   "Diámetro interno"
            Height          =   255
            Left            =   1920
            TabIndex        =   76
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton bcombinacion 
         Caption         =   "Desea combinar diámetros"
         Height          =   495
         Left            =   4920
         MaskColor       =   &H008080FF&
         Picture         =   "hidraulica de prncipal.frx":29F2
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame14 
         Caption         =   "Combinación de Diámetros"
         Height          =   2895
         Left            =   7560
         TabIndex        =   56
         Top             =   360
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txthft1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   61
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtl1 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   60
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txthft 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   59
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txthft2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2040
            TabIndex        =   58
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtl2 
            BackColor       =   &H80000016&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   57
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label33 
            Caption         =   "Pérdidas tramo 1"
            Height          =   255
            Left            =   2040
            TabIndex        =   71
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label34 
            Caption         =   "Longitud tramo 1"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label35 
            Caption         =   "Pérdidas efectivas en todo el tramo"
            Height          =   255
            Left            =   360
            TabIndex        =   69
            Top             =   2160
            Width           =   2535
         End
         Begin VB.Label Label36 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   68
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label37 
            Caption         =   "m"
            Height          =   255
            Left            =   1560
            TabIndex        =   67
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label39 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   66
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label fsfff 
            Caption         =   "Pérdidas tramo 2"
            Height          =   255
            Left            =   2040
            TabIndex        =   65
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label fsf 
            Caption         =   "Longitud tramo 2"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label42 
            Caption         =   "m"
            Height          =   255
            Left            =   3360
            TabIndex        =   63
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label43 
            Caption         =   "m"
            Height          =   255
            Left            =   1560
            TabIndex        =   62
            Top             =   1560
            Width           =   375
         End
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Cálculos en la principal"
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
      Left            =   360
      TabIndex        =   54
      Top             =   360
      Width           =   4815
   End
   Begin VB.Menu zmam 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu ml 
         Caption         =   "Cálculos en laterales"
      End
      Begin VB.Menu mb 
         Caption         =   "Selección de bombas"
      End
   End
   Begin VB.Menu zmpp 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmHprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Double
Dim X As Double
Dim L As Double
Dim c As Double
Dim Y As Double

Dim ysup As Double
Dim ymin As Double
Dim ysup1 As Double
Dim ymin1 As Double

Private Sub bcombinacion_Click()
On Error GoTo mensaje:
hf1 = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ysup ^ -4.872
   
hf2 = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ymin ^ -4.872

l1 = (X - (hf2 * L)) / (hf1 - hf2)
l2 = L - l1
hft1 = hf1 * l1
hft2 = hf2 * l2
hft = hft1 + hft2
txtl1 = Format(l1, "##0.0#")
txtl2 = Format(l2, "##0.0#")
txthft1 = Format(hft1, "##0.0##")
txthft2 = Format(hft2, "##0.0##")
txthft = Format(hft, "##0.0##")
Frame14.Visible = True
Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"
End Sub

Private Sub bevaluar_Click()
On Error GoTo mensaje:
q = Val(txtQ.text)
L = Val(txtL.text)
X = Val(txtX.text)
If q = 0 Then
    MsgBox "Ingrese el valor del caudal del sistema", 64, "Cálculo en la principal"
    txtQ.SetFocus
    Exit Sub
End If
If L = 0 Then
    MsgBox "Ingrese el valor de la longitud del tramo", 64, "Cálculo en la principal"
    txtL.SetFocus
    Exit Sub
End If
If X = 0 Then
    If Option1.Value = True Then
    MsgBox "Ingrese de las pérdidas admisibles", 64, "Cálculo en la principal"
    Else
    MsgBox "Ingrese el valor del díametro", 64, "Cálculo en la principal"
    End If
    txtX.SetFocus
    Exit Sub
End If

Frame7.Visible = False
Frame8.Visible = False
bcombinacion.Visible = False

c = Combo1.ListIndex
If c = -1 Then
    MsgBox "Seleccione adecuadamente el tipo de tubería", 64, "Cálculo en la principal"
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


If Option1.Value = True Then
    Y = (L * 1.131 * 10 ^ 9 * (q / c) ^ 1.852 / X) ^ 0.20525
    
'aluminio**********************************************
    
    If Combo1.ListIndex = 1 Then
        If Y > 249.38 Then
           MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
           Exit Sub
        End If
        If Y < 48.81 Then
            ymin = Val(gridAL.TextMatrix(1, 1))
            ysup = Val(gridAL.TextMatrix(1, 1))
            ymin1 = Val(gridAL.TextMatrix(1, 0))
            ysup1 = Val(gridAL.TextMatrix(1, 0))
        End If

        For j% = 1 To 7
            If (Y > Val(gridAL.TextMatrix(j%, 1))) And (Y < Val(gridAL.TextMatrix(j% + 1, 1))) Then
                ymin = Val(gridAL.TextMatrix(j%, 1))
                ymin1 = Val(gridAL.TextMatrix(j%, 0))
                ysup = Val(gridAL.TextMatrix(j% + 1, 1))
                ysup1 = Val(gridAL.TextMatrix(j% + 1, 0))
            End If
            If Y = Val(gridAL.TextMatrix(j%, 1)) Then
                ymin = Val(gridAL.TextMatrix(j%, 1))
                ymin1 = Val(gridAL.TextMatrix(j%, 0))
                ysup = Val(gridAL.TextMatrix(j%, 1))
                ysup1 = Val(gridAL.TextMatrix(j%, 0))
            End If
        Next j%
        PC = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ysup ^ -4.872 * L
        PC1 = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ymin ^ -4.872 * L
        txtDC = ysup
        txtDCC = ysup1
        txtPC = Format(PC, "###0.0###")
        txtDC1 = ymin
        txtDCC1 = ymin1
        txtPC1 = Format(PC1, "###0.0###")
        Frame7.Visible = True
        Frame8.Visible = True
        bcombinacion.Visible = True
        
    End If
    
    If Combo1.ListIndex = 4 Then
        If Y > 249.38 Then
           MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
           Exit Sub
        End If
        If Y < 48.81 Then
            ymin = Val(gridAL.TextMatrix(1, 1))
            ysup = Val(gridAL.TextMatrix(1, 1))
            ymin1 = Val(gridAL.TextMatrix(1, 0))
            ysup1 = Val(gridAL.TextMatrix(1, 0))
        End If

        For j% = 1 To 7
            If (Y > Val(gridAL.TextMatrix(j%, 1))) And (Y < Val(gridAL.TextMatrix(j% + 1, 1))) Then
                ymin = Val(gridAL.TextMatrix(j%, 1))
                ymin1 = Val(gridAL.TextMatrix(j%, 0))
                ysup = Val(gridAL.TextMatrix(j% + 1, 1))
                ysup1 = Val(gridAL.TextMatrix(j% + 1, 0))
            End If
            If Y = Val(gridAL.TextMatrix(j%, 1)) Then
                ysup = Val(gridAL.TextMatrix(j%, 1))
                ysup1 = Val(gridAL.TextMatrix(j%, 0))
                ymin = Val(gridAL.TextMatrix(j%, 1))
                ymin1 = Val(gridAL.TextMatrix(j%, 0))
            End If
        Next j%
        PC = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ysup ^ -4.872 * L
        PC1 = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ymin ^ -4.872 * L
        txtDC = ysup
        txtDCC = ysup1
        txtPC = Format(PC, "###0.0###")
        txtDC1 = ymin
        txtDCC1 = ymin1
        txtPC1 = Format(PC1, "###0.0###")
        Frame7.Visible = True
        Frame8.Visible = True
        bcombinacion.Visible = True
    End If
    
'polietileno*******************
    
    If Combo1.ListIndex = 7 Then
            
            'blando 2.5
        If Combo3.ListIndex = 0 Then
            If Y > 45 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 9.8 Then
            ymin = Val(gridP25.TextMatrix(1, 1))
            ysup = Val(gridP25.TextMatrix(1, 1))
            ymin1 = Val(gridP25.TextMatrix(1, 0))
            ysup1 = Val(gridP25.TextMatrix(1, 0))
        End If

            For j% = 1 To 6
                If (Y > Val(gridP25.TextMatrix(j%, 1))) And (Y < Val(gridP25.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridP25.TextMatrix(j%, 1))
                    ymin1 = Val(gridP25.TextMatrix(j%, 0))
                    ysup = Val(gridP25.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridP25.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridP25.TextMatrix(j%, 1)) Then
                    ysup = Val(gridP25.TextMatrix(j%, 1))
                    ysup1 = Val(gridP25.TextMatrix(j%, 0))
                    ymin = Val(gridP25.TextMatrix(j%, 1))
                    ymin1 = Val(gridP25.TextMatrix(j%, 0))
                End If
            Next j%
        End If
        
        ' blando 4.0
        If Combo3.ListIndex = 1 Then
            If Y > 42.3 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 9.6 Then
            ymin = Val(gridP40.TextMatrix(1, 1))
            ysup = Val(gridP40.TextMatrix(1, 1))
            ymin1 = Val(gridP40.TextMatrix(1, 0))
            ysup1 = Val(gridP40.TextMatrix(1, 0))
        End If

            For j% = 1 To 6
                If (Y > Val(gridP40.TextMatrix(j%, 1))) And (Y < Val(gridP40.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridP40.TextMatrix(j%, 1))
                    ymin1 = Val(gridP40.TextMatrix(j%, 0))
                    ysup = Val(gridP40.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridP40.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridP40.TextMatrix(j%, 1)) Then
                    ysup = Val(gridP40.TextMatrix(j%, 1))
                    ysup1 = Val(gridP40.TextMatrix(j%, 0))
                    ymin = Val(gridP40.TextMatrix(j%, 1))
                    ymin1 = Val(gridP40.TextMatrix(j%, 0))
                End If
            Next j%
        End If
           
        ' blando 6.0
        If Combo3.ListIndex = 2 Then
            If Y > 38.4 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 9.2 Then
            ymin = Val(gridP60.TextMatrix(1, 1))
            ysup = Val(gridP60.TextMatrix(1, 1))
            ymin1 = Val(gridP60.TextMatrix(1, 0))
            ysup1 = Val(gridP60.TextMatrix(1, 0))
            End If
            For j% = 1 To 6
                If (Y > Val(gridP60.TextMatrix(j%, 1))) And (Y < Val(gridP60.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridP60.TextMatrix(j%, 1))
                    ymin1 = Val(gridP60.TextMatrix(j%, 0))
                    ysup = Val(gridP60.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridP60.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridP60.TextMatrix(j%, 1)) Then
                    ysup = Val(gridP60.TextMatrix(j%, 1))
                    ysup1 = Val(gridP60.TextMatrix(j%, 0))
                    ymin = Val(gridP60.TextMatrix(j%, 1))
                    ymin1 = Val(gridP60.TextMatrix(j%, 0))
                End If
            Next j%
        End If
                
         ' blando 8.0
       If Combo3.ListIndex = 3 Then
            If Y > 33 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 7.9 Then
            ymin = Val(gridP80.TextMatrix(1, 1))
            ysup = Val(gridP80.TextMatrix(1, 1))
            ymin1 = Val(gridP80.TextMatrix(1, 0))
            ysup1 = Val(gridP80.TextMatrix(1, 0))
            End If
            For j% = 1 To 6
                If (Y > Val(gridP80.TextMatrix(j%, 1))) And (Y < Val(gridP80.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridP80.TextMatrix(j%, 1))
                    ymin1 = Val(gridP80.TextMatrix(j%, 0))
                    ysup = Val(gridP80.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridP80.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridP80.TextMatrix(j%, 1)) Then
                    ysup = Val(gridP80.TextMatrix(j%, 1))
                    ysup1 = Val(gridP80.TextMatrix(j%, 0))
                    ymin = Val(gridP80.TextMatrix(j%, 1))
                    ymin1 = Val(gridP80.TextMatrix(j%, 0))
                End If
            Next j%
        End If
        
              'duro 2.5
        If Combo3.ListIndex = 4 Then
            If Y > 104.4 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 46.7 Then
            ymin = Val(gridPd25.TextMatrix(1, 1))
            ysup = Val(gridPd25.TextMatrix(1, 1))
            ymin1 = Val(gridPd25.TextMatrix(1, 0))
            ysup1 = Val(gridPd25.TextMatrix(1, 0))
            End If

            For j% = 1 To 4
                If (Y > Val(gridPd25.TextMatrix(j%, 1))) And (Y < Val(gridPd25.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridPd25.TextMatrix(j%, 1))
                    ymin1 = Val(gridPd25.TextMatrix(j%, 0))
                    ysup = Val(gridPd25.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridPd25.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridPd25.TextMatrix(j%, 1)) Then
                    ysup = Val(gridPd25.TextMatrix(j%, 1))
                    ysup1 = Val(gridPd25.TextMatrix(j%, 0))
                    ymin = Val(gridPd25.TextMatrix(j%, 1))
                    ymin1 = Val(gridPd25.TextMatrix(j%, 0))
                End If
            Next j%
        End If
        
        ' duro 4.0
        If Combo3.ListIndex = 5 Then
            If Y > 101.3 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 28.7 Then
            ymin = Val(gridPd40.TextMatrix(1, 1))
            ysup = Val(gridPd40.TextMatrix(1, 1))
            ymin1 = Val(gridPd40.TextMatrix(1, 0))
            ysup1 = Val(gridPd40.TextMatrix(1, 0))
            End If

            For j% = 1 To 6
                If (Y > Val(gridPd40.TextMatrix(j%, 1))) And (Y < Val(gridPd40.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridPd40.TextMatrix(j%, 1))
                    ymin1 = Val(gridPd40.TextMatrix(j%, 0))
                    ysup = Val(gridPd40.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridPd40.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridPd40.TextMatrix(j%, 1)) Then
                    ysup = Val(gridPd40.TextMatrix(j%, 1))
                    ysup1 = Val(gridPd40.TextMatrix(j%, 0))
                    ymin = Val(gridPd40.TextMatrix(j%, 1))
                    ymin1 = Val(gridPd40.TextMatrix(j%, 0))
                End If
            Next j%
        End If
           
    
        ' durp 6.0
        If Combo3.ListIndex = 6 Then
            If Y > 96.4 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 21.7 Then
            ymin = Val(gridpD60.TextMatrix(1, 1))
            ysup = Val(gridpD60.TextMatrix(1, 1))
            ymin1 = Val(gridpD60.TextMatrix(1, 0))
            ysup1 = Val(gridpD60.TextMatrix(1, 0))
            End If

            For j% = 1 To 7
                If (Y > Val(gridpD60.TextMatrix(j%, 1))) And (Y < Val(gridpD60.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridpD60.TextMatrix(j%, 1))
                    ymin1 = Val(gridpD60.TextMatrix(j%, 0))
                    ysup = Val(gridpD60.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridpD60.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridpD60.TextMatrix(j%, 1)) Then
                    ysup = Val(gridpD60.TextMatrix(j%, 1))
                    ysup1 = Val(gridpD60.TextMatrix(j%, 0))
                    ymin = Val(gridpD60.TextMatrix(j%, 1))
                    ymin1 = Val(gridpD60.TextMatrix(j%, 0))
                End If
            Next j%
        End If
                
         ' duro 8.0
       If Combo3.ListIndex = 7 Then
            If Y > 93.4 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 16.7 Then
            ymin = Val(gridpD80.TextMatrix(1, 1))
            ysup = Val(gridpD80.TextMatrix(1, 1))
            ymin1 = Val(gridpD80.TextMatrix(1, 0))
            ysup1 = Val(gridpD80.TextMatrix(1, 0))
            End If

            For j% = 1 To 8
                If (Y > Val(gridpD80.TextMatrix(j%, 1))) And (Y < Val(gridpD80.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridpD80.TextMatrix(j%, 1))
                    ymin1 = Val(gridpD80.TextMatrix(j%, 0))
                    ysup = Val(gridpD80.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridpD80.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridpD80.TextMatrix(j%, 1)) Then
                    ysup = Val(gridpD80.TextMatrix(j%, 1))
                    ysup1 = Val(gridpD80.TextMatrix(j%, 0))
                    ymin = Val(gridpD80.TextMatrix(j%, 1))
                    ymin1 = Val(gridpD80.TextMatrix(j%, 0))
                End If
            Next j%
        End If
      
        PC = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ysup ^ -4.872 * L
        PC1 = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ymin ^ -4.872 * L
        txtDC = ysup
        txtDCC = ysup1
        txtPC = Format(PC, "###0.0###")
        txtDC1 = ymin
        txtDCC1 = ymin1
        txtPC1 = Format(PC1, "###0.0###")
        Frame7.Visible = True
        Frame8.Visible = True
        bcombinacion.Visible = True
    End If
    '---
'**************P.V.C.*******************
    
    If Combo1.ListIndex = 6 Then
            
            'SDR 17
        If Combo4.ListIndex = 0 Then
            If Y > 285.8 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 23.53 Then
            ymin = Val(gridSDR17.TextMatrix(1, 1))
            ysup = Val(gridSDR17.TextMatrix(1, 1))
            ymin1 = Val(gridSDR17.TextMatrix(1, 0))
            ysup1 = Val(gridSDR17.TextMatrix(1, 0))
            End If

            For j% = 1 To 9
                If (Y > Val(gridSDR17.TextMatrix(j%, 1))) And (Y < Val(gridSDR17.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridSDR17.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR17.TextMatrix(j%, 0))
                    ysup = Val(gridSDR17.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridSDR17.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridSDR17.TextMatrix(j%, 1)) Then
                    ysup = Val(gridSDR17.TextMatrix(j%, 1))
                    ysup1 = Val(gridSDR17.TextMatrix(j%, 0))
                    ymin = Val(gridSDR17.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR17.TextMatrix(j%, 0))
                End If
            Next j%
        End If
        
        ' SDR 26
        If Combo4.ListIndex = 1 Then
            If Y > 298.95 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 30.36 Then
            ymin = Val(gridSDR26.TextMatrix(1, 1))
            ysup = Val(gridSDR26.TextMatrix(1, 1))
            ymin1 = Val(gridSDR26.TextMatrix(1, 0))
            ysup1 = Val(gridSDR26.TextMatrix(1, 0))
            End If
    
            For j% = 1 To 10
                If (Y > Val(gridSDR26.TextMatrix(j%, 1))) And (Y < Val(gridSDR26.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridSDR26.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR26.TextMatrix(j%, 0))
                    ysup = Val(gridSDR26.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridSDR26.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridSDR26.TextMatrix(j%, 1)) Then
                    ysup = Val(gridSDR26.TextMatrix(j%, 1))
                    ysup1 = Val(gridSDR26.TextMatrix(j%, 0))
                    ymin = Val(gridSDR26.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR26.TextMatrix(j%, 0))
                End If
            Next j%
        End If
           
    
        ' SDR 32.5
        If Combo4.ListIndex = 2 Then
            If Y > 303.93 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 39 Then
            ymin = Val(gridSDR325.TextMatrix(1, 1))
            ysup = Val(gridSDR325.TextMatrix(1, 1))
            ymin1 = Val(gridSDR325.TextMatrix(1, 0))
            ysup1 = Val(gridSDR325.TextMatrix(1, 0))
            End If
            For j% = 1 To 9
                If (Y > Val(gridSDR325.TextMatrix(j%, 1))) And (Y < Val(gridSDR325.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridSDR325.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR325.TextMatrix(j%, 0))
                    ysup = Val(gridSDR325.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridSDR325.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridSDR325.TextMatrix(j%, 1)) Then
                    ysup = Val(gridSDR325.TextMatrix(j%, 1))
                    ysup1 = Val(gridSDR325.TextMatrix(j%, 0))
                    ymin = Val(gridSDR325.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR325.TextMatrix(j%, 0))
                End If
            Next j%
        End If
                
         ' SDR 41
       If Combo4.ListIndex = 3 Then
            If Y > 369.7 Then
               MsgBox "El diámetro calculado excede a los diámetros comerciales", 16, "Mensaje de error"
               Exit Sub
            End If
            If Y < 39.8 Then
            ymin = Val(gridSDR41.TextMatrix(1, 1))
            ysup = Val(gridSDR41.TextMatrix(1, 1))
            ymin1 = Val(gridSDR41.TextMatrix(1, 0))
            ysup1 = Val(gridSDR41.TextMatrix(1, 0))
            End If

            For j% = 1 To 10
                If (Y > Val(gridSDR41.TextMatrix(j%, 1))) And (Y < Val(gridSDR41.TextMatrix(j% + 1, 1))) Then
                    ymin = Val(gridSDR41.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR41.TextMatrix(j%, 0))
                    ysup = Val(gridSDR41.TextMatrix(j% + 1, 1))
                    ysup1 = Val(gridSDR41.TextMatrix(j% + 1, 0))
                End If
                If Y = Val(gridSDR41.TextMatrix(j%, 1)) Then
                    ysup = Val(gridSDR41.TextMatrix(j%, 1))
                    ysup1 = Val(gridSDR41.TextMatrix(j%, 0))
                    ymin = Val(gridSDR41.TextMatrix(j%, 1))
                    ymin1 = Val(gridSDR41.TextMatrix(j%, 0))
                End If
            Next j%
        End If
    'sdr 13.5
    If Combo4.ListIndex = 4 Then
            ymin = 18.2
            ysup = 18.2
            ymin1 = 12
            ysup1 = 12
    End If
    
        PC = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ysup ^ -4.872 * L
        PC1 = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * ymin ^ -4.872 * L
        txtDC = ysup
        txtDCC = ysup1
        txtPC = Format(PC, "###0.0###")
        txtDC1 = ymin
        txtDCC1 = ymin1
        txtPC1 = Format(PC1, "###0.0###")
        Frame7.Visible = True
        Frame8.Visible = True
        bcombinacion.Visible = True
    End If
v1 = q / (3.142 * 3600 * (ysup / 2000) ^ 2)
v2 = q / (3.142 * 3600 * (ymin / 2000) ^ 2)
Text1 = v1
Text2 = v2

    '---

Else
    Y = 1.131 * 10 ^ 9 * (q / c) ^ 1.852 * X ^ -4.872 * L
End If
txtqsis = Format(q, "###0.0###")
txtY = Format(Y, "###0.0###")
Txtf = Format(f, "###0.0###")
txtM = Format(m, "###0.0###")
Frame3.Visible = True
StatusBar1.Panels(1).text = ""

Exit Sub
mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"

End Sub


Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
txtQ.text = ""
txtX.text = ""
txtL.text = ""
txtl1.text = ""
txtl2.text = ""
txthft1.text = ""
txthft2.text = ""
txthft.text = ""
txtY.text = ""
txtDCC.text = ""
txtDC.text = ""
txtPC.text = ""
txtDCC1.text = ""
txtDC1.text = ""
txtPC1.text = ""
Frame1.Visible = False
Frame3.Visible = False
Option1.ForeColor = &H80000012
Option2.ForeColor = &H80000012
Option1.Value = False
Option2.Value = False
Combo1.text = ""
Combo3.text = ""
Combo4.text = ""
lbl1.Visible = False
lbl2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
StatusBar1.Panels(1).text = "Seleccione la opción de cálculo que desea realizar"
End Sub

Private Sub Combo1_Click()
THC = Combo1.ListIndex
Select Case THC
    Case 0
        lbl1.Visible = False
        lbl2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    Case 1
        lbl1.Visible = False
        lbl2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    Case 2
        lbl1.Visible = False
        lbl2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    Case 3
        lbl1.Visible = False
        lbl2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    Case 4
        lbl1.Visible = False
        lbl2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    Case 5
        lbl1.Visible = False
        lbl2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    Case 6
        lbl1.Visible = False
        lbl2.Visible = True
        Combo3.Visible = False
        Combo4.Visible = True
    Case 7
        lbl1.Visible = True
        lbl2.Visible = False
        Combo3.Visible = True
        Combo4.Visible = False
End Select

End Sub


Private Sub mb_Click()
frmbomba.Show
End Sub

Private Sub ml_Click()
FrmHLaterales.Show
End Sub

Private Sub Option1_Click()
lbletiqueta.Caption = "Pérdidas admisibles"
lbletiqueta1.Caption = "Diámetro"
lblunidades.Caption = "m"
lblunidades1.Caption = "mm"
Frame1.Visible = True
Option1.ForeColor = &HC0&
Option2.ForeColor = &H80000012
txtQ.SetFocus

StatusBar1.Panels(1).text = "Digite los datos básicos y oprima el botón de Evaluar "
End Sub

Private Sub Option2_Click()
lbletiqueta.Caption = "Diámetro"
lbletiqueta1.Caption = "Pérdidas admisibles"
lblunidades.Caption = "mm"
lblunidades1.Caption = "m"
Frame1.Visible = True
Option2.ForeColor = &HC0&
Option1.ForeColor = &H80000012
txtQ.SetFocus
StatusBar1.Panels(1).text = "Digite los datos básicos y oprima el botón de Evaluar "
End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
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
With Combo3
    .AddItem "Blando de 2.5 kg/cm2"
    .AddItem "Blando de 4.0 kg/cm2"
    .AddItem "Blando de 6.0 kg/cm2"
    .AddItem "Blando de 8.0 kg/cm2"
    .AddItem "Duro de 2.5 kg/cm2"
    .AddItem "Duro de 4.0 kg/cm2"
    .AddItem "Duro de 6.0 kg/cm2"
    .AddItem "Duro de 8.0 kg/cm2"
End With
With Combo4
    .AddItem "SDR 17"
    .AddItem "SDR 26"
    .AddItem "SDR 32.5"
    .AddItem "SDR 41"
    .AddItem "SDR 13.5"
End With
StatusBar1.Panels(1).text = "Seleccione la opción de cálculo que desea realizar"


With gridAL
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "50"
    .TextMatrix(2, 0) = "75"
    .TextMatrix(3, 0) = "100"
    .TextMatrix(4, 0) = "125"
    .TextMatrix(5, 0) = "150"
    .TextMatrix(6, 0) = "175"
    .TextMatrix(7, 0) = "200"
    .TextMatrix(8, 0) = "250"
    .TextMatrix(1, 1) = "48.81"
    .TextMatrix(2, 1) = "74.01"
    .TextMatrix(3, 1) = "99.21"
    .TextMatrix(4, 1) = "124.36"
    .TextMatrix(5, 1) = "149.45"
    .TextMatrix(6, 1) = "174.55"
    .TextMatrix(7, 1) = "199.54"
    .TextMatrix(8, 1) = "249.38"
End With
With gridP25
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "12"
    .TextMatrix(2, 0) = "16"
    .TextMatrix(3, 0) = "20"
    .TextMatrix(4, 0) = "25"
    .TextMatrix(5, 0) = "32"
    .TextMatrix(6, 0) = "40"
    .TextMatrix(7, 0) = "50"
    .TextMatrix(1, 1) = "9.8"
    .TextMatrix(2, 1) = "13.1"
    .TextMatrix(3, 1) = "16.9"
    .TextMatrix(4, 1) = "21.7"
    .TextMatrix(5, 1) = "28.7"
    .TextMatrix(6, 1) = "36"
    .TextMatrix(7, 1) = "45"
End With

With gridP40
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "12"
    .TextMatrix(2, 0) = "16"
    .TextMatrix(3, 0) = "20"
    .TextMatrix(4, 0) = "25"
    .TextMatrix(5, 0) = "32"
    .TextMatrix(6, 0) = "40"
    .TextMatrix(7, 0) = "50"
    .TextMatrix(1, 1) = "9.6"
    .TextMatrix(2, 1) = "12.7"
    .TextMatrix(3, 1) = "16.5"
    .TextMatrix(4, 1) = "21.1"
    .TextMatrix(5, 1) = "27"
    .TextMatrix(6, 1) = "33.8"
    .TextMatrix(7, 1) = "42.3"
End With

With gridP60
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "12"
    .TextMatrix(2, 0) = "16"
    .TextMatrix(3, 0) = "20"
    .TextMatrix(4, 0) = "25"
    .TextMatrix(5, 0) = "32"
    .TextMatrix(6, 0) = "40"
    .TextMatrix(7, 0) = "50"
    .TextMatrix(1, 1) = "9.2"
    .TextMatrix(2, 1) = "12.3"
    .TextMatrix(3, 1) = "15.1"
    .TextMatrix(4, 1) = "19.2"
    .TextMatrix(5, 1) = "24.5"
    .TextMatrix(6, 1) = "30.8"
    .TextMatrix(7, 1) = "38.4"
End With



With gridP80
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "12"
    .TextMatrix(2, 0) = "16"
    .TextMatrix(3, 0) = "20"
    .TextMatrix(4, 0) = "25"
    .TextMatrix(5, 0) = "32"
    .TextMatrix(6, 0) = "40"
    .TextMatrix(7, 0) = "50"
    .TextMatrix(1, 1) = "7.9"
    .TextMatrix(2, 1) = "10.4"
    .TextMatrix(3, 1) = "13"
    .TextMatrix(4, 1) = "16.3"
    .TextMatrix(5, 1) = "20.9"
    .TextMatrix(6, 1) = "26.3"
    .TextMatrix(7, 1) = "33"
End With
With gridPd25
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "50"
    .TextMatrix(2, 0) = "63"
    .TextMatrix(3, 0) = "75"
    .TextMatrix(4, 0) = "90"
    .TextMatrix(5, 0) = "110"
    .TextMatrix(1, 1) = "46.7"
    .TextMatrix(2, 1) = "59.7"
    .TextMatrix(3, 1) = "71.1"
    .TextMatrix(4, 1) = "85.5"
    .TextMatrix(5, 1) = "104.4"
End With

With gridPd40
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "32"
    .TextMatrix(2, 0) = "40"
    .TextMatrix(3, 0) = "50"
    .TextMatrix(4, 0) = "63"
    .TextMatrix(5, 0) = "75"
    .TextMatrix(6, 0) = "90"
    .TextMatrix(7, 0) = "110"
    .TextMatrix(1, 1) = "28.7"
    .TextMatrix(2, 1) = "36.7"
    .TextMatrix(3, 1) = "45.8"
    .TextMatrix(4, 1) = "58"
    .TextMatrix(5, 1) = "69"
    .TextMatrix(6, 1) = "82.8"
    .TextMatrix(7, 1) = "101.3"
End With

With gridpD60
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "25"
    .TextMatrix(2, 0) = "32"
    .TextMatrix(3, 0) = "40"
    .TextMatrix(4, 0) = "50"
    .TextMatrix(5, 0) = "63"
    .TextMatrix(6, 0) = "75"
    .TextMatrix(7, 0) = "90"
    .TextMatrix(8, 0) = "110"
    .TextMatrix(1, 1) = "21.7"
    .TextMatrix(2, 1) = "28.1"
    .TextMatrix(3, 1) = "35"
    .TextMatrix(4, 1) = "43.8"
    .TextMatrix(5, 1) = "55.2"
    .TextMatrix(6, 1) = "65.7"
    .TextMatrix(7, 1) = "78.9"
    .TextMatrix(8, 1) = "96.4"
End With



With gridpD80
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "20"
    .TextMatrix(2, 0) = "25"
    .TextMatrix(3, 0) = "32"
    .TextMatrix(4, 0) = "40"
    .TextMatrix(5, 0) = "50"
    .TextMatrix(6, 0) = "63"
    .TextMatrix(7, 0) = "75"
    .TextMatrix(8, 0) = "90"
    .TextMatrix(9, 0) = "110"
    .TextMatrix(1, 1) = "16.7"
    .TextMatrix(2, 1) = "21.7"
    .TextMatrix(3, 1) = "27"
    .TextMatrix(4, 1) = "33.8"
    .TextMatrix(5, 1) = "42.4"
    .TextMatrix(6, 1) = "53.3"
    .TextMatrix(7, 1) = "63.7"
    .TextMatrix(8, 1) = "76.4"
    .TextMatrix(9, 1) = "93.4"
    
End With

With gridSDR17
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "18"
    .TextMatrix(2, 0) = "25"
    .TextMatrix(3, 0) = "31"
    .TextMatrix(4, 0) = "38"
    .TextMatrix(5, 0) = "50"
    .TextMatrix(6, 0) = "62"
    .TextMatrix(7, 0) = "75"
    .TextMatrix(8, 0) = "100"
    .TextMatrix(9, 0) = "150"
    .TextMatrix(10, 0) = "200"
    .TextMatrix(11, 0) = "250"
    .TextMatrix(12, 0) = "300"
    .TextMatrix(1, 1) = "23.53"
    .TextMatrix(2, 1) = "29.48"
    .TextMatrix(3, 1) = "37.18"
    .TextMatrix(4, 1) = "42.58"
    .TextMatrix(5, 1) = "53.21"
    .TextMatrix(6, 1) = "54.45"
    .TextMatrix(7, 1) = "78.44"
    .TextMatrix(8, 1) = "100.84"
    .TextMatrix(9, 1) = "148.46"
    .TextMatrix(10, 1) = "193.28"
    .TextMatrix(11, 1) = "240.90"
    .TextMatrix(12, 1) = "285.80"
End With


With gridSDR26
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "25"
    .TextMatrix(2, 0) = "31"
    .TextMatrix(3, 0) = "38"
    .TextMatrix(4, 0) = "50"
    .TextMatrix(5, 0) = "62"
    .TextMatrix(6, 0) = "75"
    .TextMatrix(7, 0) = "100"
    .TextMatrix(8, 0) = "150"
    .TextMatrix(9, 0) = "200"
    .TextMatrix(10, 0) = "250"
    .TextMatrix(11, 0) = "300"
    .TextMatrix(1, 1) = "30.36"
    .TextMatrix(2, 1) = "38.9"
    .TextMatrix(3, 1) = "44.56"
    .TextMatrix(4, 1) = "55.71"
    .TextMatrix(5, 1) = "67.45"
    .TextMatrix(6, 1) = "82.04"
    .TextMatrix(7, 1) = "105.52"
    .TextMatrix(8, 1) = "155.32"
    .TextMatrix(9, 1) = "202.22"
    .TextMatrix(10, 1) = "252.07"
    .TextMatrix(11, 1) = "298.95"
End With

With gridSDR325
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "31"
    .TextMatrix(2, 0) = "38"
    .TextMatrix(3, 0) = "50"
    .TextMatrix(4, 0) = "62"
    .TextMatrix(5, 0) = "75"
    .TextMatrix(6, 0) = "100"
    .TextMatrix(7, 0) = "150"
    .TextMatrix(8, 0) = "200"
    .TextMatrix(9, 0) = "250"
    .TextMatrix(10, 0) = "300"
    .TextMatrix(1, 1) = "39"
    .TextMatrix(2, 1) = "45.22"
    .TextMatrix(3, 1) = "56.63"
    .TextMatrix(4, 1) = "68.55"
    .TextMatrix(5, 1) = "83.42"
    .TextMatrix(6, 1) = "107.28"
    .TextMatrix(7, 1) = "157.92"
    .TextMatrix(8, 1) = "205.62"
    .TextMatrix(9, 1) = "256.23"
    .TextMatrix(10, 1) = "303.93"
End With
With gridSDR41
    .TextMatrix(0, 0) = "Dia nom"
    .TextMatrix(0, 1) = "Dia int"
    .TextMatrix(1, 0) = "31"
    .TextMatrix(2, 0) = "38"
    .TextMatrix(3, 0) = "50"
    .TextMatrix(4, 0) = "62"
    .TextMatrix(5, 0) = "75"
    .TextMatrix(6, 0) = "100"
    .TextMatrix(7, 0) = "150"
    .TextMatrix(8, 0) = "200"
    .TextMatrix(9, 0) = "250"
    .TextMatrix(10, 0) = "300"
    .TextMatrix(1, 1) = "39.8"
    .TextMatrix(2, 1) = "45.9"
    .TextMatrix(3, 1) = "57.38"
    .TextMatrix(4, 1) = "69.46"
    .TextMatrix(5, 1) = "84.58"
    .TextMatrix(6, 1) = "108.72"
    .TextMatrix(7, 1) = "160.08"
    .TextMatrix(8, 1) = "208.42"
    .TextMatrix(9, 1) = "259.75"
    .TextMatrix(10, 1) = "308.05"
    .TextMatrix(11, 1) = "369.7"
    
End With
End Sub

Private Sub zmpp_Click()
Unload Me
frmGeneral.Show
End Sub
