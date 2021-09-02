VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmetodocostos 
   Caption         =   "Comparación de costos energía vrs conducción"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmmetododecomparaciondecostos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   2520
      TabIndex        =   1
      Top             =   6960
      Width           =   6495
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   480
         Picture         =   "frmmetododecomparaciondecostos.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2400
         Picture         =   "frmmetododecomparaciondecostos.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   4320
         Picture         =   "frmmetododecomparaciondecostos.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   8220
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
            TextSave        =   "5/7/2008"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Método de comparación de Costos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Consumo de la bomba"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame9 
      Height          =   5895
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   11535
      Begin VB.Frame Frame1 
         Caption         =   "Resultados"
         Height          =   2655
         Left            =   720
         TabIndex        =   42
         Top             =   3120
         Width           =   9735
         Begin VB.Frame Frame6 
            Caption         =   "Costos de la Red A"
            ForeColor       =   &H00000080&
            Height          =   1695
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   4455
            Begin VB.TextBox ct 
               BackColor       =   &H80000004&
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   2280
               TabIndex        =   53
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox cat 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   2280
               TabIndex        =   52
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox cae 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   2280
               TabIndex        =   51
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label28 
               Caption         =   "$"
               Height          =   375
               Left            =   3600
               TabIndex        =   86
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label27 
               Caption         =   "$"
               Height          =   375
               Left            =   3600
               TabIndex        =   85
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label Label26 
               Caption         =   "$"
               Height          =   375
               Left            =   3600
               TabIndex        =   57
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "Costo anual por enegía"
               Height          =   375
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label2 
               Caption         =   "Costos totales anuales"
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   120
               TabIndex        =   55
               Top             =   1200
               Width           =   2175
            End
            Begin VB.Label Label1 
               Caption         =   "Costo anual por proyecto"
               Height          =   375
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   2055
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Costos de la Red B"
            ForeColor       =   &H00000080&
            Height          =   1695
            Left            =   4920
            TabIndex        =   43
            Top             =   600
            Width           =   4455
            Begin VB.TextBox caeb 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   2280
               TabIndex        =   46
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox catb 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   2280
               TabIndex        =   45
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox ctb 
               BackColor       =   &H80000004&
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   2280
               TabIndex        =   44
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label45 
               Caption         =   "$"
               Height          =   375
               Left            =   3600
               TabIndex        =   89
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label41 
               Caption         =   "$"
               Height          =   375
               Left            =   3600
               TabIndex        =   88
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label37 
               Caption         =   "$"
               Height          =   375
               Left            =   3600
               TabIndex        =   87
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Costo anual por proyecto"
               Height          =   375
               Left            =   120
               TabIndex        =   49
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label Label13 
               Caption         =   "Costos totales anuales"
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   120
               TabIndex        =   48
               Top             =   1200
               Width           =   2175
            End
            Begin VB.Label Label14 
               Caption         =   "Costo anual por enegía"
               Height          =   375
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.TextBox txtfrc 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   3600
            TabIndex        =   58
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "FRC"
            Height          =   375
            Left            =   1440
            TabIndex        =   61
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "Mas económico"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1920
            TabIndex        =   60
            Top             =   2280
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label32 
            Caption         =   "Mas económico"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   6720
            TabIndex        =   59
            Top             =   2280
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Finacieros Costo de Red"
         Height          =   1335
         Left            =   8160
         TabIndex        =   35
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txttasint 
            Height          =   285
            Left            =   1440
            TabIndex        =   37
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtviduti 
            Height          =   285
            Left            =   1440
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Tasa de interés"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Vida útil"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "%"
            Height          =   375
            Left            =   2760
            TabIndex        =   39
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "años"
            Height          =   375
            Index           =   0
            Left            =   2760
            TabIndex        =   38
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos de la Energía"
         Height          =   2295
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtcoscom 
            Height          =   285
            Left            =   1560
            TabIndex        =   26
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtnh 
            Height          =   285
            Left            =   1560
            TabIndex        =   25
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1560
            TabIndex        =   24
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox re 
            Height          =   285
            Left            =   1560
            TabIndex        =   23
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Combustible"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Costo del combustible"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Numero de horas por año"
            Height          =   495
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label24 
            Caption         =   "$/lts"
            Height          =   375
            Left            =   2880
            TabIndex        =   31
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "$/kw"
            Height          =   375
            Left            =   2880
            TabIndex        =   30
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label30 
            Caption         =   "hp-hora/ kw"
            Height          =   375
            Left            =   2880
            TabIndex        =   29
            Top             =   1800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Rendimiento del equipo"
            Height          =   495
            Left            =   240
            TabIndex        =   28
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "hp-hora/ lts"
            Height          =   375
            Left            =   2880
            TabIndex        =   27
            Top             =   1800
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Opción de costos A"
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3615
         Begin VB.TextBox txtcdt 
            Height          =   285
            Left            =   1680
            TabIndex        =   17
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtcostub 
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "hp"
            Height          =   375
            Left            =   3000
            TabIndex        =   21
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Potencia requerida"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label9 
            Caption         =   "Costo proyecto"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label10 
            Caption         =   "$"
            Height          =   375
            Left            =   3000
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Opción de costos B"
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
         Begin VB.TextBox txtcdtB 
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtcostubB 
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Costo proyecto"
            Height          =   375
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label33 
            Caption         =   "Potencia requerida"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label34 
            Caption         =   "hp"
            Height          =   375
            Left            =   3000
            TabIndex        =   13
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label36 
            Caption         =   "$"
            Height          =   375
            Left            =   3000
            TabIndex        =   12
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton BcalGOTERO 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   8520
         Picture         =   "frmmetododecomparaciondecostos.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      Height          =   5895
      Left            =   120
      TabIndex        =   62
      Top             =   960
      Visible         =   0   'False
      Width           =   11535
      Begin VB.Frame Frame12 
         Caption         =   "Consumo y costos"
         ForeColor       =   &H00000080&
         Height          =   1695
         Left            =   4080
         TabIndex        =   75
         Top             =   3720
         Width           =   4455
         Begin VB.TextBox consumo 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2280
            TabIndex        =   77
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox costo 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2280
            TabIndex        =   76
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "Costo total"
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label40 
            Caption         =   "Consumo de energía"
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label39 
            Caption         =   "Lts"
            Height          =   375
            Left            =   3600
            TabIndex        =   79
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label38 
            Caption         =   " $"
            Height          =   375
            Left            =   3600
            TabIndex        =   78
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Datos de la Energía"
         Height          =   2775
         Left            =   600
         TabIndex        =   63
         Top             =   480
         Width           =   5295
         Begin VB.TextBox pot2 
            Height          =   285
            Left            =   2760
            TabIndex        =   70
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox cc2 
            Height          =   285
            Left            =   2760
            TabIndex        =   65
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox nh2 
            Height          =   285
            Left            =   2760
            TabIndex        =   67
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2760
            TabIndex        =   64
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox re2 
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   2760
            TabIndex        =   69
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label59 
            Caption         =   "$/kw"
            Height          =   375
            Left            =   4080
            TabIndex        =   84
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label44 
            Caption         =   "Potencia de la bomba"
            Height          =   495
            Left            =   240
            TabIndex        =   83
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label Label43 
            Caption         =   "hp"
            Height          =   375
            Left            =   4080
            TabIndex        =   82
            Top             =   2280
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "Combustible"
            Height          =   375
            Left            =   240
            TabIndex        =   74
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label56 
            Caption         =   "Costo del combustible"
            Height          =   375
            Left            =   240
            TabIndex        =   73
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label57 
            Caption         =   "Numero de horas de trabajo"
            Height          =   495
            Left            =   240
            TabIndex        =   71
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label61 
            Caption         =   "Rendimiento del equipo"
            Height          =   495
            Left            =   240
            TabIndex        =   68
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label62 
            Caption         =   "hp-hora/ lts"
            Height          =   375
            Left            =   4080
            TabIndex        =   66
            Top             =   1800
            Width           =   975
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   8520
         Picture         =   "frmmetododecomparaciondecostos.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Comparación de costos                                              Energía vrs Costo del proyecto"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Menu aaa 
      Caption         =   "Hidráulica de tuberías"
      Begin VB.Menu bbbb 
         Caption         =   "Cálculos de la principal"
      End
      Begin VB.Menu ccc 
         Caption         =   "Selección de la bomba"
      End
   End
   Begin VB.Menu cccc 
      Caption         =   "Menú principal"
   End
End
Attribute VB_Name = "frmmetodocostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bbbb_Click()
frmHprincipal.Show

End Sub

Private Sub BcalGOTERO_Click()
On Error GoTo mensaje:

ii = Val(txttasint.text)
n = Val(txtviduti.text)
costub = Val(txtcostub.text)
potbomba = Val(txtcdt.text)
NH = Val(txtnh.text)
COSENE = Val(txtcoscom.text)
i = ii / 100
rendimiento = Val(re.text)
costubB = Val(txtcostubB.text)
potbombaB = Val(txtcdtB.text)
energia = Val(Combo1.ListIndex)
Select Case energia
    Case 0
        factorpotencia = 1.15
    Case 1
        factorpotencia = 1.15
    Case 2
        factorpotencia = 1.15
    Case 3
        factorpotencia = 1.1
End Select


COSTUNPOT = COSENE / rendimiento * NH
cospotanu = COSTUNPOT * potbomba
cospotanub = COSTUNPOT * potbombaB
FRC = (i + 1) ^ n * i / ((i + 1) ^ n - 1)
costofijoanualtubo = costub * FRC
costofijoanualtubob = costubB * FRC

total = costofijoanualtubo + cospotanu

totalb = costofijoanualtubob + cospotanub

cae = Format(cospotanu, "##0.0#")
txtfrc = Format(FRC, "##0.00#")
cat = Format(costofijoanualtubo, "##0.0#")
ct = Format(total, "##0.0#")

caeb = Format(cospotanub, "##0.0#")
catb = Format(costofijoanualtubob, "##0.0#")
ctb = Format(totalb, "##0.0#")

'If total < totalb Then
 '   Label32.Visible = True
 '   Label31.Visible = False
'Else
 '   Label32.Visible = False
 '   Label31.Visible = True
'End If

Exit Sub

mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"

End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub blimpiar_Click()
txtcostub = ""
txtcdt = ""
txtcostubB = ""
txtcdtB = ""
txtviduti = ""
txttasint = ""
txtcoscom = ""
txtnh = ""
cae = ""
cat = ""
ct = ""
caeb = ""
catb = ""
ctb = ""
Label31.Visible = False
Label32.Visible = False
txtcostub.SetFocus
txtfrc = ""

End Sub

Private Sub ccc_Click()
frmbomba.Show
End Sub

Private Sub cccc_Click()
Unload Me
frmGeneral.Show

End Sub

Private Sub Combo1_Click()
THC = Combo1.ListIndex
Select Case THC
    Case 0
     Label24.Visible = True
     Label25.Visible = True
     Label29.Visible = False
     Label30.Visible = False
     re = 3.96
    Case 1
     Label24.Visible = True
     Label25.Visible = True
     Label29.Visible = False
     Label30.Visible = False
     re = 2.77
     Case 2
     Label24.Visible = True
     Label25.Visible = True
     Label29.Visible = False
     Label30.Visible = False
     re = 2.51
     Case 3
      Label24.Visible = False
     Label25.Visible = False
     Label29.Visible = True
     Label30.Visible = True
     re = 1.2
End Select
End Sub
Private Sub Combo2_Click()
THC = Combo2.ListIndex
Select Case THC
    Case 0
     Label59.Caption = "$/lts"
     Label62.Caption = "Hp-hr/lts"
     Label39.Caption = "lts"
     re2 = 3.96
    Case 1
     Label59.Caption = "$/lts"
     Label62.Caption = "Hp*hr/lts"
     Label39.Caption = "lts"
     re2 = 2.77
     Case 2
     Label59.Caption = "$/lts"
     Label62.Caption = "Hp*hr/lts"
     Label39.Caption = "lts"
     re2 = 2.51
     Case 3
     Label59.Caption = "$/watts"
     Label62.Caption = "Hp*hr/watts"
     Label39.Caption = "watts"
     re2 = 1.2
End Select
End Sub


Private Sub Command1_Click()
On Error GoTo mensaje:
cc22 = Val(cc2.text)
nh22 = Val(nh2.text)
re22 = Val(re2.text)
pot22 = Val(pot2.text)

consum = nh22 * pot22 / re22
cost = consum * cc22


consumo = Format(consum, "##0.0#")
costo = Format(cost, "##0.0#")
Exit Sub

mensaje:
MsgBox "Error: Digite todos los datos adecuadamente"

End Sub

Private Sub Form_Load()
With Combo1
    .AddItem "Diesel"
    .AddItem "Gasolina"
    .AddItem "Propano"
    .AddItem "Electricidad"
    .ListIndex = 0
    
End With
With Combo2
    .AddItem "Diesel"
    .AddItem "Gasolina"
    .AddItem "Propano"
    .AddItem "Electricidad"
    .ListIndex = 0
    
End With
StatusBar1.Panels(1).text = "Digite los datos de dos diseños y evalue cual es el más económico"

End Sub




Private Sub TabStrip1_Click()
s = TabStrip1.SelectedItem.Index
Select Case s
    Case 1
    Frame9.Visible = True
    Frame10.Visible = False
    StatusBar1.Panels(1).text = "Digite los datos de dos diseños y evalue cual es el más económico"
    txtcostub.SetFocus
    Label19.Caption = "Comparación de costos Energía vrs condución"
    
    Case 2
    Frame9.Visible = False
    Frame10.Visible = True
    StatusBar1.Panels(1).text = "Digite los datos y oprima el boton de calcular para determinar el consumo"
    cc2.SetFocus
    Label19.Caption = "Estimación de consumo de energía por unidad de bombeo"

End Select

End Sub
