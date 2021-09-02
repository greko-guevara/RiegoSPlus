VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtextura 
   Caption         =   "Textura"
   ClientHeight    =   7320
   ClientLeft      =   1290
   ClientTop       =   915
   ClientWidth     =   8430
   Icon            =   "frmtextura.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      Caption         =   "Datos de laboratorio"
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   308
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Frame Frame4 
         Caption         =   "40 segundos"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   7575
         Begin VB.TextBox txtt1 
            Height          =   285
            Left            =   5160
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txth1 
            Height          =   285
            Left            =   2040
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "° Celsius"
            Height          =   255
            Left            =   6480
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Temperatura"
            Height          =   255
            Left            =   3960
            TabIndex        =   33
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Lectura del Hidrómetro "
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "120 minutos"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   7575
         Begin VB.TextBox txtt2 
            Height          =   285
            Left            =   5160
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txth2 
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label56 
            Caption         =   "° Celsius"
            Height          =   255
            Left            =   6480
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Temperatura"
            Height          =   255
            Left            =   3960
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Lectura del Hidrómetro"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox txtpm 
         Height          =   285
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Peso de la muestra"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label55 
         Caption         =   "gramos"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame11 
      Height          =   1215
      Left            =   668
      TabIndex        =   36
      Top             =   5520
      Width           =   7095
      Begin VB.CommandButton BLIMPIAR 
         Caption         =   "&Limpiar"
         Height          =   855
         Left            =   1920
         Picture         =   "frmtextura.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bS 
         Caption         =   "&Menú principal"
         Height          =   855
         Left            =   5280
         Picture         =   "frmtextura.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   3600
         Picture         =   "frmtextura.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Bcalcular 
         Caption         =   "&Calcular"
         Height          =   855
         Left            =   120
         Picture         =   "frmtextura.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "No"
      Height          =   255
      Left            =   6210
      TabIndex        =   14
      Top             =   795
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sí"
      Height          =   255
      Left            =   5610
      TabIndex        =   13
      Top             =   795
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   35
      Top             =   6945
      Width           =   8430
      _ExtentX        =   14870
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
   Begin VB.Frame Frame2 
      Caption         =   "% de Partículas "
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   240
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox txtArena 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLimo 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtTextura 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   4200
         TabIndex        =   20
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtArcilla 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Left            =   5160
         Picture         =   "frmtextura.frx":29F2
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label14 
         Caption         =   "%"
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "%"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "%"
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Limo"
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Arcilla"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Arena"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1035
      Left            =   5400
      Picture         =   "frmtextura.frx":7868
      Top             =   4200
      Width           =   1590
   End
   Begin VB.Label Label6 
      Caption         =   "¿Conoce Usted las relaciones porcentuales  de  Arcilla, Limo y Arena?"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label10 
      Caption         =   "Determinación de Textura por el Método de Bouyoucos"
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
      TabIndex        =   12
      Top             =   120
      Width           =   7695
   End
   Begin VB.Menu motrps 
      Caption         =   "Otros Parámetros"
      Begin VB.Menu mcon 
         Caption         =   "Conductividad"
      End
      Begin VB.Menu mgerm 
         Caption         =   "Generales del Suelo"
      End
      Begin VB.Menu mreger 
         Caption         =   "evapotranspiración"
      End
   End
   Begin VB.Menu mm 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmtextura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Bcalcular_Click()
On Error GoTo mensaje
If Option1.Value = True Then
    arena = Val(txtArena.text)
    arcilla = Val(txtArcilla.text)
    limo = Val(txtLimo.text)
    If arena + arcilla + limo <> 100 Then
    MsgBox "la suma de las partículas debe ser igual a 100", 64, "Textura"
    txtArena.SetFocus
    Exit Sub
    End If
Else
    pm = Val(txtpm.text)
    h1 = Val(txth1.text)
    h2 = Val(txth2.text)
    t1 = Val(txtt1.text)
    t2 = Val(txtt2.text)
    If pm = 0 Then
    MsgBox "Ingrese el valor del peso de la muestra", 64, "Textura"
    txtpm.SetFocus
    Exit Sub
    End If
    If h1 = 0 Then
    MsgBox "Ingrese Lectura del hidrómetro a los 40 seg", 64, "Textura"
    txth1.SetFocus
    Exit Sub
    End If
    If t1 = 0 Then
    MsgBox "Ingrese temperatura (Lectura hidrómetro a los 40 seg)", 64, "Textura"
    txtt1.SetFocus
    Exit Sub
    End If
    If h2 = 0 Then
    MsgBox "Ingrese Lectura del hidrómetro a las 2 hrs", 64, "Textura"
    txth2.SetFocus
    Exit Sub
    End If
    If t2 = 0 Then
    MsgBox "Ingrese temperatura (Lectura hidrómetro a las 2 hrs)", 64, "Textura"
    txtt2.SetFocus
    Exit Sub
    End If
    
    If t1 >= 19.4 Then
            dt = t1 - 19.4
            hh1 = h1 + dt * 0.3
    Else
            dt = 19.4 - t1
            hh1 = h1 + dt * 0.3
    End If
    If t2 >= 19.4 Then
            dt1 = t2 - 19.4
            hh2 = h2 + dt1 * 0.3
    Else
            dt1 = 19.4 - t2
            hh2 = h2 + dt1 * 0.3
    End If
    arena = 100 - hh1 / pm * 100
    arcilla = hh2 / pm * 100
    limo = 100 - arena - arcilla
    txtArena.text = Format(arena, "##0.0")
    txtArcilla.text = Format(arcilla, "##0.0")
    txtLimo = Format(limo, "##0.0")
End If
w = arcilla + arena + limo
If w >= 100.3 Then
    MsgBox "La suma de las tres fracciones no puede ser mayor al 100%", 64, "Determinación de la Textura"
Else
    If arcilla < 25 And arcilla > 7 And limo > 28 And limo < 50 And arena < 52 Then
        textura = "Franco"
    End If
    
    If limo >= 50 And arcilla <= 25 And arcilla > 12 Then
        textura = "Franco Limoso"
    Else
        If limo > 50 And limo < 80 And arcilla < 12 Then
            textura = "Franco Limoso"
        End If
    End If
    
    If limo >= 80 And arcilla <= 12 Then
        textura = "Limoso"
    End If
    
    If arcilla >= 20 And arcilla < 35 And limo < 28 And arena > 45 Then
        textura = "Franco Arcillo Arenoso"
    End If
    
    If arcilla >= 25 And arcilla <= 40 And arena > 20 And arena <= 45 Then
        textura = "Franco Arcilloso"
    End If
    
    If arcilla >= 25 And arcilla <= 40 And arena <= 20 Then
        textura = "Franco Arcillo Limoso"
    End If
    
    If arcilla > 35 And arena > 45 Then
        textura = "Arcillo Arenoso"
    End If
    
    If arcilla > 40 And limo > 40 Then
        textura = "Arcillo Limoso"
    End If
    
    If arcilla > 40 And arena < 45 And limo < 40 Then
        textura = "Arcilloso"
    End If
    
    If arena > 85 And limo + arcilla < 15 Then
        textura = "Arenoso"
    End If
    
    If arena > 75 And arena < 90 And arcilla + limo > 15 And arcilla + limo < 30 Then
        textura = "Arenoso Franco"
    End If
    
    
    If arcilla < 20 And limo + 2 * arcilla > 30 And arena > 52 Then
        textura = "Franco Arenoso"
    Else
        If arcilla < 7 And limo < 50 And arena > 43 And arena < 52 Then
            textura = "Franco Arenoso"
        End If
    End If
    
    txtTextura.text = (textura)
    Frame2.Visible = True
End If
Exit Sub
mensaje:
MsgBox "ingrese valores adecuados", 64, "Textura"
End Sub




Private Sub bimprimir_Click()
Print Form
End Sub

Private Sub blimpiar_Click()
txtpm.text = ""
txth1.text = ""
txth2.text = ""
txtt1.text = ""
txtt2.text = ""
txtArena.text = ""
txtArcilla.text = ""
txtLimo.text = ""
txtTextura.text = ""
Option1.Value = False
Option2.Value = False
Frame1.Visible = False
Frame2.Visible = False

StatusBar1.Panels(1).text = "Determinación de la Textura Método de Bouyoucos"

End Sub

Private Sub bS_Click()
Unload Me

End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False

StatusBar1.Panels(1).text = "Determinación de la Textura Método de Bouyoucos"
End Sub





Private Sub mcon_Click()
frmconductividad.Show
End Sub

Private Sub mgerm_Click()
frmGeneral.Show
End Sub

Private Sub mm_Click()
frmGeneral.Show
Unload Me
End Sub

Private Sub mreger_Click()
frmETO.Show


End Sub

Private Sub Option1_Click()
Frame1.Visible = True
Frame1.Enabled = False
Frame2.Visible = True
StatusBar1.Panels(1).text = "Introduzca las fracciones porcentuales de las partículas y oprima Calcular para Estimar Textura"
txtArena.SetFocus
End Sub

Private Sub Option2_Click()
Frame1.Visible = True
Frame1.Enabled = True
StatusBar1.Panels(1).text = "Introduzca los datos de la prueba de Bouyoucos y oprima Calcular para Estimar la Textura"
txtpm.SetFocus

End Sub


