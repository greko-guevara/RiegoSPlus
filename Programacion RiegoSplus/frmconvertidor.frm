VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmconvertidor 
   Caption         =   "Convertidor de Unidades"
   ClientHeight    =   3315
   ClientLeft      =   1080
   ClientTop       =   2925
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmconvertidor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   8535
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2940
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15169
            MinWidth        =   15169
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
   Begin VB.CommandButton blimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Picture         =   "frmconvertidor.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton bFinalizar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Picture         =   "frmconvertidor.frx":13B4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lbltitulo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   3720
      X2              =   4320
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   3720
      X2              =   4320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Menu mnudist 
      Caption         =   "Distancia"
   End
   Begin VB.Menu mnuarea 
      Caption         =   "Area"
   End
   Begin VB.Menu mnuvol 
      Caption         =   "Volumen"
   End
   Begin VB.Menu mnuvel 
      Caption         =   "Velocidad"
   End
   Begin VB.Menu mq 
      Caption         =   "Caudal"
   End
   Begin VB.Menu mnumasa 
      Caption         =   "Masa"
   End
   Begin VB.Menu mnufuer 
      Caption         =   "Fuerza"
   End
   Begin VB.Menu mnuener 
      Caption         =   "Energía"
   End
   Begin VB.Menu mnupot 
      Caption         =   "Potencia"
   End
   Begin VB.Menu mnupre 
      Caption         =   "Presión"
   End
   Begin VB.Menu mnutemp 
      Caption         =   "Temperatura"
   End
End
Attribute VB_Name = "frmconvertidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UN1 As Double
Dim un2 As Double
Private Sub bFinalizar_Click()
Unload Me

End Sub
Private Sub blimpiar_Click()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
End Sub



Private Sub Form_Load()
StatusBar1.Panels(1).Text = "Seleccione en el menú el tipo de Unidades con que sea trabajar"
End Sub

Rem Area
Private Sub mnuarea_Click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Area"

Combo1.AddItem "metro^2"
Combo1.AddItem "centimetro^2"
Combo1.AddItem "pulgada^2"
Combo1.AddItem "pie^2"
Combo1.AddItem "hectarea"
Combo1.AddItem "acre"
Combo1.AddItem "manzana"

Combo2.AddItem "metro^2"
Combo2.AddItem "centimetro^2"
Combo2.AddItem "pulgada^2"
Combo2.AddItem "pie^2"
Combo2.AddItem "hectarea"
Combo2.AddItem "acre"
Combo2.AddItem "manzana"

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"

Text1.SetFocus
End Sub

Rem DISTANCIA
Private Sub mnudist_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Distancia"

Combo1.AddItem "metros"
Combo1.AddItem "centimetros"
Combo1.AddItem "pulgadas"
Combo1.AddItem "pies"
Combo1.AddItem "yardas"
Combo1.AddItem "milla"
Combo1.AddItem "kilometro"

Combo2.AddItem "metros"
Combo2.AddItem "centimetros"
Combo2.AddItem "pulgadas"
Combo2.AddItem "pies"
Combo2.AddItem "yardas"
Combo2.AddItem "milla"
Combo2.AddItem "kilometro"

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
Text1.SetFocus
End Sub
Rem VOOLUMEN
Private Sub mnuvol_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Volumen"

Combo1.AddItem "metro^3"
Combo1.AddItem "centimetro^3"
Combo1.AddItem "pie^3"
Combo1.AddItem "pulgada^3"
Combo1.AddItem "litro"
Combo1.AddItem "galón"

Combo2.AddItem "metro^3"
Combo2.AddItem "centimetro^3"
Combo2.AddItem "pie^3"
Combo2.AddItem "pulgada^3"
Combo2.AddItem "litro"
Combo2.AddItem "galón"

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"

Text1.SetFocus
End Sub
Rem velocidad
Private Sub mnuvel_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Velocidad"

Combo1.AddItem "mts/seg"
Combo1.AddItem "ft/seg"
Combo1.AddItem "km/h"
Combo1.AddItem "mi/h"
Combo1.AddItem "nudos"

Combo2.AddItem "mts/seg"
Combo2.AddItem "ft/seg"
Combo2.AddItem "km/h"
Combo2.AddItem "mi/h"
Combo2.AddItem "nudos"

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
Text1.SetFocus
End Sub
Rem masa
Private Sub mnumasa_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Masa"

Combo1.AddItem "Kilogramo"
Combo1.AddItem "libra"
Combo1.AddItem "onza"
Combo1.AddItem "tonelada"
Combo1.AddItem "slug"

Combo2.AddItem "Kilogramo"
Combo2.AddItem "libra"
Combo2.AddItem "onza"
Combo2.AddItem "tonelada"
Combo2.AddItem "slug"

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
Text1.SetFocus
End Sub

Rem fuerza

Private Sub mnufuer_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Fuerza"

Combo1.AddItem "Newton"
Combo1.AddItem "kilogramo F"
Combo1.AddItem "libra F"

Combo2.AddItem "Newton"
Combo2.AddItem "kilogramo F"
Combo2.AddItem "libra F"

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
Text1.SetFocus
End Sub

Rem Energia

Private Sub mnuener_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Energía"

Combo1.AddItem "Joules (N*m)"
Combo1.AddItem "Calorias"
Combo1.AddItem "ft*lb"

Combo2.AddItem "Joules (N*m)"
Combo2.AddItem "Calorias"
Combo2.AddItem "ft*lb"
Text1.SetFocus

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
End Sub

Rem potencia
Private Sub mnupot_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Potencia"

Combo1.AddItem "watt"
Combo1.AddItem "HP"

Combo2.AddItem "watt"
Combo2.AddItem "HP"
Text1.SetFocus

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
End Sub

Rem presión
Private Sub mnupre_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Presión"

Combo1.AddItem "kg/cm^2"
Combo1.AddItem "Psi"
Combo1.AddItem "Pa"
Combo1.AddItem "atm"
Combo1.AddItem "bar"
Combo1.AddItem "m.c.H2O"

Combo2.AddItem "kg/cm^2"
Combo2.AddItem "Psi"
Combo2.AddItem "Pa"
Combo2.AddItem "atm"
Combo2.AddItem "bar"
Combo2.AddItem "m.c.H2O"
Text1.SetFocus

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
End Sub

Rem Temperatura
Private Sub mnutemp_click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Temperatura"

Combo1.AddItem "°C"
Combo1.AddItem "°F"
Combo1.AddItem "Kelvin"

Combo2.AddItem "°C"
Combo2.AddItem "°F"
Combo2.AddItem "Kelvin"
Text1.SetFocus

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"
End Sub

Private Sub Combo2_Click()
UN1 = Val(Text1.Text)
unidad = Combo2.Text
If UN1 = 0 Then
    MsgBox "Introduzca el valor a transformar", 64, "convertidor de unidades"
    Text1.SetFocus
    Exit Sub
End If
If Combo1 = "" Then
    MsgBox "Seleccione la unidad de entrada", 64, "convertidor de unidades"
    Combo1.SetFocus
    Exit Sub
End If

Rem distancia

If Combo1.Text = "metros" Then
    Select Case unidad
    Case "metros"
        un2 = UN1
    Case "centimetros"
        un2 = UN1 * 100
    Case "pulgadas"
        un2 = UN1 * 39.37008
    Case "pies"
        un2 = UN1 * 3.28084
    Case "yardas"
        un2 = UN1 * 1.09361
    Case "milla"
        un2 = UN1 * 0.00062
    Case "kilometro"
        un2 = UN1 * 0.001
    End Select
End If
If Combo1.Text = "centimetros" Then
    Select Case unidad
    Case "metros"
        un2 = UN1 / 100
    Case "centimetros"
        un2 = UN1
    Case "pulgadas"
        un2 = UN1 * 39.37008 / 100
    Case "pies"
        un2 = UN1 * 3.28084 / 100
    Case "yardas"
        un2 = UN1 * 1.09361 / 100
    Case ",milla"
        un2 = UN1 * 0.000062
    Case "kilometro"
        un2 = UN1 * 0.00001
    End Select
End If
If Combo1.Text = "pulgadas" Then
    Select Case unidad
    Case "metros"
        un2 = UN1 * 0.0254
    Case "centimetros"
        un2 = UN1 * 2.54
    Case "pulgadas"
        un2 = UN1
    Case "pies"
        un2 = UN1 * 0.08333
    Case "yardas"
        un2 = UN1 * 0.02778
    Case "milla"
        un2 = UN1 * 0.00002
    Case "kilometro"
        un2 = UN1 * 0.00003
    End Select
End If
If Combo1.Text = "pies" Then
    Select Case unidad
    Case "metros"
        un2 = UN1 * 0.3048
    Case "centimetros"
        un2 = UN1 * 30.48
    Case "pulgadas"
        un2 = UN1 * 12
    Case "pies"
        un2 = UN1
    Case "yardas"
        un2 = UN1 * 0.33333
    Case "milla"
        un2 = UN1 * 0.00019
    Case "kilometro"
        un2 = UN1 * 0.0003
    End Select
End If
If Combo1.Text = "yardas" Then
    Select Case unidad
    Case "metros"
        un2 = UN1 * 0.9144
    Case "centimetros"
        un2 = UN1 * 91.44
    Case "pulgadas"
        un2 = UN1 * 36
    Case "pies"
        un2 = UN1 * 3
    Case "yardas"
        un2 = UN1
    Case "milla"
        un2 = UN1 * 0.00057
    Case "kilometro"
        un2 = UN1 * 0.00091
    End Select
End If
If Combo1.Text = "milla" Then
    Select Case unidad
    Case "metros"
        un2 = UN1 * 1609.344
    Case "centimetros"
        un2 = UN1 * 160934.4
    Case "pulgadas"
        un2 = UN1 * 63360
    Case "pies"
        un2 = UN1 * 5280
    Case "yardas"
        un2 = UN1 * 1760
    Case "milla"
        un2 = UN1
    Case "kilometro"
        un2 = UN1 * 1.60934
    End Select
End If
If Combo1.Text = "kilometro" Then
    Select Case unidad
    Case "metros"
        un2 = UN1 * 1000
    Case "centimetros"
        un2 = UN1 * 100000
    Case "pulgadas"
        un2 = UN1 * 39370.07874
    Case "pies"
        un2 = UN1 * 3280.8399
    Case "yardas"
        un2 = UN1 * 1093.6133
    Case "milla"
        un2 = UN1 * 0.62137
    Case "kilometro"
        un2 = UN1
    End Select
End If
Rem area

If Combo1.Text = "metro^2" Then
    Select Case unidad
    Case "metro^2"
        un2 = UN1
    Case "centimetro^2"
        un2 = UN1 * 10000
    Case "pulgada^2"
        un2 = UN1 * 39.37008 ^ 2
    Case "pie^2"
        un2 = UN1 * 3.28084 ^ 2
    Case "hectarea"
        un2 = UN1 * 0.0001
    Case "acre"
        un2 = UN1 * 0.00025
    Case "manzana"
        un2 = UN1 * 0.01
    End Select
End If
If Combo1.Text = "centimetro^2" Then
    Select Case unidad
    Case "metro^2"
        un2 = UN1 / 10000
    Case "centimetro^2"
        un2 = UN1
    Case "pulgada^2"
        un2 = UN1 * 39.37008 ^ 2 / 10000
    Case "pie^2"
        un2 = UN1 * 3.28084 ^ 2 / 10000
    Case "hectarea"
        un2 = UN1 * 0.0001 / 10000
    Case "acre"
        un2 = UN1 * 0.00025 / 10000
    Case "manzana"
        un2 = UN1 * 0.01 / 10000
    End Select
End If
If Combo1.Text = "pulgada^2" Then
    Select Case unidad
    Case "metro^2"
        un2 = UN1 * 0.00065
    Case "centimetro^2"
        un2 = UN1 * 6.4516
    Case "pulgada^2"
        un2 = UN1
    Case "pie^2"
        un2 = UN1 * 0.00694
    Case "hectarea"
        un2 = UN1 * 6.45 * 10 ^ -8
    Case "acre"
        un2 = UN1 * 1.59422 * 10 ^ -7
    Case "manzana"
        un2 = UN1 * 0.00001
    End Select
End If
If Combo1.Text = "pie^2" Then
    Select Case unidad
    Case "metro^2"
        un2 = UN1 * 0.0929
    Case "centimetro^2"
        un2 = UN1 * 929.0304
    Case "pulgada^2"
        un2 = UN1 * 144
    Case "pie^2"
        un2 = UN1
    Case "hectarea"
        un2 = UN1 * 0.00001
    Case "acre"
        un2 = UN1 * 0.00002
    Case "manzana"
        un2 = UN1 * 0.00093
    End Select
End If
If Combo1.Text = "hectarea" Then
    Select Case unidad
    Case "metro^2"
        un2 = UN1 * 10000
    Case "centimetro^2"
        un2 = UN1 * 100000000
    Case "pulgada^2"
        un2 = UN1 * 15500031.0001
    Case "pie^2"
        un2 = UN1 * 107639.10417
    Case "hectarea"
        un2 = UN1
    Case "acre"
        un2 = UN1 * 2.47104
    Case "manzana"
        un2 = UN1 * 100
    End Select
End If
If Combo1.Text = "acre" Then
    Select Case unidad
    Case "metro^2"
        un2 = UN1 * 4046.87261
    Case "centimetro^2"
        un2 = UN1 * 4046.87261 * 10000
    Case "pulgada^2"
        un2 = UN1 * 6272665.09063
    Case "pie^2"
        un2 = UN1 * 43560.17424
    Case "hectarea"
        un2 = UN1 * 0.40469
    Case "acre"
        un2 = UN1
    Case "manzana"
        un2 = UN1 * 40.46873
    End Select
End If
If Combo1.Text = "manzana" Then
    Select Case unidad
    Case "metro^2"
        un2 = UN1 * 100
    Case "centimetro^2"
        un2 = UN1 * 1000000
    Case "pulgada^2"
        un2 = UN1 * 155000.31
    Case "pie^2"
        un2 = UN1 * 1076.39104
    Case "hectarea"
        un2 = UN1 * 0.01
    Case "acre"
        un2 = UN1 * 0.02471
    Case "manzana"
        un2 = UN1
    End Select
End If

Rem volumen

If Combo1.Text = "metro^3" Then
    Select Case unidad
    Case "metro^3"
        un2 = UN1
    Case "centimetro^3"
        un2 = UN1 * 1000000
    Case "pulgada^3"
        un2 = UN1 * 61023.74409
    Case "pie^3"
        un2 = UN1 * 35.31467
    Case "litro"
        un2 = UN1 * 1000
    Case "galón"
        un2 = UN1 * 264.17205
    End Select
End If
If Combo1.Text = "centimetro^3" Then
    Select Case unidad
    Case "metro^3"
        un2 = UN1 / 1000000
    Case "centimetro^3"
        un2 = UN1
    Case "pulgada^3"
        un2 = UN1 * 61023.74409 / 1000000
    Case "pie^3"
        un2 = UN1 * 35.31467 / 1000000
    Case "litro"
        un2 = UN1 * 1000 / 1000000
    Case "galón"
        un2 = UN1 * 264.17205 / 1000000
    End Select
End If
If Combo1.Text = "pulgada^3" Then
    Select Case unidad
    Case "metro^3"
        un2 = UN1 * 0.00002
    Case "centimetro^3"
        un2 = UN1 * 16.38706
    Case "pulgada^3"
        un2 = UN1
    Case "pie^3"
        un2 = UN1 * 0.00058
    Case "litro"
        un2 = UN1 * 0.01639
    Case "galón"
        un2 = UN1 * 0.00433
    End Select
End If
If Combo1.Text = "pie^3" Then
    Select Case unidad
    Case "metro^3"
        un2 = UN1 * 0.02832
    Case "centimetro^3"
        un2 = UN1 * 28316.84659
    Case "pulgada^3"
        un2 = UN1 * 1728
    Case "pie^3"
        un2 = UN1
    Case "litro"
        un2 = UN1 * 28.31685
    Case "galón"
        un2 = UN1 * 7.48052
    End Select
End If
If Combo1.Text = "litro" Then
    Select Case unidad
    Case "metro^3"
        un2 = UN1 * 0.001
    Case "centimetro^3"
        un2 = UN1 * 1000
    Case "pulgada^3"
        un2 = UN1 * 61.02374
    Case "pie^3"
        un2 = UN1 * 0.03531
    Case "litro"
        un2 = UN1
    Case "galón"
        un2 = UN1 * 0.26417
    End Select
End If
If Combo1.Text = "galón" Then
    Select Case unidad
    Case "metro^3"
        un2 = UN1 * 0.00379
    Case "centimetro^3"
        un2 = UN1 * 3785.41178
    Case "pulgada^3"
        un2 = UN1 * 231
    Case "pie^3"
        un2 = UN1 * 0.13368
    Case "litro"
        un2 = UN1 * 3.78541
    Case "galón"
        un2 = UN1
    End Select
End If

Rem velocidad

If Combo1.Text = "mts/seg" Then
    Select Case unidad
    Case "mts/seg"
        un2 = UN1
    Case "ft/seg"
        un2 = UN1 * 3.28084
    Case "km/h"
        un2 = UN1 * 3.6
    Case "mi/h"
        un2 = UN1 * 2.23694
    Case "nudos"
        un2 = UN1 * 1.94384
    End Select
End If
If Combo1.Text = "ft/seg" Then
    Select Case unidad
    Case "mts/seg"
        un2 = UN1 * 0.3048
    Case "ft/seg"
        un2 = UN1
    Case "km/h"
        un2 = UN1 * 1.09728
    Case "mi/h"
        un2 = UN1 * 0.68182
    Case "nudos"
        un2 = UN1 * 0.59248
    End Select
End If
If Combo1.Text = "km/h" Then
    Select Case unidad
    Case "mts/seg"
        un2 = UN1 * 0.27778
    Case "ft/seg"
        un2 = UN1 * 0.91134
    Case "km/h"
        un2 = UN1
    Case "mi/h"
        un2 = UN1 * 0.62137
    Case "nudos"
        un2 = UN1 * 0.53996
    End Select
End If
If Combo1.Text = "mi/h" Then
    Select Case unidad
    Case "mts/seg"
        un2 = UN1 * 0.44704
    Case "ft/seg"
        un2 = UN1 * 1.46667
    Case "km/h"
        un2 = UN1 * 1.60934
    Case "mi/h"
        un2 = UN1
    Case "nudos"
        un2 = UN1 * 0.86898
    End Select
End If
If Combo1.Text = "nudos" Then
    Select Case unidad
    Case "mts/seg"
        un2 = UN1 * 0.51444
    Case "ft/seg"
        un2 = UN1 * 1.68781
    Case "km/h"
        un2 = UN1 * 1.852
    Case "mi/h"
        un2 = UN1 * 1.15078
    Case "nudos"
        un2 = UN1
    End Select
End If

Rem masa

If Combo1.Text = "Kilogramo" Then
    Select Case unidad
    Case "Kilogramo"
        un2 = UN1
    Case "libra"
        un2 = UN1 * 2.20462
    Case "onza"
        un2 = UN1 * 35.27396
    Case "tonelada"
        un2 = UN1 * 0.001
    Case "slug"
        un2 = UN1 * 0.06852
    End Select
End If
If Combo1.Text = "libra" Then
    Select Case unidad
    Case "Kilogramo"
        un2 = UN1 * 0.45359
    Case "libra"
        un2 = UN1
    Case "onza"
        un2 = UN1 * 16
    Case "tonelada"
        un2 = UN1 * 0.00045
    Case "slug"
        un2 = UN1 * 0.03108
    End Select
End If
If Combo1.Text = "onza" Then
    Select Case unidad
    Case "Kilogramo"
        un2 = UN1 * 0.02835
    Case "libra"
        un2 = UN1 * 0.0625
    Case "onza"
        un2 = UN1
    Case "tonelada"
        un2 = UN1 * 0.00003
    Case "slug"
        un2 = UN1 * 0.00194
    End Select
End If
If Combo1.Text = "tonelada" Then
    Select Case unidad
    Case "Kilogramo"
        un2 = UN1 * 1000
    Case "libra"
        un2 = UN1 * 2204.62262
    Case "onza"
        un2 = UN1 * 35273.96195
    Case "tonelada"
        un2 = UN1
    Case "slug"
        un2 = UN1 * 68.52177
    End Select
End If
If Combo1.Text = "slug" Then
    Select Case unidad
    Case "Kilogramo"
        un2 = UN1 * 14.5939
    Case "libra"
        un2 = UN1 * 32.1705
    Case "onza"
        un2 = UN1 * 514.78478
    Case "tonelada"
        un2 = UN1 * 0.0145939
    Case "slug"
        un2 = UN1
    End Select
End If

Rem fuerza

If Combo1.Text = "Newton" Then
    Select Case unidad
    Case "Newton"
        un2 = UN1
    Case "kilogramo F"
        un2 = UN1 * 0.10197
    Case "libra F"
        un2 = UN1 * 0.22481
    End Select
End If
If Combo1.Text = "kilogramo F" Then
    Select Case unidad
    Case "Newton"
        un2 = UN1 * 9.80665
    Case "kilogramo F"
        un2 = UN1
    Case "libra F"
        un2 = UN1 * 2.20462
    End Select
End If
If Combo1.Text = "libra F" Then
    Select Case unidad
    Case "Newton"
        un2 = UN1 * 4.44822
    Case "kilogramo F"
        un2 = UN1 * 0.45359
    Case "libra F"
        un2 = UN1
    End Select
End If

Rem energia

If Combo1.Text = "Joules (N*m)" Then
    Select Case unidad
    Case "Joules (N*m)"
        un2 = UN1
    Case "Calorias"
        un2 = UN1 * 0.23885
    Case "ft*lb"
        un2 = UN1 * 0.73756
    End Select
End If
If Combo1.Text = "Calorias" Then
    Select Case unidad
    Case "Joules (N*m)"
        un2 = UN1 * 4.1868
    Case "Calorias"
        un2 = UN1
    Case "ft*lb"
        un2 = UN1 * 3.08803
    End Select
End If
If Combo1.Text = "ft*lb" Then
    Select Case unidad
    Case "Joules (N*m)"
        un2 = UN1 * 1.35582
    Case "Calorias"
        un2 = UN1 * 0.32383
    Case "ft*lb"
        un2 = UN1
    End Select
End If

Rem potencia

If Combo1.Text = "watt" Then
    Select Case unidad
    Case "watt"
        un2 = UN1
    Case "HP"
        un2 = UN1 * 0.00134
    End Select
End If
If Combo1.Text = "HP" Then
    Select Case unidad
    Case "watt"
        un2 = UN1 * 745.69987
    Case "HP"
        un2 = UN1 * 1
    End Select
End If

Rem presion

If Combo1.Text = "kg/cm^2" Then
    Select Case unidad
    Case "kg/cm^2"
        un2 = UN1
    Case "Psi"
        un2 = UN1 * 14.22273
    Case "Pa"
        un2 = UN1 * 98062.2714
    Case "atm"
        un2 = un * 0.9678
    Case "bar"
        un2 = UN1 * 0.98062
    Case "m.c.H2O"
        un2 = UN1 * 10
    End Select
End If
If Combo1.Text = "Psi" Then
    Select Case unidad
    Case "kg/cm^2"
        un2 = UN1 * 0.07031
    Case "Psi"
        un2 = UN1
    Case "Pa"
        un2 = UN1 * 6894.75729
    Case "atm"
        un2 = un * 0.06805
    Case "bar"
        un2 = UN1 * 0.06895
    Case "m.c.H2O"
        un2 = UN1 * 0.7031
    End Select
End If
If Combo1.Text = "Pa" Then
    Select Case unidad
    Case "kg/cm^2"
        un2 = UN1 / 98062.2714
    Case "Psi"
        un2 = UN1 * 0.00015
    Case "Pa"
        un2 = UN1
    Case "atm"
        un2 = un * 0.00001
    Case "bar"
        un2 = UN1 * 0.00001
    Case "m.c.H2O"
        un2 = UN1 * 0.000102
    End Select
End If
If Combo1.Text = "atm" Then
    Select Case unidad
    Case "kg/cm^2"
        un2 = UN1 * 1.03327
    Case "Psi"
        un2 = UN1 * 14.69595
    Case "Pa"
        un2 = UN1 * 101325
    Case "atm"
        un2 = un
    Case "bar"
        un2 = UN1 * 1.01325
    Case "m.c.H2O"
        un2 = UN1 * 10.3426097
    End Select
End If
If Combo1.Text = "bar" Then
    Select Case unidad
    Case "kg/cm^2"
        un2 = UN1 * 1.01976
    Case "Psi"
        un2 = UN1 * 14.50377
    Case "Pa"
        un2 = UN1 * 100000
    Case "atm"
        un2 = un * 0.98692
    Case "bar"
        un2 = UN1
    Case "m.c.H2O"
        un2 = UN1 * 10.19762
    End Select
End If
If Combo1.Text = "m.c.H2O" Then
    Select Case unidad
    Case "kg/cm^2"
        un2 = UN1 / 10
    Case "Psi"
        un2 = UN1 * 1.422273
    Case "Pa"
        un2 = UN1 * 9806.206456
    Case "atm"
        un2 = un * 0.09678
    Case "bar"
        un2 = UN1 * 0.09806206456
    Case "m.c.H2O"
        un2 = UN1
    End Select
End If

Rem temperatura

If Combo1.Text = "°C" Then
    Select Case unidad
    Case "°C"
        un2 = UN1
    Case "°F"
        un2 = (9 * UN1 / 5) + 32
    Case "Kelvin"
        un2 = UN1 + 273.15
    End Select
End If
If Combo1.Text = "°F" Then
    Select Case unidad
    Case "°C"
        un2 = (UN1 - 32) * 5 / 9
    Case "°F"
        un2 = UN1
    Case "Kelvin"
        un2 = ((UN1 - 32) * 5 / 9) + 273.15
    End Select
End If
If Combo1.Text = "Kelvin" Then
    Select Case unidad
    Case "°C"
        un2 = UN1 - 273.15
    Case "°F"
        un2 = ((UN1 - 273.15) * 9 / 5) + 32
    Case "Kelvin"
        un2 = UN1
    End Select
End If

'determinacion del caudal *********-------*

If Combo1.Text = "m3/seg" Then
    Select Case unidad
    Case "m3/seg"
        un2 = UN1
    Case "m3/dia"
        un2 = UN1 * 86400
    Case "l/seg"
        un2 = UN1 * 1000
    Case "l/min"
        un2 = UN1 * 60000
    Case "gal/min"
        un2 = UN1 * 60000 / 3.7854
    End Select
End If
If Combo1.Text = "m3/dia" Then
    Select Case unidad
    Case "m3/seg"
        un2 = UN1 / 86400
    Case "m3/dia"
        un2 = UN1
    Case "l/seg"
        un2 = UN1 * 0.01157
    Case "l/min"
        un2 = UN1 * 0.69444
    Case "gal/min"
        un2 = UN1 * 0.69444 / 3.7854
    End Select
End If
If Combo1.Text = "l/seg" Then
    Select Case unidad
    Case "m3/seg"
        un2 = UN1 / 1000
    Case "m3/dia"
        un2 = UN1 / 0.01157
    Case "l/seg"
        un2 = UN1
    Case "l/min"
        un2 = UN1 * 60
    Case "gal/min"
        un2 = UN1 * 60 / 3.7854
    End Select
End If
If Combo1.Text = "l/min" Then
    Select Case unidad
    Case "m3/seg"
        un2 = UN1 / 60000
    Case "m3/dia"
        un2 = UN1 / 0.69444
    Case "l/seg"
        un2 = UN1 / 60
    Case "l/min"
        un2 = UN1
    Case "gal/min"
        un2 = UN1 / 3.7854
    End Select
End If
If Combo1.Text = "gal/min" Then
    Select Case unidad
    Case "m3/seg"
        un2 = UN1 * 3.7854 / 60000
    Case "m3/dia"
        un2 = UN1 * 3.7854 / 0.69444
    Case "l/seg"
        un2 = UN1 * 3.7854 / 60
    Case "l/min"
        un2 = UN1 * 3.7854
    Case "gal/min"
        un2 = UN1
    End Select
End If
Text2.Text = Format(un2, "##,##0.0000")
End Sub


Private Sub mq_Click()
Combo1.Clear
Combo2.Clear
lbltitulo.Caption = "Unidades de Caudal"

Combo1.AddItem "m3/seg"
Combo1.AddItem "m3/dia"
Combo1.AddItem "l/seg"
Combo1.AddItem "l/min"
Combo1.AddItem "gal/min"
Combo2.AddItem "m3/seg"
Combo2.AddItem "m3/dia"
Combo2.AddItem "l/seg"
Combo2.AddItem "l/min"
Combo2.AddItem "gal/min"

StatusBar1.Panels(1).Text = "Digite el valor a convertir, seleccione la unidad de entrada y luego seleccine la unidad de Salida"

Text1.SetFocus

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case 13
 Combo1.SetFocus
Case Is < 46, Is > 57
 KeyAscii = 0
End Select
End Sub
