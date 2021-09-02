VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalendario 
   Caption         =   "Calendario de Riego"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11850
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   480
      TabIndex        =   12
      Top             =   6000
      Width           =   7215
      Begin VB.CommandButton blimpiar 
         Caption         =   "&Limpiar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmCalendario.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bimprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   3720
         Picture         =   "frmCalendario.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bfinailizar 
         Caption         =   "&Menú principal"
         Height          =   735
         Left            =   5520
         Picture         =   "frmCalendario.frx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton bevaluar 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   120
         Picture         =   "frmCalendario.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Introduzca datos:"
      ForeColor       =   &H00800000&
      Height          =   4215
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame Frame8 
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Selec. Mes"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdD 
         Height          =   3255
         Left            =   2280
         TabIndex        =   2
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   4
         FixedCols       =   2
         BackColorFixed  =   -2147483626
         GridColor       =   16761024
      End
      Begin VB.Label Label5 
         Caption         =   "Oprima  (*) para eliminar líneas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Oprima Enter para insertar  líneas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lbletiqueta 
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   3375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grddd 
      Height          =   3495
      Left            =   8040
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   50
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483626
      ForeColor       =   128
      GridColor       =   16761024
      AllowUserResizing=   1
   End
   Begin MSComCtl2.MonthView MonthView1 
      Bindings        =   "frmCalendario.frx":29F2
      Height          =   2370
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      StartOfWeek     =   20971521
      CurrentDate     =   38171
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
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
            TextSave        =   "23/04/2008"
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
   Begin MSComDlg.CommonDialog cdCrear 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir Proyecto"
      Filter          =   ".dat"
   End
   Begin MSComDlg.CommonDialog cdAccesar 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Guardar como"
      Filter          =   ".dat"
   End
   Begin VB.Label Label4 
      Caption         =   "Calendario de Riego"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   6120
      Picture         =   "frmCalendario.frx":29FD
      Top             =   600
      Width           =   1650
   End
   Begin VB.Label Label3 
      Caption         =   "Seleccione en el calendario la fecha de Inicio del Riego"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "Calendario de Riego"
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
      TabIndex        =   8
      Top             =   360
      Width           =   3015
   End
   Begin VB.Menu marchi 
      Caption         =   "Archivo"
      Begin VB.Menu mguardar 
         Caption         =   "Guardar como "
         Shortcut        =   ^G
      End
      Begin VB.Menu mabrir 
         Caption         =   "Abrir archivo"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mpaesuelcli 
      Caption         =   "Parámetros Suelo- Clima"
      Begin VB.Menu mgs 
         Caption         =   "Generales Suelo"
      End
      Begin VB.Menu mt 
         Caption         =   "Textura"
      End
      Begin VB.Menu mc 
         Caption         =   "Conductividad Hidráulica"
      End
      Begin VB.Menu me 
         Caption         =   "Evaportranspiración"
      End
   End
   Begin VB.Menu ma 
      Caption         =   "Otros"
      Begin VB.Menu mconv 
         Caption         =   "Convertidor de Unidades"
      End
   End
   Begin VB.Menu mm 
      Caption         =   "Menú Principal"
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u As Integer
Dim h As Integer

Dim valorrmes(0 To 50) As Single
Dim valormes(0 To 50) As String
Dim valorn(1 To 50) As Single
Dim valornd(0 To 50) As Single
Dim valoretr(1 To 50) As Single
Dim valoretra(0 To 50) As Single
Dim valorln(1 To 50) As Single
Dim valormes1(1 To 50) As String
Dim valornd1(1 To 50) As Single
Dim valorln1(1 To 50) As Single

Public dia As String
Public mes As String

Rem funcion de la grid
Dim row, col, numero, num(0 To 1000)
Dim n As Integer, i, ii, punto




Private Sub bevaluar_Click()
'On Error GoTo mensaje
If dia = "" Then
MsgBox "Seleccione en el calendario la fecha de inicio", 64, "Calendario de Riego"
Exit Sub
Else
'declaracion de matrices
sx = 0
sx2 = 0
For j% = 1 To (u + 1)
    valormes(j%) = grdD.TextMatrix(j%, 0)
    valorn(j%) = grdD.TextMatrix(j%, 1)
    valoretr(j%) = grdD.TextMatrix(j%, 2)
    valorln(j%) = grdD.TextMatrix(j%, 3)
    sx = sx + valorn(j%)
    sx2 = sx2 + valoretr(j%)
    valornd(j%) = sx
    valoretra(j%) = sx2
Next j%

    '-------------primer riego----------------------
    ln1 = valorln(1)
    h = Val(u)
    u = 1
    valormes1(u) = meess
    valornd1(u) = dia
    valorln1(u) = ln1
    grddd.Visible = True
    Label4.Visible = True
    
    With grddd
        .Clear
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(1) = 4
        .TextMatrix(0, 0) = "Mes"
        .TextMatrix(0, 1) = "día"
        .TextMatrix(0, 2) = "Lámina"
        .TextMatrix(1, 0) = (valormes(1))
        .TextMatrix(1, 1) = Format(dia, "##0")
        .TextMatrix(1, 2) = Format(ln1, "##0.0##")
    End With
   If dia <> 1 Then
    ln1 = ln1 + (valoretra(1) - valoretra(0)) / valornd(1) * dia
    End If
    
    For k% = 1 To 35
    u = u + 1
     
       For j% = 1 To h + 1
        'proximo riego
        If ln1 > valoretra((j%) - 1) And ln1 < valoretra(j%) Then
            m = ((valornd(j%) - valornd((j%) - 1)) / (valoretra(j%) - valoretra((j%) - 1))) ^ -1
            T = Int((ln1 - valoretra((j%) - 1)) / m)
            'determinacion de mes dia lamina
            
            mes1 = valormes(j%)
            lnn1 = valorln(j%)
        
        Else
            If ln1 = valoretra(j%) Then
            T = 1
            'determinacion de mes dia lamina
            
            mes1 = valormes(j% + 1)
            lnn1 = valorln(j% + 1)
            End If
             
        End If
            valorln1(u) = lnn1
            valormes1(u) = mes1
            valornd1(u) = T
            With grddd
                .TextMatrix(u, 0) = (mes1)
                .TextMatrix(u, 1) = Format(T, "##0")
                .TextMatrix(u, 2) = Format(lnn1, "##0.00#")
                
            End With
            Next j%
        ln1 = ln1 + valorln1(k%)
    
    Next k%

End If
Exit Sub
mensaje:
MsgBox "Introduzca adecuados Valores de ETR, Ln y fecha de Inicio", 64, "Calendario de Riego"

End Sub

Private Sub bfinailizar_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub bimprimir_Click()
Print Form

End Sub

Private Sub blimpiar_Click()


u = 0
With grdD
    .Clear
    .Rows = 2
    .ColAlignment(0) = 4
    .ColAlignment(1) = 4
    .ColAlignment(2) = 4
    .ColAlignment(3) = 4
    .TextMatrix(0, 0) = "Mes "
    .TextMatrix(0, 1) = "N° días"
    .TextMatrix(0, 2) = "ETR (mm/día)"
    .TextMatrix(0, 3) = "Lámina (mm)"
End With
With grddd
    .Visible = False
    .Clear
End With
Frame1.Visible = False
dia = ""
mes = ""
año = ""

lbletiqueta.Caption = ""
Label4.Caption = ""
End Sub




Private Sub Form_Load()
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
u = 0
With grdD
    .ColAlignment(0) = 4
    .ColAlignment(1) = 4
    .ColAlignment(2) = 4
    .ColAlignment(3) = 4
    .ColWidth(0) = 1200
    
    .ColWidth(1) = 800
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .TextMatrix(0, 0) = "Mes "
    .TextMatrix(0, 1) = "N° días"
    .TextMatrix(0, 2) = "ETR mm/mes"
    .TextMatrix(0, 3) = "Lámina mm"
End With
StatusBar1.Panels(1).text = "Seleccione la fecha de Inicio y cargue los valores de de ETr y Ln para los meses de Operación del riego"
End Sub




Private Sub grdD_Click()
i = ""
punto = 0
End Sub

Private Sub grdD_KeyPress(KeyAscii As Integer)
On Error GoTo mensaje
If grdD.col <> col Or grdD.row <> row Then
i = ""
punto = 0
For s = 1 To numero
num(s) = ""
Next s
numero = 0
End If
num(0) = ""
If KeyAscii = 45 And i = "" Then
i = i + "-"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 48 Then
    i = i + "0"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If


If punto <> 1 Then
If KeyAscii = 44 Or KeyAscii = 46 Then
    numero = numero + 1
    If i = "" Then
    i = i + "0."
    grdD.text = i
    num(numero) = i
    punto = 1
Else
    i = i + "."
    grdD.text = i
    num(numero) = i
    punto = 1
End If
End If
End If


If KeyAscii = 49 Then
    i = i + "1"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 50 Then
    i = i + "2"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 51 Then
    i = i + "3"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 52 Then
    i = i + "4"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 53 Then
    i = i + "5"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 54 Then
    i = i + "6"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 55 Then
    i = i + "7"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 56 Then
    i = i + "8"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If

If KeyAscii = 57 Then
    i = i + "9"
    grdD.text = i
    numero = numero + 1
    num(numero) = i
End If
Rem tecla de borrado
If numero >= 1 Then
If KeyAscii = 8 Then
i = num(numero - 1)
numero = numero - 1
grdD.text = i
End If
Else

grdD.text = ""
End If

Rem tecla para eliminar
If KeyAscii = 42 Then
If u >= 2 Then
    u = u - 1
    grdD.Rows = u + 2
End If
'*******************************************************************************
End If
If KeyAscii = 13 Then
u = u + 1
grdD.Rows = u + 2

    If grdD.TextMatrix(1, 0) <> "" Then
    If valorrmes(u) = 11 Then
        valorrmes(u + 1) = 0
    Else
        valorrmes(u + 1) = Val(valorrmes(u)) + 1
    End If
           Select Case valorrmes(u)
                Case 0
                    NN = 31
                Case 1
                    NN = 28
                Case 2
                    NN = 31
                Case 3
                    NN = 30
                Case 4
                    NN = 31
                Case 5
                    NN = 30
                Case 6
                    NN = 31
                Case 7
                    NN = 31
                Case 8
                    NN = 30
                Case 9
                    NN = 31
                Case 10
                    NN = 30
                Case 11
                    NN = 31
            End Select
        grdD.TextMatrix(u + 1, 1) = NN
        If Combo1.ListIndex = 11 Then
            Combo1.ListIndex = 0
            grdD.TextMatrix(u + 1, 0) = Combo1.text
        Else
        Combo1.ListIndex = Combo1.ListIndex + 1
        grdD.TextMatrix(u + 1, 0) = Combo1.text
        End If
    End If
End If
'*********************************************************************************



Rem pruevas grid1.TextMatrix(numero, 6) = num(numero)

Rem grdDatos.Text = KeyAscii
col = grdD.col
row = grdD.row
Exit Sub
mensaje:
MsgBox "Error desconocido al digitar"

End Sub


Private Sub mabrir_Click()
On Error GoTo SinArchivo:
 cdAccesar.ShowOpen
 NombreArch = cdAccesar.FileName
 u = 0
 Frame1.Visible = True
  
 Open NombreArch For Random As #1 Len = Len(cuatroMNEL)
 NumReg = LOF(1) \ Len(cuatroMNEL)
 grdD.Rows = NumReg + 1
 For j% = 1 To NumReg
  Get #1, j%, cuatroMNEL
  
  mmm = cuatroMNEL.mm
  nnn = cuatroMNEL.NN
  eee = cuatroMNEL.ee
  lll = cuatroMNEL.LL
  Frame1.Visible = True
  grdD.TextMatrix(j%, 0) = mmm
  grdD.TextMatrix(j%, 1) = nnn
  grdD.TextMatrix(j%, 2) = eee
  grdD.TextMatrix(j%, 3) = lll
  
 Next j%
 Close
 u = NumReg - 1

 Exit Sub
 
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al abrir el archivo " & NombreArch
 End If


End Sub

Private Sub mc_Click()
frmconductividad.Show
End Sub

Private Sub mconv_Click()
frmconvertidor.Show
End Sub

Private Sub me_Click()
frmETO.Show
End Sub

Private Sub mgs_Click()
frmgenerales.Show
End Sub

Private Sub mguardar_Click()
 On Error GoTo SinArchivo
 ChDir App.Path
 cdCrear.ShowSave
 NombreArch = cdCrear.FileName
 ' Salvar archivo
 Open NombreArch For Random As #1 Len = Len(cuatroMNEL)
 If (LOF(1) <> 0) Then
  Close #1
  Kill NombreArch
  Open NombreArch For Random As #1 Len = Len(cuatroMNEL)
 End If
 For j% = 1 To (grdD.Rows - 1)
 
  valormes(j%) = grdD.TextMatrix(j%, 0)
  valorn(j%) = grdD.TextMatrix(j%, 1)
  valoretr(j%) = grdD.TextMatrix(j%, 2)
  valorln(j%) = grdD.TextMatrix(j%, 3)
  cuatroMNEL.mm = valormes(j%)
  cuatroMNEL.NN = valorn(j%)
  cuatroMNEL.ee = valoretr(j%)
  cuatroMNEL.LL = valorln(j%)
  Put #1, j%, cuatroMNEL
 Next j%
 Close
 Exit Sub
SinArchivo:
 If Err.Number = 32755 Then
  MsgBox "Error desconocido al salvar el archivo " & NombreArch
 End If
 
End Sub

Private Sub mm_Click()
Unload Me
frmGeneral.Show
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
dia = MonthView1.Day
mes = MonthView1.Month
aaa = MonthView1.Year
valorrmes(0) = 0
valorrmes(1) = mes
Select Case mes
Case 1
meess = "Enero"
Case 2
meess = "Febrero"
Case 3
meess = "Marzo"
Case 4
meess = "Abril"
Case 5
meess = "Mayo"
Case 6
meess = "Junio"
Case 7
meess = "Julio"
Case 8
meess = "Agosto"
Case 9
meess = "Setiembre"
Case 10
meess = "Octubre"
Case 11
meess = "Noviembre"
Case 1
meess = "Diciembre"
End Select
grdD.TextMatrix(1, 0) = meess
   Select Case mes
        Case 1
            NN = 31
        Case 2
            NN = 28
        Case 3
            NN = 31
        Case 4
            NN = 30
        Case 5
            NN = 31
        Case 6
            NN = 30
        Case 7
            NN = 31
        Case 8
            NN = 31
        Case 9
            NN = 30
        Case 10
            NN = 31
        Case 11
            NN = 30
        Case 12
            NN = 31
    End Select

gg = NN - dia
grdD.TextMatrix(1, 1) = NN
Combo1.ListIndex = mes - 1
lbletiqueta.Caption = "Inicio del Riego: " + (dia) + " de " + (meess) + " del " + Str(aaa)
Frame1.Visible = True
grdD.SetFocus
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


Private Sub mt_Click()
frmtextura.Show
End Sub
