VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGeneral 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Riegos Plus"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   ForeColor       =   &H00000000&
   Icon            =   "frmprincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture1 
      Height          =   7575
      Left            =   1920
      Picture         =   "frmprincipal.frx":0CCA
      ScaleHeight     =   7515
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Salir"
      Height          =   795
      Left            =   10200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmprincipal.frx":CC37
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7815
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
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
   Begin VB.Menu mgenerales 
      Caption         =   "Generales"
      Begin VB.Menu mpeda 
         Caption         =   "Parámetros edafológicos"
      End
      Begin VB.Menu mtex 
         Caption         =   "Textura"
      End
      Begin VB.Menu mconhidr 
         Caption         =   "Conductividad hidráulica"
      End
      Begin VB.Menu mraya 
         Caption         =   "-"
      End
      Begin VB.Menu metor 
         Caption         =   "Determinación de ETo y ETr"
      End
      Begin VB.Menu dfd 
         Caption         =   "-"
      End
      Begin VB.Menu mcuuuuuu 
         Caption         =   "Convertidor de unidades"
      End
   End
   Begin VB.Menu mda 
      Caption         =   "Diseño agrónomicos"
      Begin VB.Menu ty66 
         Caption         =   "Riego por aspersión"
      End
      Begin VB.Menu zaa 
         Caption         =   "Riego por goteo"
      End
      Begin VB.Menu m8882 
         Caption         =   "Riego por microaspersión"
      End
      Begin VB.Menu llll 
         Caption         =   "-"
      End
      Begin VB.Menu ms 
         Caption         =   "Surcos"
      End
      Begin VB.Menu mmp 
         Caption         =   "Melgas pendiente"
      End
      Begin VB.Menu msp 
         Caption         =   "Melgas sin pendiente"
      End
      Begin VB.Menu ma 
         Caption         =   "Arroceras"
      End
      Begin VB.Menu ee 
         Caption         =   "-"
      End
      Begin VB.Menu mpo 
         Caption         =   "Pozas"
      End
      Begin VB.Menu msu 
         Caption         =   "Subirrigación"
      End
      Begin VB.Menu jj 
         Caption         =   "-"
      End
      Begin VB.Menu mferti 
         Caption         =   "Fertirrigación"
      End
      Begin VB.Menu mnmnmn 
         Caption         =   "-"
      End
      Begin VB.Menu mcalendario 
         Caption         =   "Calendario de riego"
      End
   End
   Begin VB.Menu mmeeeccdd 
      Caption         =   "Evaluación de sistemas"
      Begin VB.Menu maspeva 
         Caption         =   " Riego por Aspersión"
      End
      Begin VB.Menu mgoteoaseva 
         Caption         =   " Riego por Goteo"
      End
      Begin VB.Menu rrr 
         Caption         =   "-"
      End
      Begin VB.Menu mpa 
         Caption         =   "Prueba de avance"
      End
      Begin VB.Menu mpar 
         Caption         =   "Prueba de avance recesión"
      End
      Begin VB.Menu mpsi 
         Caption         =   "Surcos infiltrómetros"
      End
   End
   Begin VB.Menu cvb 
      Caption         =   "Hidráulica de tuberias"
      Begin VB.Menu fsdfsdfsdf 
         Caption         =   "Diseño del  lateral"
      End
      Begin VB.Menu fsdfksdfkldskfñsdk 
         Caption         =   "Diseño de la principal"
      End
      Begin VB.Menu p33oo 
         Caption         =   "Combinación de Diámetros"
      End
      Begin VB.Menu q1q1 
         Caption         =   "-"
      End
      Begin VB.Menu ljlñ 
         Caption         =   "Presiones en el lateral"
      End
      Begin VB.Menu nr 
         Caption         =   "-"
      End
      Begin VB.Menu bbbbbbbb 
         Caption         =   "Selección de bomba"
      End
      Begin VB.Menu dgdfgdgdfgdfgdfgdfgdf 
         Caption         =   "-"
      End
      Begin VB.Menu klklklkllklklkl 
         Caption         =   "Costos bombeo vrs costos red"
      End
   End
   Begin VB.Menu moc 
      Caption         =   "Hidráulica de canales"
      Begin VB.Menu mhccccc 
         Caption         =   "Cálculo en flujo uniforme"
      End
      Begin VB.Menu minfcan 
         Caption         =   "Infiltración en canales"
      End
   End
   Begin VB.Menu macerca 
      Caption         =   "Acerca de:"
   End
   Begin VB.Menu msalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bbbbbbbb_Click()
frmbomba.Show
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub fsdfksdfkldskfñsdk_Click()
frmHprincipal.Show
Unload Me

End Sub

Private Sub fsdfsdfsdf_Click()
FrmHLaterales.Show
Unload Me
End Sub




Private Sub klklklkllklklkl_Click()
frmmetodocostos.Show
End Sub

Private Sub ljlñ_Click()
frmPresiones.Show
Unload Me
End Sub

Private Sub m8882_Click()
frmDAMicro.Show
Unload Me
End Sub

Private Sub ma_Click()
frmarroceras.Show
Unload Me
End Sub

Private Sub macerca_Click()
frmAbout.Show


End Sub

Private Sub maspeva_Click()
frmevalucion.Show
Unload Me
End Sub

Private Sub mcalendario_Click()
frmCalendario.Show
Unload Me
End Sub

Private Sub mconhidr_Click()
frmconductividad.Show

End Sub

Private Sub mcuuuuuu_Click()
frmconvertidor.Show

End Sub

Private Sub metor_Click()
frmETO.Show

End Sub

Private Sub mferti_Click()
frmfertirrigacion.Show
End Sub

Private Sub mgoteoaseva_Click()
frmEVAGOTERO.Show
Unload Me
End Sub

Private Sub mhccccc_Click()
Frmhidraulica.Show
Unload Me
End Sub

Private Sub minfcan_Click()
frminfiltracioncanales.Show
Unload Me
End Sub

Private Sub mmp_Click()
frmmelgaspendiente.Show
Unload Me
End Sub

Private Sub mpa_Click()
frmpruebaavance.Show
Unload Me
End Sub

Private Sub mpar_Click()
frmpruebaavancerecesion.Show
Unload Me
End Sub

Private Sub mpeda_Click()
frmgenerales.Show
End Sub

Private Sub mpo_Click()
frmpozas.Show
Unload Me
End Sub

Private Sub mpsi_Click()
frmsurcosinfiltrometros.Show
Unload Me
End Sub

Private Sub ms_Click()
frmriegosurcos.Show
Unload Me
End Sub

Private Sub msalir_Click()
End
End Sub

Private Sub msp_Click()
frmmelgassinpendiente.Show
Unload Me
End Sub

Private Sub msu_Click()
frmsubirrigacion.Show
Unload Me
End Sub

Private Sub mtex_Click()
frmtextura.Show

End Sub

Private Sub p33oo_Click()
frmcombDia.Show
Unload Me
End Sub

Private Sub ty66_Click()
frmDAaspersion.Show
Unload Me
End Sub

Private Sub zaa_Click()
frmDAgoteo.Show
Unload Me
End Sub
