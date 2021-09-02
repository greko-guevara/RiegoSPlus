VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Porcentaje de área  regada"
   ClientHeight    =   3015
   ClientLeft      =   150
   ClientTop       =   3480
   ClientWidth     =   3885
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "&Regresar"
      Height          =   735
      Left            =   2400
      Picture         =   "Dialog.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Si su lámina la cálculos como la relación entre CC, PMP ingrese un PAR estimado "
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Si ingresa una lámina diaria el PAR debe ser de una 100%"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()
Unload Me
End Sub
