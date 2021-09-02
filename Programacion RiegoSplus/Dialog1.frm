VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KC de cultivo"
   ClientHeight    =   3600
   ClientLeft      =   2010
   ClientTop       =   1665
   ClientWidth     =   7245
   Icon            =   "Dialog1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Dialog1.frx":0CCA
      Left            =   120
      List            =   "Dialog1.frx":0D5B
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "Dialog1.frx":1020
      Left            =   2520
      List            =   "Dialog1.frx":10B1
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.ListBox List3 
      Height          =   2400
      ItemData        =   "Dialog1.frx":136F
      Left            =   4920
      List            =   "Dialog1.frx":1400
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio de Temporada"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Mediados de Temporada"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Fin de Temporada"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OKButton_Click()
Unload Me
End Sub
