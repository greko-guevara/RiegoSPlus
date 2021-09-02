Attribute VB_Name = "Module1"
Global PARRa As Double
Global eggg As Double
Global sss11 As Double


Public NumReg As Integer
Public NombreArch As String

'Guardar grid prueba de avance
Type ParesValore
 T As Single
 L As Single
End Type

'Guardar grid prueba de surcoinfiltrometros
Type ParesValores11
 qs As Single
 tT As Single
 kkqe As Double
 kkef As Double
 kkln As Double
 kkw As Double
 kkl As Double
 kktext1 As Double
 kktext2 As Double
End Type

'Guardar grid prueba de campo en melga
Type ParesValores2
 d As Single
 ta As Single
 tr As Single
 kln As Double
 ka As Double
 kb As Double
End Type

Type ParesValores
 T As Single
 L As Single
 xx1 As Double
 xx2 As Double
End Type

Type caudal
q As Double
xx3 As Double
End Type

Type ParesValores1
 tT As Single
 ntp As Double
 nar As Double
 nec As Double
 numcol As Double
 numfila As Double
End Type

'Guardar grid calendario de riego
Type paresvalores3
 mm As String * 15
 dddia As Single
 mmmes As Single
 aaaño As Single
 NN As Single
 ee As Single
 LL As Single
End Type

Public Parestl As ParesValores
Public Paresqst As ParesValores1
Public cuatroMNEL As paresvalores3
Public caudales As caudal
Public Paresqst1 As ParesValores11
Public TrioDTaTR As ParesValores2
Public Pares As ParesValore
