Attribute VB_Name = "M�dulo1"
Sub constantes()
Const a1 As String = "A1"
Const a2 As String = "A2"
Dim nome As String
Dim numero As Integer

nome = InputBox("Digite seu nome")
numero = InputBox("Digite seu n�mero")

Range(a1).Value = nome

If (numero Mod 2 = 0) Then
Range(a2).Value = "Este n�mero � Par"
Else
Range(a2).Value = "Este n�mero � �mpar"
End If
End Sub
Sub mediaescolar()

Const media_aprovacao As Double = 7
'Para notas maiores ou iguais a 7-> Aprovado
'Para notas menores ou iguais a 4-> Reprovado
'Para o restante recupera��o
Dim notas As Double
nota = InputBox("Digite a nota do aluno")
If (nota > 10 Or nota < 0) Then
MsgBox "Nota Inv�lida"
Else

 If (nota >= media_aprovacao) Then
 MsgBox "Aprovado"
 ElseIf (nota <= 4) Then
 MsgBox "Reprovado"
 Else
 MsgBox "Recupera��o"
 End If
 End If
 
End Sub
