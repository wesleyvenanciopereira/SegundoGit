Attribute VB_Name = "M�dulo1"
Sub estrutura()
'para declarar vari�vel no VBA usamos o comando DIM
Dim produto As String
Dim preco As Double
Dim desconto As Double
Dim precofinal As Double

'vamos utilizar a caixa de entrada inputbox para as vari�veis
produto = InputBox("Digite o nome do produto", "Produto")
preco = InputBox("Digite o pre�o do produto", "Pre�o")
desconto = InputBox("Digite o desconto", "Desconto")
precofinal = preco - preco * desconto

Range("A1").Value = produto
Range("A2").Value = preco
Range("A3").Value = desconto
Range("A4").Value = precofinal

End Sub
