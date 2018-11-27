# Vb6-Style-Guide
Uma abordagem prática para escrita de codigos VB6

## Conteúdo

1. [Convenção Capitalização](#convenção-capitalização)
1. [Convenção Nomeação](#convenção-nomeação)
1. [Nomes de Classes, Estruturas e Interfaces](#nomes-de-classes-estruturas-e-interfaces)
1. [Comentários](#comentários)
1. [Referências](#referências)

## Convenção Capitalização

Existem basicamente duas formas de capitalizações, PascalCasing e camelCasing

PascalCasing: Primeira letra de cada palavra em maiusculo
camelCasing: Primeira Letrar minusculo e o restante maiusculo. O camelCasing é basicamente utilizado para parametros e variáveis de scopo local.

Exemplos de uso:
- Interface
```vb
'ruim
Option Explicit
Implements Pessoa { ... }

'bom 
Implements IPessoa

Option Explicit
{ ... }
```

- Métodos
```vb
'ruim 
Option Explicit
public function recuperarPorCodigo() as long { ... }
public function recuperar_por_codigo() as long { ... }

'bom 
Option Explicit
public function RecuperarPorCodigo() as long { ... }
```

- Propriedades 
```vb
'ruim
Option Explicit
public codigoInterno as Integer
public codigo_interno as Integer

'bom
Option Explicit
Private strNome As String
Public Property Get Nome() As String
    Nome = strNome
End Property

Public Property Let Nome(nome_ As String)
    strNome = nome_
End Property

```

- Variáveis locais 
```vb
'ruim
Option Explicit
Private sub CalcularJuros()
  Dim TaxaJuros as Currency
End Sub

Public sub CalcularJuros()
  Dim TaxaJuros as Currency
End Sub


'bom
Option Explicit
Private sub CalcularJuros()
  Dim taxaJuros as Currency
End Sub

Public sub CalcularJuros()
  Dim taxaJuros as Currency
End Sub

```

- Parâmetros 
```vb
'ruim
Option Explicit
public function RecuperarPorCodigo(CodigoInterno as long) as long {...}
public function RecuperarPorCodigo(codigo_interno as long) as long {...}

'bom
Option Explicit
public function RecuperarPorCodigo(codigoInterno_ as long) as long {...}
```

## Convenção Nomeação

- Escolha termos de fácil leitura e boa coerência.

```vb
'ruim 
MensalExtrato

'bom
ExtratoMensal
```

- Prefira Legibilidade ao invés de abreviações
```vb
'ruim 
Dim LimSqDia as Integer
LimSqDia = 2

'bom 
Dim LimiteSaqueDiario as Integer
LimiteSaqueDiario = 2
```

- Utilizar PARCIALMENTE Hungarian Notation
Obs: Hungarian notation é a técnica de adicionar prefixo aos nomes e, em alguns momentos, será necessária no vb6.

```vb
Option Explicit
'Utilizar esta técnica para:
' - Componentes (TextBox, ComboBox, Grid);
'ruim
Dim TNome as TextBox
Dim GProdutos as MSFlexGrid

'bom
Dim txtNome as TextBox
Dim grdProdutos as MSFLexGrid

' - Variável privada que compõe uma propriedade;
'ruim
Private nome as String
Private DataEmissao as Date

'bom 
Private strNome         as String
Private datDataEmissao  as Date

' - Estruturas do tipo Type e Enum;
'ruim
Public Enum Operacao
    Gravar = 1
    Excluir = 2
End Enum

'bom
Public Enum enuOperacao
    adGravar = 1
    adExcluir = 2
End Enum
```

- Evite o uso de palavras reservadas da Linguagem
```vb
' ruim
Dim Select as String;
```
## Nomes de Classes, Estruturas e Interfaces

- Nomes de Classes e Estruturas utilize substantivos ou frases nominais 
- Nomes de Interfaces dê preferencia para o uso de adjetivos e ocasionalmente substantivos ou frases nominais 
- Utilize o prefixo "I" para nomes de interfaces
- Utilize o prefixo "model" para nomes de Modelos (que representam entidades do banco de dados);
- Utilize o prefixo "cls" para classes auxiliares (exemplo: clsUtil);
- Utilize o prefixo "frm" para Formulários;
- Utilize o prefixo "filtro" para classes que representem filtros de uma busca;
- Nomes de Enuns
	* Use nomes singulares

- Nomes de Métodos
	* Utilize Verbos ou Frases verbais

```vb
Option Explicit
Public Function ObterSaldo() as Double {...}
Public Function Salvar() as Boolean { ... }
```

- Nomes de Propriedades
	* Utilize substantivos, frases nominais ou adjetivos

## Comentários
- Comentário de uma linha
```vb
Public Function ValidarDescricaoProdutoImportado(descricao_ as String) as String
{
  ' Foi necessário remover a quebra de linha, para que não ocorra erro ao Exportar Carga de produtos para o PDV 	
  ValidarDescricaoProdutoImportado = Replace(descricao_,"\n","")
}
```

## Referências
* https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation
