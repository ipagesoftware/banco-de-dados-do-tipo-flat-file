# Arquivos Texto Randômicos no Visual Basic 6

## Tratamento de arquivos texto randômicos (Random Access Files)

Este material apresenta, de forma prática e didática, o tratamento de **arquivos randômicos no Visual Basic 6 (VB6)**.

O exemplo utiliza um cadastro de produtos em que cada registro possui **100 caracteres**, divididos em campos de tamanho fixo.

Esse tipo de arquivo permite que o programa leia ou altere diretamente um determinado registro sem precisar percorrer todos os registros anteriores.

---

# 1. O que é um arquivo randômico?

Em um arquivo sequencial, normalmente os dados são lidos na ordem em que foram gravados. Para encontrar o registro de número 500, por exemplo, pode ser necessário passar pelos registros anteriores.

Em um arquivo randômico de tamanho fixo, cada registro ocupa sempre a mesma quantidade de bytes. Assim, podemos acessar diretamente o registro desejado.

### Exemplo

Se cada registro possui 100 caracteres:

```text
Registro 1    → 100 caracteres
Registro 2    → 100 caracteres
Registro 3    → 100 caracteres
...
Registro 500  → 100 caracteres
```

No VB6, os principais recursos utilizados são:

- `Open ... For Random` — abre o arquivo para acesso randômico.
- `Len` — informa o tamanho do registro.
- `Get` — lê um registro.
- `Put` — grava ou substitui um registro.
- `LOF` — informa o tamanho do arquivo em bytes.

---

# 2. Estrutura do registro

Vamos utilizar a seguinte estrutura:

| Campo | Tamanho | Posição inicial | Posição final |
|---|---:|---:|---:|
| Nome | 20 | 1 | 20 |
| Código | 5 | 21 | 25 |
| Preço de compra | 15 | 26 | 40 |
| Preço de venda | 10 | 41 | 50 |
| Quantidade | 10 | 51 | 60 |
| Total pago | 20 | 61 | 80 |
| Tabela | 10 | 81 | 90 |
| Último | 10 | 91 | 100 |
| **Total** | **100** | | |

A ideia é que cada campo tenha uma largura previamente definida.

No VB6, podemos representar essa estrutura utilizando `Type` e strings de tamanho fixo:

```vb
Private Type TProduto
    Nome        As String * 20
    Codigo      As String * 5
    PrecoCompra As String * 15
    PrecoVenda  As String * 10
    Quantidade  As String * 10
    TotalPago   As String * 20
    Tabela      As String * 10
    Ultimo      As String * 10
End Type
```

Quando um texto ocupa menos espaço que o campo definido, o VB6 completa o restante com espaços.

Por exemplo:

```vb
Produto.Nome = "CHAPARIA"
```

O campo continuará ocupando 20 caracteres.

---

# 3. Criando um arquivo randômico

O exemplo abaixo cria ou abre o arquivo `produtos.dat` e grava um produto na primeira posição.

```vb
Private Sub GravarProduto()

    Dim Arquivo As Integer
    Dim Produto As TProduto
    Dim NumeroRegistro As Long

    Arquivo = FreeFile

    Open App.Path & "\produtos.dat" _
        For Random As #Arquivo _
        Len = Len(Produto)

    Produto.Nome = "CHAPARIA"
    Produto.Codigo = "0002"
    Produto.PrecoCompra = "6,603"
    Produto.PrecoVenda = "9,60"
    Produto.Quantidade = "1,462"
    Produto.TotalPago = "9,60"
    Produto.Tabela = "0"
    Produto.Ultimo = "8,0000"

    NumeroRegistro = 1

    Put #Arquivo, NumeroRegistro, Produto

    Close #Arquivo

    MsgBox "Produto gravado com sucesso!"

End Sub
```

A instrução:

```vb
Open App.Path & "\produtos.dat" _
    For Random As #Arquivo _
    Len = Len(Produto)
```

informa ao VB6 que estamos trabalhando com registros de tamanho fixo.

A instrução:

```vb
Put #Arquivo, 1, Produto
```

grava o conteúdo de `Produto` no registro número 1.

Se o registro já existir, seu conteúdo será substituído.

---

# 4. Por que usamos `Len = Len(Produto)`?

O VB6 precisa saber qual é o tamanho de cada registro.

Como nosso `Type` possui campos de tamanho fixo, podemos utilizar:

```vb
Len(Produto)
```

Assim, o VB6 conhece o tamanho do registro.

A ideia é:

```text
Arquivo
│
├── Registro 1 → 100 bytes
├── Registro 2 → 100 bytes
├── Registro 3 → 100 bytes
├── Registro 4 → 100 bytes
└── ...
```

Isso é o que permite o acesso direto.

---

# 5. Lendo um registro específico

Uma das principais vantagens do acesso randômico é poder buscar diretamente um registro.

```vb
Private Sub LerProduto(ByVal Numero As Long)

    Dim Arquivo As Integer
    Dim Produto As TProduto

    Arquivo = FreeFile

    Open App.Path & "\produtos.dat" _
        For Random As #Arquivo _
        Len = Len(Produto)

    Get #Arquivo, Numero, Produto

    Close #Arquivo

    Debug.Print "Nome: "; Trim(Produto.Nome)
    Debug.Print "Código: "; Trim(Produto.Codigo)
    Debug.Print "Compra: "; Trim(Produto.PrecoCompra)
    Debug.Print "Venda: "; Trim(Produto.PrecoVenda)

End Sub
```

Para ler o registro número 10:

```vb
Call LerProduto(10)
```

O VB6 calcula a posição do registro com base no tamanho fixo definido para o arquivo.

Não precisamos ler os registros 1 a 9.

---

# 6. Alterando um registro existente

Podemos ler um registro, modificar seus dados e gravá-lo novamente na mesma posição.

```vb
Private Sub AlterarProduto(ByVal Numero As Long)

    Dim Arquivo As Integer
    Dim Produto As TProduto
    Dim TotalRegistros As Long

    Arquivo = FreeFile

    Open App.Path & "\produtos.dat" _
        For Random As #Arquivo _
        Len = Len(Produto)

    TotalRegistros = LOF(Arquivo) \ Len(Produto)

    If Numero < 1 Or Numero > TotalRegistros Then
        Close #Arquivo
        MsgBox "Registro não encontrado!"
        Exit Sub
    End If

    Get #Arquivo, Numero, Produto

    Produto.PrecoVenda = "12,50"

    Put #Arquivo, Numero, Produto

    Close #Arquivo

    MsgBox "Produto alterado com sucesso!"

End Sub
```

A sequência é:

1. Descobrir quantos registros existem.
2. Verificar se o registro solicitado existe.
3. Ler o registro com `Get`.
4. Alterar os campos necessários.
5. Gravar novamente com `Put`.

Como o registro continua ocupando o mesmo espaço, os outros registros não são deslocados.

---

# 7. Incluindo um novo registro

Para incluir um novo produto no final do arquivo, podemos descobrir qual será o próximo número de registro.

```vb
Private Sub IncluirProduto()

    Dim Arquivo As Integer
    Dim Produto As TProduto
    Dim NovoRegistro As Long

    Arquivo = FreeFile

    Open App.Path & "\produtos.dat" _
        For Random As #Arquivo _
        Len = Len(Produto)

    NovoRegistro = (LOF(Arquivo) \ Len(Produto)) + 1

    Produto.Nome = "TUBO"
    Produto.Codigo = "0004"
    Produto.PrecoCompra = "6,432"
    Produto.PrecoVenda = "1,158"
    Produto.Quantidade = "179,20"
    Produto.TotalPago = "1.158,20"
    Produto.Tabela = "0"
    Produto.Ultimo = "5,0000"

    Put #Arquivo, NovoRegistro, Produto

    Close #Arquivo

    MsgBox "Produto incluído no registro " & NovoRegistro

End Sub
```

A expressão:

```vb
NovoRegistro = (LOF(Arquivo) \ Len(Produto)) + 1
```

é muito importante.

### `LOF(Arquivo)`

Retorna o tamanho do arquivo em bytes.

### `Len(Produto)`

Retorna o tamanho de um registro.

### `\`

Realiza uma divisão inteira no VB6.

### `+ 1`

Obtém o próximo registro.

Por exemplo, se o arquivo possui 24 registros de 100 bytes:

```text
LOF = 2400
Len = 100

2400 \ 100 = 24

24 + 1 = 25
```

Portanto, o próximo registro será o número 25.

---

# 8. Percorrendo todos os registros

Também podemos percorrer todos os registros de um arquivo randômico.

```vb
Private Sub ListarProdutos()

    Dim Arquivo As Integer
    Dim Produto As TProduto
    Dim Numero As Long
    Dim TotalRegistros As Long

    Arquivo = FreeFile

    Open App.Path & "\produtos.dat" _
        For Random As #Arquivo _
        Len = Len(Produto)

    TotalRegistros = LOF(Arquivo) \ Len(Produto)

    For Numero = 1 To TotalRegistros

        Get #Arquivo, Numero, Produto

        Debug.Print Numero
        Debug.Print Trim(Produto.Nome)
        Debug.Print Trim(Produto.Codigo)
        Debug.Print Trim(Produto.PrecoCompra)
        Debug.Print Trim(Produto.PrecoVenda)

    Next Numero

    Close #Arquivo

End Sub
```

Esse procedimento pode ser utilizado para carregar os dados em controles como:

- `ListBox`
- `ListView`
- `MSFlexGrid`
- `DataGrid`
- outros controles de interface do VB6

---

# 9. Como o registro é localizado?

Se cada registro possui 100 bytes, podemos visualizar o arquivo desta maneira:

| Registro | Deslocamento inicial |
|---:|---:|
| 1 | 0 |
| 2 | 100 |
| 3 | 200 |
| 4 | 300 |
| 5 | 400 |
| 10 | 900 |
| 100 | 9.900 |
| 500 | 49.900 |

A fórmula é:

```text
Deslocamento = (Número do registro - 1) × Tamanho do registro
```

Para o registro 500:

```text
(500 - 1) × 100

499 × 100

49.900
```

No código, entretanto, não precisamos fazer esse cálculo manualmente.

Basta:

```vb
Get #Arquivo, 500, Produto
```

O VB6 se encarrega do posicionamento.

---

# 10. Arquivo sequencial × arquivo randômico

| Característica | Sequencial | Randômico |
|---|---|---|
| Organização | Geralmente linha após linha | Registros de tamanho fixo |
| Acesso direto pelo registro | Não é automático | Sim |
| Tamanho dos registros | Pode variar | Fixo |
| Alteração no meio do arquivo | Pode exigir reconstrução | Pode substituir diretamente |
| Uso comum | Configuração, importação, exportação | Cadastros de registros fixos |

No VB6, os arquivos sequenciais normalmente utilizam:

```vb
Open ... For Input
Open ... For Output
Open ... For Append
```

Enquanto os arquivos randômicos utilizam:

```vb
Open ... For Random
```

---

# 11. `Get` e `Put`

Essas duas instruções são fundamentais.

## `Get`

Lê um registro:

```vb
Get #Arquivo, NumeroRegistro, Produto
```

Podemos pensar:

```text
ARQUIVO
   ↓
Get
   ↓
Produto
```

## `Put`

Grava um registro:

```vb
Put #Arquivo, NumeroRegistro, Produto
```

Podemos pensar:

```text
Produto
   ↓
Put
   ↓
ARQUIVO
```

---

# 12. Um exemplo completo

Abaixo temos uma pequena demonstração que inclui um produto e depois o lê.

```vb
Private Type TProduto

    Nome        As String * 20
    Codigo      As String * 5
    PrecoCompra As String * 15
    PrecoVenda  As String * 10
    Quantidade  As String * 10
    TotalPago   As String * 20
    Tabela      As String * 10
    Ultimo      As String * 10

End Type


Private Sub Form_Load()

    Dim Arquivo As Integer
    Dim Produto As TProduto

    Arquivo = FreeFile

    Open App.Path & "\produtos.dat" _
        For Random As #Arquivo _
        Len = Len(Produto)

    '--------------------------------------------------
    ' GRAVA
    '--------------------------------------------------

    Produto.Nome = "CHAPARIA"
    Produto.Codigo = "0002"
    Produto.PrecoCompra = "6,603"
    Produto.PrecoVenda = "9,60"
    Produto.Quantidade = "1,462"
    Produto.TotalPago = "9,60"
    Produto.Tabela = "0"
    Produto.Ultimo = "8,0000"

    Put #Arquivo, 1, Produto

    '--------------------------------------------------
    ' LIMPA A VARIÁVEL
    '--------------------------------------------------

    Produto = Empty

    '--------------------------------------------------
    ' LÊ
    '--------------------------------------------------

    Get #Arquivo, 1, Produto

    Close #Arquivo

    Debug.Print "Nome: " & Trim(Produto.Nome)
    Debug.Print "Código: " & Trim(Produto.Codigo)
    Debug.Print "Preço de compra: " & Trim(Produto.PrecoCompra)
    Debug.Print "Preço de venda: " & Trim(Produto.PrecoVenda)

End Sub
```

---

# 13. Atenção ao `Trim`

Como estamos utilizando campos de tamanho fixo:

```vb
Nome As String * 20
```

se armazenarmos:

```text
CHAPARIA
```

o campo ocupará os 20 caracteres.

Na prática teremos algo semelhante a:

```text
CHAPARIA            |
```

Por isso, quando queremos exibir o conteúdo sem os espaços adicionais, utilizamos:

```vb
Trim(Produto.Nome)
```

O `Trim` remove os espaços do início e do final da string.

Exemplo:

```vb
Debug.Print Trim(Produto.Nome)
```

em vez de:

```vb
Debug.Print Produto.Nome
```

---

# 14. Uma observação importante sobre arquivos "texto"

Embora o exemplo utilize campos `String`, um arquivo aberto com:

```vb
For Random
```

não deve ser confundido com um arquivo texto sequencial tradicional.

Um arquivo sequencial normalmente possui estruturas como:

```text
CHAPARIA|0002|6,603|9,60
TUBO|0004|6,432|1,158
```

ou:

```text
CHAPARIA;0002;6,603;9,60
TUBO;0004;6,432;1,158
```

Já o arquivo randômico trabalha com posições fixas:

```text
[20 caracteres][5][15][10][10][20][10][10]
```

Isso permite localizar diretamente um registro.

---

# 15. Atenção aos tipos de dados

Neste exemplo, utilizamos `String` de tamanho fixo:

```vb
Nome As String * 20
```

Isso faz sentido quando queremos reproduzir um arquivo de largura fixa composto por caracteres.

Entretanto, um `Type` do VB6 também pode conter outros tipos:

```vb
Private Type TProduto

    Codigo As Long
    Preco  As Currency
    Nome   As String * 30

End Type
```

Nesse caso, os dados numéricos podem ser gravados em representação binária.

Portanto, se o objetivo for criar um arquivo realmente textual e compatível com outro sistema, é importante definir cuidadosamente o formato de cada campo.

---

# 16. Exercícios para os alunos

## Exercício 1 — Criar um cadastro

Crie um programa VB6 capaz de:

1. Criar o arquivo `produtos.dat`.
2. Cadastrar produtos.
3. Gravar cada produto em um registro.
4. Informar o número do registro criado.

---

## Exercício 2 — Consultar produto

Crie uma tela com:

```text
Número do registro: [     ]

[ CONSULTAR ]
```

Ao clicar no botão, o programa deverá utilizar:

```vb
Get #Arquivo, NumeroRegistro, Produto
```

e apresentar os dados na tela.

---

## Exercício 3 — Alterar produto

Crie uma tela para informar o número do registro e permitir a alteração do preço de venda.

O programa deverá:

```text
Ler
 ↓
Alterar
 ↓
Gravar novamente
```

---

## Exercício 4 — Listar produtos

Percorra todos os registros e apresente:

```text
Código     Nome                    Preço
0001       BLOCO DE ALUMÍNIO      6,2882
0002       CHAPARIA               6,6063
0003       PERFIL                 10,6841
```

---

## Exercício 5 — Localização direta

Faça o aluno consultar diretamente o registro 10.

Pergunta:

> Quantos registros anteriores precisam ser lidos para chegar ao registro 10?

Resposta esperada:

**Nenhum.**

Essa é uma das principais características do acesso randômico.

---

# 17. Desafio final

Crie um pequeno sistema de cadastro de produtos em VB6 utilizando arquivo randômico.

O sistema deverá possuir:

```text
┌───────────────────────────────────┐
│       CADASTRO DE PRODUTOS        │
├───────────────────────────────────┤
│ Código:       [              ]    │
│ Nome:         [              ]    │
│ Preço compra: [              ]    │
│ Preço venda:  [              ]    │
│ Quantidade:   [              ]    │
├───────────────────────────────────┤
│ [NOVO] [CONSULTAR] [ALTERAR]      │
│ [EXCLUIR] [PRIMEIRO] [PRÓXIMO]    │
└───────────────────────────────────┘
```

O programa deverá utilizar:

```vb
Open
Get
Put
Close
Len
LOF
```

E deverá permitir:

- inclusão;
- consulta;
- alteração;
- navegação;
- listagem;
- validação do número do registro.

---

# 18. Conceito fundamental

O conceito mais importante que o aluno deve compreender é:

> **No arquivo randômico, todos os registros possuem tamanho conhecido e fixo.**

Por exemplo:

```text
Registro 1 → 100 bytes
Registro 2 → 100 bytes
Registro 3 → 100 bytes
Registro 4 → 100 bytes
Registro 5 → 100 bytes
```

Consequentemente, o programa consegue determinar diretamente onde está determinado registro.

```text
             ARQUIVO

┌────────────┬────────────┬────────────┬────────────┐
│ Registro 1 │ Registro 2 │ Registro 3 │ Registro 4 │
│ 100 bytes  │ 100 bytes  │ 100 bytes  │ 100 bytes  │
└────────────┴────────────┴────────────┴────────────┘
      ↑             ↑
      │             │
     Get           Put
```

---

# 19. Resumo dos principais comandos

| Comando | Função |
|---|---|
| `FreeFile` | Obtém um número de arquivo disponível |
| `Open` | Abre o arquivo |
| `For Random` | Define acesso randômico |
| `Len` | Define/informa o tamanho do registro |
| `Get` | Lê um registro |
| `Put` | Grava um registro |
| `LOF` | Informa o tamanho do arquivo |
| `Close` | Fecha o arquivo |
| `Trim` | Remove espaços das extremidades |

---

# 20. Conclusão

O acesso randômico é uma técnica clássica de programação que foi muito utilizada em sistemas desenvolvidos em **Visual Basic 6, Clipper, Delphi, Pascal, COBOL e outras tecnologias de sua época**.

Seu grande diferencial é permitir que registros de tamanho fixo sejam acessados diretamente.

No VB6, os elementos fundamentais são:

```text
Type
  ↓
String * tamanho
  ↓
Open ... For Random
  ↓
Get / Put
  ↓
Len / LOF
```

A ideia pode ser resumida assim:

```text
                 ARQUIVO RANDÔMICO

        ┌──────────────────────────────┐
        │ Registro 1 - 100 bytes      │
        ├──────────────────────────────┤
        │ Registro 2 - 100 bytes      │
        ├──────────────────────────────┤
        │ Registro 3 - 100 bytes      │
        ├──────────────────────────────┤
        │ Registro 4 - 100 bytes      │
        ├──────────────────────────────┤
        │ ...                          │
        └──────────────────────────────┘
                    ↑
                    │
              acesso direto
                    │
                  Get/Put
```

O mais importante para o aluno não é apenas decorar `Get` e `Put`, mas entender **por que o tamanho fixo do registro permite o acesso direto aos dados**.

---

## Checklist de aprendizagem

- [ ] Entendi o que é um arquivo randômico.
- [ ] Entendi a diferença entre arquivo sequencial e randômico.
- [ ] Sei utilizar `Type`.
- [ ] Sei utilizar `String * N`.
- [ ] Sei abrir um arquivo com `For Random`.
- [ ] Sei utilizar `Get`.
- [ ] Sei utilizar `Put`.
- [ ] Sei utilizar `LOF`.
- [ ] Sei utilizar `Len`.
- [ ] Sei calcular o número do próximo registro.
- [ ] Sei consultar diretamente um registro.
- [ ] Sei alterar um registro existente.
- [ ] Entendi por que os registros precisam ter tamanho fixo.

---

**Material didático — Visual Basic 6**

**Tema:** Arquivos Randômicos / Random Access Files

**Objetivo:** compreender como armazenar, consultar, alterar e percorrer registros de tamanho fixo utilizando arquivos randômicos no Visual Basic 6.
