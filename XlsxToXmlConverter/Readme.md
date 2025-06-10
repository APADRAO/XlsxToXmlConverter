
# Conversor XML ⇄ Excel (.NET 8, C# 12)

Este projeto fornece métodos utilitários para converter arquivos XML em planilhas Excel (XLSX) e vice-versa, utilizando C# 12 e .NET 8. As conversões são realizadas inteiramente em memória, com arquivos trafegando em formato Base64, facilitando integrações com APIs e sistemas que não trabalham diretamente com arquivos físicos.

## Resumo do Código

O código principal está no arquivo `Program.cs` e implementa três métodos:

- **XmlToExcelAsinc**: Converte um arquivo XML (em Base64) para Excel (XLSX, também em Base64). O XML é lido, transformado em uma lista de dicionários (cada um representando uma linha), e então exportado para uma planilha Excel. Os dados são organizados em colunas baseadas nos nomes dos elementos e atributos do XML.
- **FlattenElement**: Método auxiliar recursivo que "achata" a estrutura hierárquica do XML, convertendo elementos e atributos em pares chave-valor para facilitar a exportação para Excel.
- **ExcelToXml**: Converte um arquivo Excel (em Base64) para XML (também em Base64). Lê os dados da planilha, utiliza a primeira linha como cabeçalho e cria elementos XML para cada linha de dados, agrupando todos sob um elemento raiz.

Todos os métodos trabalham apenas com dados em memória e retornam o resultado como string Base64.

## Funcionalidades

- **XmlToExcelAsinc**: Converte um arquivo XML (Base64) em uma planilha Excel (Base64).
- **ExcelToXml**: Converte uma planilha Excel (Base64) em um arquivo XML (Base64).

## Dependências

- [.NET 8](https://dotnet.microsoft.com/download)
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) (para manipulação de arquivos Excel)

## Instalação

Adicione o pacote ClosedXML ao seu projeto.

## Como Funciona

### 1. XmlToExcelAsinc

Converte um arquivo XML (em Base64) para um arquivo Excel (também em Base64).

**Passos:**
1. Validação inicial dos parâmetros.
2. Decodificação do Base64 e leitura do XML.
3. Processamento dos dados com o método `FlattenElement`, que transforma a estrutura do XML em linhas e colunas.
4. Criação do arquivo Excel em memória, preenchendo os dados.
5. Retorno do arquivo Excel convertido em Base64.

### 2. FlattenElement

Transforma um elemento XML (e seus filhos) em uma estrutura plana, adequada para ser representada em linhas e colunas de uma planilha.

- Para cada atributo do elemento, adiciona ao dicionário atual.
- Se o elemento não tem filhos, adiciona seu valor ao dicionário.
- Se tem filhos, verifica se há elementos repetidos (arrays) e processa cada filho recursivamente, criando novas linhas conforme necessário.
- Adiciona cada linha processada à lista de linhas.

### 3. ExcelToXml

Converte um arquivo Excel (em Base64) para um arquivo XML (também em Base64).

**Passos:**
1. Validação inicial dos parâmetros.
2. Decodificação do Base64 e leitura do Excel.
3. Processamento dos dados: lê os cabeçalhos e cria elementos XML para cada linha.
4. Criação do XML agrupando todos os registros sob um elemento raiz.
5. Retorno do arquivo XML convertido em Base64.

## Exemplos de Uso

### Converter XML para Excel

