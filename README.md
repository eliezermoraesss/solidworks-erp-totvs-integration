# Cadastro Automático de Produtos e Estrutura (BOM) no ERP TOTVS direto pelo ambiente do SolidWorks

## Visão Geral

Este script em Python foi desenvolvido para economizar horas de trabalho manual pelo simples clicar de um botão.
Realiza a validação e cadastro de produtos e estruturas (BOM) em um banco de dados Microsoft SQL Server. Ele oferece funcionalidades para verificar a consistência dos dados da BOM, criar ou atualizar estruturas no banco de dados TOTVS® Protheus, e interagir com o usuário por meio de mensagens.

## Requisitos
Certifique-se de que você tenha os seguintes requisitos instalados:

### Importações

- **pyodbc**: Conexão ao SQL Server.
- **pandas**: Manipulação de dados.
- **ctypes**: Interação com bibliotecas compartilhadas.
- **os**: Interação com o sistema operacional.
- **re**: Expressões regulares.
- **datetime**: Manipulação de datas e horários.
- **tkinter**: Ferramentas gráficas para interfaces de usuário.

## Configuração - Parâmetros de Conexão

Configurações para conectar-se ao SQL Server, incluindo servidor, banco de dados, usuário, senha e driver.

```python
server = 'SERVIDOR,PORTA'
database = 'NOME_DO_BANCO'  # Substitua pelo nome do banco de dados
username = 'NOME_DO_USUARIO'
password = 'SENHA'
driver = '{ODBC Driver 17 for SQL Server}'
```
Certifique-se de ter permissões adequadas para acessar o banco de dados especificado.

## Utilização
### Ambiente de Execução:
Certifique-se de que o Python esteja instalado no ambiente.
Execute o script utilizando um ambiente Python compatível.

### Código do Desenho:
O código do desenho deve ser definido como uma variável de ambiente chamada CODIGO_DESENHO.

### Arquivo Excel:
O script espera um arquivo Excel com os dados da BOM no formato adequado.

### Formato do Código Pai:
O código do desenho deve seguir um dos formatos padrão ENAPLIC:

```
C-###-###-###

M-###-###-###

E####-###-###
```
### Execução:
Execute o script para validar e cadastrar os dados da BOM no ERP TOTVS.

## Funções Principais

### validação_formato_codigo_pai

Valida o formato de um código pai usando expressões regulares.

### validacao_formato_codigos_filho

Valida o formato de códigos filhos em um arquivo Excel.

### ler_variavel_ambiente_codigo_desenho

Lê uma variável de ambiente chamada CODIGO_DESENHO.

### obter_caminho_arquivo_excel

Constrói o caminho para um arquivo Excel com base no código do desenho.

### excluir_arquivo_excel_bom

Exclui o arquivo Excel especificado, se existir.

### verificar_codigo_repetido

Verifica códigos duplicados na BOM.

### verificar_cadastro_codigo_filho

Verifica o cadastro dos códigos filhos no banco de dados SQL Server.

### remover_linhas_duplicadas_e_consolidar_quantidade

Consolida quantidades para linhas duplicadas na BOM.

### validar_descricao

Valida descrições na BOM.

### validacao_de_dados_bom

Valida vários aspectos dos dados da BOM e retorna um DataFrame limpo.

### atualizar_campo_revisao_do_codigo_pai

Atualiza o campo de revisão para um código pai no SQL Server.

### verificar_se_existe_estrutura_totvs

Verifica se existe uma estrutura existente no SQL Server para um código pai.

### obter_ultima_pk_tabela_estrutura

Recupera a última chave primária da tabela de estrutura.

### obter_revisao_inicial_codigo_pai

Recupera a revisão inicial para um código pai.

### obter_unidade_medida_codigo_filho

Recupera a unidade de medida para um código filho.

### formatar_data_atual

Formata a data atual.

### verificar_cadastro_codigo_pai

Verifica se um código pai está registrado no SQL Server.

### criar_nova_estrutura_totvs

Cria uma nova estrutura no SQL Server com base nos dados da BOM.

### ask_user_for_action

Pergunta ao usuário se deseja alterar uma estrutura existente.

### alterar_estrutura_existente

Compara os dados da BOM com a estrutura existente e exibe as diferenças.

## Execução Principal

- Lê o código do desenho da variável de ambiente.
- Constrói o caminho para o arquivo Excel.
- Valida o formato do código pai.
- Verifica o cadastro do código pai no SQL Server.
- Valida os dados da BOM, verifica a existência de uma estrutura e trata a criação ou alteração conforme necessário.

## Conclusão

O script oferece uma solução que integra o SolidWorks e o ERP TOTVS, reduzindo em 80% o tempo de cadastro de produtos e estruturas (BOM), economizando horas e horas de trabalho manual do usuário pelo simples clicar de um botão.

## Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests para melhorias ou correções.

## Desenvolvido por:
Eliezer Moraes Silva
eliezer.moraes@outlook.com

## Licença
Este projeto é licenciado sob a Licença MIT.