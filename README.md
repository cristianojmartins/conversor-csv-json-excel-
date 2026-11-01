# Conversor CSV, JSON e Excel

Aplicação desenvolvida em Python para converter arquivos CSV e JSON
entre si e para planilhas Excel (.xlsx), com validação de dados
e testes automatizados.

O projeto faz parte de um portfólio prático de programação,
com foco em tratamento, organização e conversão de dados.

## Funcionalidades

- **CSV → JSON** — Conversão com validação completa
- **CSV → Excel (.xlsx)** — Planilha formatada com cabeçalhos estilizados
- **JSON → CSV** — Conversão de lista de objetos para CSV
- **JSON → Excel (.xlsx)** — Planilha formatada a partir de JSON
- Interface web para enviar arquivos pelo navegador
- Execução também pelo terminal (CLI)
- Preservação de caracteres acentuados e Unicode
- Preservação de zeros à esquerda e códigos identificadores
- Não converte automaticamente campos numéricos (mantém como string)
- Validação de cabeçalhos duplicados, campos ausentes e extras
- Aceita CSV com delimitadores comuns (`,`, `;`, `\t`, `|`)
- Trata objetos e listas aninhados no JSON (serializa como string JSON)
- Limite de upload de 5 MB na aplicação web
- Proteção do nome do arquivo enviado (nomes seguros)
- Testes automatizados para conversões e erros

## Tecnologias

- Python 3
- Flask
- openpyxl
- unittest
- HTML e CSS na interface web

## Requisitos

Python 3 instalado e acesso ao terminal.

Instale as dependências:

```bash
python -m pip install flask openpyxl
```

## Executar a versão web

Entre na pasta do projeto:

```bash
cd ~/freelancer_dados/conversor_csv_json
```

Inicie a aplicação:

```bash
python app_web.py
```

No navegador do mesmo dispositivo, abra:

http://127.0.0.1:8000

Selecione um arquivo CSV ou JSON e escolha uma das quatro conversões:
- CSV → JSON
- CSV → Excel (.xlsx)
- JSON → CSV
- JSON → Excel (.xlsx)

Depois, envie o arquivo para iniciar a conversão.

Para encerrar a aplicação, volte ao terminal e pressione
Ctrl+C.

## Executar pelo terminal (CLI)

O conversor aceita quatro modos de operação:

```bash
# CSV para JSON
python conversor.py csv2json entrada.csv saida.json

# CSV para Excel
python conversor.py csv2xlsx entrada.csv saida.xlsx

# JSON para CSV
python conversor.py json2csv entrada.json saida.csv

# JSON para Excel
python conversor.py json2xlsx entrada.json saida.xlsx
```

O primeiro argumento é o modo de conversão.
O segundo argumento é o arquivo de entrada.
O terceiro argumento é o arquivo de saída.

As extensões dos arquivos são validadas automaticamente.

## Executar os testes

Execute todos os testes automatizados:

```bash
python -m unittest discover -v
```

Os testes verificam:
- Conversões válidas nos 4 modos
- Dados com acentos e caracteres Unicode
- Preservação de zeros à esquerda
- Não conversão automática de números
- Detecção de delimitadores CSV (`,`, `;`, `\t`, `|`)
- Objetos e listas aninhados no JSON
- Cabeçalhos duplicados
- Campos ausentes ou extras
- Arquivos vazios ou sem cabeçalho
- JSON inválido (não é lista, lista vazia, itens não-objetos)
- Extensões de arquivo inválidas
- Modo de conversão inválido
- Preservação de arquivo de saída em caso de erro
- Limite de upload de 5 MB (web)
- Nome seguro de arquivo (web)
- Tratamento de erros internos sem exposição de detalhes (web)

## Arquivos de exemplo e teste

- `clientes.csv`: exemplo de entrada CSV
- `clientes.json`: exemplo de entrada JSON
- `incompleto.csv`: exemplo de campos ausentes
- `extras.csv`: exemplo de campos extras
- `duplicado.csv`: exemplo de cabeçalhos duplicados
- `idade_invalida.csv`: exemplo de valor não numérico em campo idade
- `vazio.csv`: exemplo de arquivo vazio
- `test_conversor.py`: testes do conversor de terminal
- `test_app_web.py`: testes da aplicação web

## Segurança e limitações

A aplicação foi configurada para testes locais. Não a exponha
diretamente à internet sem uma revisão adicional de segurança
e uma configuração apropriada para produção.

O limite de upload da aplicação web é de 5 MB.

## Comportamento importante

### Preservação de dados
- Todos os valores do CSV são mantidos como **strings** no JSON
  (não há conversão automática para int, float, etc.)
- Zeros à esquerda em códigos identificadores são preservados
- Caracteres acentuados e Unicode são preservados em todas as conversões

### Objetos e listas aninhados no JSON
Ao converter JSON → CSV ou JSON → Excel, objetos e listas aninhados
são serializados como strings JSON para não perder informações:

```json
{"nome": "Ana", "endereco": {"rua": "Rua A", "numero": 123}}
```

Vira no CSV:
```csv
nome,endereco
Ana,"{""rua"": ""Rua A"", ""numero"": 123}"
```

### Validações do CSV
- Arquivo deve ter cabeçalho (primeira linha)
- Nenhuma coluna pode ter nome vazio
- Nomes de colunas não podem ser duplicados
- Todas as linhas devem ter o mesmo número de campos do cabeçalho
- Não são permitidos campos ausentes ou extras

### Validações do JSON
- Deve ser uma lista (array) de objetos
- Lista não pode ser vazia
- Cada item deve ser um objeto (dicionário)

## Objetivo profissional

Demonstrar habilidades práticas em Python, manipulação de
arquivos, validação de dados, automação, testes e criação
de uma interface web para processamento de informações.

## Melhorias futuras

- Adicionar capturas de tela da aplicação
- Organizar exemplos de entrada e saída
- Ampliar a documentação e os testes
- Preparar uma demonstração para o portfólio
- Suporte a mais formatos de saída (XML, Parquet, etc.)