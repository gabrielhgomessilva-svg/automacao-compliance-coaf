# Automação de Compliance COAF

Programa em **Python** que gera, de forma automática, as planilhas diárias de classificação de risco por operação (COAF), lendo os borderôs de pagamento direto do **SQL Server** e aplicando as regras de negócio de compliance.

> Projeto de portfólio. Os dados, caminhos, nomes e a planilha modelo da empresa **não** fazem parte deste repositório; os exemplos e testes usam dados fictícios.

## O problema

Todo dia era preciso montar, uma a uma, a planilha de classificação de risco de cada operação concluída, consultando o cadastro de risco do cliente e verificando pagamentos a terceiros. Era um trabalho manual, repetitivo e sujeito a erro de digitação.

## Como funciona

![Fluxo](docs/img/fluxo.png)

1. **Extração:** consulta os borderôs concluídos do dia (factoring e FIDC).
2. **Cruzamento:** busca risco, motivo e CNPJ do cliente na planilha de controle de risco.
3. **Regras:** aplica as regras de compliance a cada operação.
4. **Geração:** preenche o modelo em Excel e salva na pasta do dia (`MÊS\dd.mm.aaaa`).

O usuário só escolhe o dia (Enter = hoje); as operações são encontradas sozinhas.

## Regras de negócio

![Regras](docs/img/regras.png)

| Situação | O que o programa faz | Resultado |
|---|---|---|
| Cliente com risco Elevado | Copia o motivo do cadastro para a fundamentação | Arquivo com sufixo `S` |
| Pagamento a terceiro acima do limite | Compara o CNPJ do favorecido com o do cliente e marca SIM | Arquivo com sufixo `S` |
| Motivo ausente no cadastro | Deixa em branco, sem inventar texto | Arquivo com sufixo `S` |
| Operação sem pagamento | Gera normalmente | Planilha padrão |
| Cliente fora do cadastro | Não gera e avisa | Pendente |

O limite de pagamento a terceiro é configurável (`LIMITE_TERCEIRO`).

## Segurança

![Segurança](docs/img/seguranca.png)

- Somente **consultas de leitura** (`SELECT`), com `READ UNCOMMITTED` e `LOCK_TIMEOUT`, sem bloquear quem usa o sistema.
- Consultas **parametrizadas** (sem concatenar texto em SQL).
- Abre a planilha de controle **somente para leitura**.
- Nunca grava, altera ou exclui dados no sistema; avisa em vez de decidir quando falta informação.
- Recomendado: usar um usuário de banco com permissão **apenas de leitura**.
- Credenciais ficam no `config.ini` local, que está no `.gitignore`. **Nunca** versione esse arquivo.

## Tecnologias

| Tecnologia | Uso |
|---|---|
| Python 3 | Regras de negócio, tratamento de erros e testes |
| SQL Server + pyodbc | Consultas de leitura parametrizadas |
| Excel via COM (pywin32) | Abre o modelo, preenche células, atualiza vínculos e salva em `.xls` |
| pdfplumber | Lê borderôs em PDF (fonte alternativa) |
| PyInstaller | Gera um executável que roda com dois cliques |

## Estrutura

```
coaf_automacao.py      # programa principal
config.example.ini     # exemplo de configuração (copie para config.ini)
gerar_exe.bat          # gera o executável (Windows)
requirements.txt
tests/test_logica.py   # testes da lógica (dados fictícios)
docs/img/              # imagens deste README
```

A planilha modelo (`modelo/modelo_planilha.xls`) **não** está no repositório. É preciso fornecer a sua, com o formulário nas células indicadas no código (constantes `CEL_*`).

## Como usar

Requisitos: Windows, Microsoft Excel, Python 3 e *ODBC Driver for SQL Server* (17 ou 18).

```bash
pip install -r requirements.txt
copy config.example.ini config.ini     # e ajuste os valores
python coaf_automacao.py               # pergunta o dia; Enter = hoje
python coaf_automacao.py 07.10.2026    # outro dia
python coaf_automacao.py 900101 900102 # operações específicas
```

Para gerar o executável: `gerar_exe.bat`.

## Testes

```bash
python tests/test_logica.py
```

Cobrem leitura de borderôs, regra de terceiros, fundamentação, listagem por data e a camada SQL (simulada em SQLite). A parte de Excel/COM exige Windows com Excel e foi validada manualmente em base de teste.

## Limitações

- Depende de Windows e Excel (usa COM).
- O esquema de tabelas consultado (`SIGBORS`, `SIGBDCHQ` etc.) é específico do ERP em que foi desenvolvido; para outro sistema, ajuste as consultas `SQL_*` no código.
- O ganho de tempo ainda será medido no uso diário.

## Licença

[MIT](LICENSE)
