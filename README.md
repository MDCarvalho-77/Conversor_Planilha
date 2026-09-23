# Conversor CRM

Aplicativo desktop para tratar planilhas Excel exportadas do CRM e gerar arquivos prontos para uso. A interface foi pensada para executar o fluxo de forma guiada, sem exigir que o usuário trabalhe com código Python.

## Objetivo

O Conversor CRM facilita o tratamento de planilhas exportadas do CRM para usuários leigos. Ele concentra regras já validadas para filtragem, deduplicação e normalização de telefones, eliminando a necessidade de executar scripts Python manualmente.

## Público-alvo

Usuários que trabalham com planilhas do CRM e precisam gerar uma planilha tratada ou uma lista de telefones sem conhecimento técnico de programação.

## Fluxo do processamento

Ao processar uma planilha, o aplicativo executa este fluxo, nesta ordem:

1. Seleciona a planilha Excel de entrada.
2. Lê somente as abas visíveis.
3. Ignora abas com estado `hidden` e `veryHidden`.
4. Normaliza os nomes das colunas, removendo espaços no início e no fim.
5. Valida a estrutura e as colunas necessárias para cada etapa.
6. Remove registros sem telefone.
7. Remove duplicatas integrais.
8. Normaliza os telefones restantes.
9. Produz internamente o DataFrame final tratado.
10. Gera uma ou ambas as saídas conforme o formato escolhido.

## Regras de negócio

### Deduplicação

A deduplicação é feita antes da normalização dos telefones e compara a linha inteira, sem definir subconjunto de colunas. Portanto:

- Somente linhas integralmente idênticas são removidas.
- Todas as colunas presentes na planilha participam da comparação.
- `Nome do produto` participa obrigatoriamente.
- `Primário (Nome do produto) (Produto)` participa obrigatoriamente.
- Se qualquer informação for diferente, inclusive uma das duas colunas de produto, os registros não são considerados duplicados.

Exemplos:

- Duas linhas iguais em todas as colunas, inclusive nos produtos, resultam em uma única linha final.
- Duas linhas iguais, mas com `Nome do produto` diferente, permanecem como dois registros.
- Duas linhas iguais, mas com `Primário (Nome do produto) (Produto)` diferente, permanecem como dois registros.

### Telefones

O tratamento de telefones segue exatamente estas regras:

- Caracteres não numéricos são removidos.
- Telefones com 10 dígitos recebem o prefixo `55`.
- Telefones com 11 dígitos recebem o prefixo `55`.
- Telefones com 12 dígitos são mantidos após a remoção de caracteres não numéricos.
- Telefones com 13 dígitos são mantidos após a remoção de caracteres não numéricos.
- Outros tamanhos são mantidos somente com seus dígitos.
- Registros sem telefone são removidos antes da geração das saídas.

Exemplos: `(48) 99999-0001` torna-se `5548999990001`; um telefone de 13 dígitos já contendo o código do país permanece com os mesmos dígitos.

## Estrutura esperada da planilha CRM

A planilha deve ser uma exportação compatível do CRM. As colunas relevantes para o processamento são:

- `Razão do Status`
- `Nome do produto`
- `Primário (Nome do produto) (Produto)`
- `CPF_CNPJ`
- ` Nome Completo`
- `E-mail`
- `Telefone da Empresa/Celular`

Os nomes das colunas têm espaços nas extremidades normalizados pelo processamento. Assim, a coluna de origem ` Nome Completo` é tratada como `Nome Completo` na etapa seguinte e na saída completa.

## Abas ignoradas

Somente abas com estado `visible` são lidas. Abas `hidden` e `veryHidden` não são tratadas como registros e não participam das métricas, validações, deduplicação ou saídas.

## Formatos de saída

As três opções utilizam o mesmo DataFrame final tratado.

### Planilha completa

Gera `<nome_original>_tratado.xlsx` com as cinco colunas:

1. `Razão do Status`
2. `Nome Completo`
3. `CPF`
4. `E-mail`
5. `Telefone da Empresa/Celular`

### Somente números

Gera `<nome_original>_telefones.xlsx` com uma única coluna:

`Telefone`

### As duas saídas

Gera os dois arquivos, `<nome_original>_tratado.xlsx` e `<nome_original>_telefones.xlsx`. Ambos usam o mesmo DataFrame final; não há segundo processamento nem nova deduplicação.

## Interface

Para usar o aplicativo:

1. Clique em **SELECIONAR PLANILHA** e escolha o arquivo exportado do CRM.
2. Escolha o formato de saída desejado.
3. Clique em **PROCESSAR PLANILHA**.
4. Consulte as métricas e os nomes dos arquivos gerados na área **RESULTADO**.
5. Clique em **ABRIR PASTA** para abrir a pasta que contém os arquivos gerados.

O processamento é executado em segundo plano para manter a interface responsiva. Durante a execução, o botão de processamento fica desabilitado para evitar solicitações duplicadas.

## Métricas

Após o processamento, a interface apresenta:

- **Registros lidos:** quantidade de linhas lidas nas abas visíveis.
- **Removidos sem telefone:** quantidade de registros descartados por não conterem telefone válido para o filtro inicial.
- **Duplicados removidos:** quantidade de linhas integrais repetidas removidas após o filtro de telefone.
- **Registros finais:** quantidade de registros no DataFrame final tratado.

## Arquitetura do projeto

- `app.py`: interface CustomTkinter, validação amigável da estrutura, execução em segundo plano e apresentação de métricas e arquivos.
- `processamento.py`: regras de leitura, validação, filtro, deduplicação, normalização de telefone e geração das saídas.
- `test/`: testes de carga com arquivos Excel temporários.
- `assets/`: recursos visuais da interface, incluindo a logomarca oficial SESI SENAI.
- `requirements.txt`: dependências Python do projeto.
- `Conversor_CRM.spec`: configuração PyInstaller preparada para o futuro executável Windows.

Também há `test_processamento.py`, que contém os testes funcionais de regressão das saídas e métricas.

## Tecnologias

- Python
- pandas
- openpyxl
- CustomTkinter
- Pillow
- PyInstaller

## Desenvolvimento

No Windows PowerShell, a partir da pasta do projeto:

```powershell
.\.venv\Scripts\Activate.ps1
python app.py
```

Se a política do PowerShell impedir a ativação, execute diretamente o interpretador do ambiente virtual:

```powershell
.\.venv\Scripts\python.exe app.py
```

## Testes

Os testes funcionais atuais validam a saída completa, a saída somente de telefones, a geração das duas saídas, a igualdade dos telefones entre elas e as métricas da planilha de cenários.

- Testes funcionais: **3 aprovados**.
- Teste de carga com dados sintéticos e arquivos temporários:
  - 5.000 registros: **1,23 s**.
  - 10.000 registros: **2,35 s**.
  - 25.000 registros: **5,83 s**.

Esses tempos são referências obtidas em ambiente de desenvolvimento com dados sintéticos; não representam garantia universal de desempenho.

Comandos utilizados:

```powershell
python -m unittest -v
python test\test_desempenho.py
```

## Distribuição

O objetivo de distribuição é gerar um executável Windows independente chamado `Conversor_CRM.exe` com PyInstaller. A configuração `Conversor_CRM.spec` já está preparada, incluindo a logomarca em `assets/`, mas o build ainda está em preparação e o executável ainda não foi gerado.

## Manutenção e expansão futura

Novas regras de negócio devem ser implementadas em `processamento.py`. Alterações de interface, textos, layout e interação devem ser feitas em `app.py`.

Cuidados importantes:

- Não alterar regras de deduplicação sem revisar e atualizar os testes.
- Não alterar o tratamento de telefone sem atualizar os testes.
- Preservar a compatibilidade com a estrutura exportada pelo CRM.
- Atualizar este README quando uma regra de negócio mudar.

## Múltiplos formatos e múltiplos arquivos (V2)

A partir da V2, o aplicativo aceita vários arquivos por execução e mais de um
formato de origem, mantendo intactas as regras e as saídas da versão anterior.

### Seleção de arquivos

- É possível selecionar **até 7 arquivos** Excel em uma única execução.
- Se mais de 7 forem selecionados, o aplicativo avisa e não processa, para que
  a seleção seja corrigida.
- Formatos diferentes podem ser combinados na mesma execução (por exemplo, dois
  arquivos do CRM 1, um do CRM 2 e uma planilha genérica).

### Formatos de origem reconhecidos

Cada arquivo é identificado individualmente pelos cabeçalhos de suas abas
visíveis (abas `hidden`/`veryHidden` continuam ignoradas), na ordem:

1. **CRM 1** — formato da V1 (assinatura: `CPF_CNPJ` + `Telefone da Empresa/Celular`).
2. **CRM 2** — segundo formato de CRM. Telefone escolhido por prioridade
   (`Telefone Celular` → `Telefone da Empresa` → `Telefone Residencial`); o
   nome usa `Nome Completo` e, quando vazio, compõe `Nome` + `Segundo Nome` +
   `Sobrenome`; o documento vem de `CPF`; o e-mail de `E-mail Comercial`.
3. **Planilha genérica** — qualquer planilha que tenha uma coluna reconhecível
   como telefone (a comparação olha apenas o **nome** da coluna, tolerando
   maiúsculas/minúsculas, acentos e espaços). Nomes aceitos incluem `Telefone`,
   `Celular`, `Telefone Celular`, `Telefone Residencial`, `Telefone da Empresa`,
   `WhatsApp`, entre outros. Os demais campos (`Nome`, `CPF`, `E-mail`) são
   aproveitados quando existirem e ficam vazios quando não houver.

Se um arquivo não corresponder a nenhum desses formatos por não ter coluna de
telefone reconhecível, o aplicativo mostra uma mensagem clara explicando o que
precisa ser ajustado, sem interromper com erro técnico.

### Modelo canônico e consolidação

Cada arquivo é convertido para um modelo interno comum com os campos: Razão do
Status, Nome Completo, CPF, E-mail, Telefone, Regional/Unidade do Produto, Nome
do Produto e Produto Primário. Todos os arquivos são então **consolidados** em
um único conjunto, sobre o qual ocorrem: remoção de registros sem telefone,
**deduplicação global** e normalização de telefones (a mesma regra da V1).

> A deduplicação global compara os campos do modelo canônico. Isso vale dentro
> do mesmo arquivo e entre arquivos e formatos diferentes. Essa regra aplica-se
> **apenas** a este fluxo (V2). O fluxo clássico da V1 — um único arquivo do
> CRM 1 com as três saídas originais — permanece inalterado, com deduplicação
> integral das linhas.

### Nova saída: Planilha completa estendida

Além das três saídas existentes, há a opção **Planilha completa estendida**,
que gera `<nome>_estendido.xlsx` com sete colunas:

1. `Razão do Status`
2. `Nome Completo`
3. `CPF`
4. `E-mail`
5. `Telefone da Empresa/Celular`
6. `Regional/Unidade do Produto`
7. `Nome do Produto`

Campos ausentes na origem (por exemplo, `Regional/Unidade do Produto` em
arquivos do CRM 1) ficam vazios. Todas as saídas partem do mesmo conjunto final;
não há segunda deduplicação por saída.

### Nomes dos arquivos de saída

- Um único arquivo de entrada: a saída usa o nome do próprio arquivo (igual à V1).
- Vários arquivos: a saída recebe o prefixo `Consolidado_` sobre o nome do
  primeiro arquivo e é gravada na pasta dele.

### Ícone do aplicativo

O executável passa a ter um ícone próprio, gerado a partir da logomarca
(`assets/conversor_crm.ico`) e configurado no `Conversor_CRM.spec`. O script
`gerar_icone.py` regenera o `.ico` a partir do PNG oficial quando necessário.

### Arquitetura da V2

- `adapters.py`: identificação de formato e mapeamento de cada origem para o
  modelo canônico. Concentra as diferenças entre CRMs, evitando condicionais de
  formato espalhadas pelo código de negócio.
- `pipeline_v2.py`: consolidação de múltiplos arquivos, deduplicação global e
  geração das saídas (inclusive a estendida). Reutiliza as funções da V1 para
  leitura de abas visíveis, limpeza de nomes de coluna e normalização de
  telefone.

Os testes da V2 (`test_adapters.py`, `test_consolidacao.py`) cobrem detecção de
formato, mapeamento de cada CRM, prioridade e fallback do CRM 2, reconhecimento
genérico de telefone, erro amigável sem telefone, consolidação e deduplicação
global entre arquivos e formatos, e a saída estendida. Os testes da V1
permanecem como referência de regressão e não devem ser removidos.

## Histórico de versões

- **V1 — concluída:** processamento inicial.
- **V2 — concluída:** três formatos de saída.
- **V3.1 — concluída:** robustez e processamento em background.
- **V3.2 — concluída:** UX e identidade visual.
- **V2 (multiformato) — concluída:** múltiplos arquivos (até 7), suporte a um
  segundo CRM e a planilhas genéricas por telefone, consolidação com
  deduplicação global, nova saída "Planilha completa estendida" e ícone próprio
  do executável. O fluxo clássico da V1 permanece inalterado.
