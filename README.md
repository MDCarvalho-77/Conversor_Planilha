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

## Exportação para o Blip (V3)

A V3 adiciona a modalidade **"Blip — lista de contatos"**, que prepara arquivos
CSV prontos para upload no Blip, com um Excel intermediário para conferência e
edição manual. Os fluxos e as saídas anteriores permanecem inalterados.

### Requisitos do CSV do Blip

Cada CSV gerado respeita:

- Exatamente duas colunas, nesta ordem: `Número ou BSUID` e `Nome`.
- Uma linha por contato, sem índice do pandas nem colunas extras.
- No máximo **100 contatos** por arquivo (sem contar o cabeçalho).
- No máximo **100 KB** de tamanho real do arquivo gravado.
- Identificador no formato de telefone internacional com `+` (ex.: `+5548999990001`)
  ou um BSUID válido (aceito na reimportação, quando o usuário fornecer).
- Codificação **UTF-8** (sem BOM) e separador vírgula, preservando acentos,
  vírgulas, aspas e o `+` inicial.

Os arquivos são nomeados sequencialmente: `Blip_001.csv`, `Blip_002.csv`, etc.
Cada CSV é revalidado após a gravação, medindo os bytes reais e reabrindo o
conteúdo. Nenhum arquivo existente é sobrescrito silenciosamente: se a pasta de
saída já contiver `Blip_*.csv`, o aplicativo avisa.

### Fluxo em duas etapas

1. **Preparação.** Selecione as planilhas (CRM 1, CRM 2 ou genéricas, até 7),
   escolha a saída "Blip — lista de contatos" e clique em processar. O
   aplicativo reaproveita a consolidação, a deduplicação global e a
   normalização de telefones da V2, mapeia o resultado para as duas colunas do
   Blip, divide em lotes (respeitando 100 contatos **e** 100 KB) e gera:
   - os CSVs iniciais (`Blip_001.csv`...);
   - um Excel de conferência `Blip_Contatos_Revisao.xlsx`, com uma aba por lote
     (`Lote_001`, `Lote_002`...), telefones como texto (preservando o `+` e sem
     notação científica), cabeçalho e autofiltro.

2. **Geração final após a revisão.** Edite o Excel livremente (inserir, remover,
   mover ou corrigir contatos) e use **"Gerar CSVs a partir da planilha
   revisada"**. O aplicativo lê as abas, valida as duas colunas e os
   identificadores, **reaplica os limites de 100 contatos e 100 KB** —
   redistribuindo automaticamente mesmo que uma aba tenha passado de 100
   contatos — e grava os CSVs definitivos na pasta escolhida. O Excel revisado
   não é modificado.

### Regras de validação

- O identificador (telefone) é obrigatório. Registros sem identificador válido
  não são descartados em silêncio: aparecem no resumo como inválidos.
- O nome pode faltar; nesse caso o contato é mantido com nome vazio (o
  aplicativo não inventa nomes) e é contabilizado no resumo.
- Se um único contato não couber em 100 KB nem sozinho, ele não é truncado: é
  reportado como inválido.

O resumo apresentado ao final informa contatos exportados, número de arquivos
gerados e eventuais registros inválidos. O processamento roda fora da thread da
interface e cliques repetidos não iniciam execuções simultâneas.

## Separação automática por DDD (V3.1)

A V3.1 é uma evolução incremental (não é uma V4): tudo o que já funcionava na
V1, V2 e V3 permanece intacto. O foco é atender à operação comercial, que
trabalha com três números de WhatsApp associados aos DDDs **47**, **48** e
**49**. Após a consolidação e a deduplicação global, os contatos preparados para
o Blip são classificados automaticamente em **quatro grupos**.

### Os quatro grupos

1. **DDD_47** — somente contatos cujo telefone brasileiro tenha DDD 47.
2. **DDD_49** — somente contatos cujo telefone brasileiro tenha DDD 49.
3. **DDD_48_E_OUTROS** — contatos com DDD 48 e todos os demais DDDs brasileiros
   válidos diferentes de 47 e 49.
4. **SEM_DDD** — contatos que não podem ser classificados com segurança: BSUID,
   telefones internacionais que não sejam brasileiros e telefones sem DDD
   brasileiro identificável.

Cada contato cai em **exatamente um** grupo, e nenhum contato é descartado.

### Regra conservadora de classificação

A classificação usa o telefone **já normalizado** pela lógica existente (a mesma
regra de telefone da V1/V2) e não inventa nem infere DDD por nome, origem da
planilha ou qualquer outro campo:

- Um telefone é reconhecido como brasileiro somente quando os dígitos começam
  com `55` **e** o total é de **12 dígitos** (`55` + DDD de 2 + número de 8) ou
  **13 dígitos** (`55` + DDD de 2 + número de 9). O DDD são os dois dígitos logo
  após o `55`, válido na faixa 11–99.
- Qualquer outra forma vai para **SEM_DDD**: BSUID e identificadores com letras,
  telefones internacionais não brasileiros (ex.: `+1...`, `+351...`, `+44...`) e
  `+55` com comprimento incompatível (que não permita inferir o DDD com
  segurança).
- Os dois primeiros dígitos de um telefone internacional **não** são confundidos
  com DDD brasileiro.

### Estrutura de pastas Blip/DDD_x/

A exportação Blip da V3.1 organiza os arquivos por grupo dentro de uma pasta
`Blip/`:

```text
Blip/
    DDD_47/
        Blip_001.csv
        Blip_Contatos_Revisao.xlsx
    DDD_49/
        Blip_001.csv
        Blip_Contatos_Revisao.xlsx
    DDD_48_E_OUTROS/
        Blip_001.csv
        Blip_Contatos_Revisao.xlsx
    SEM_DDD/
        Blip_001.csv
        Blip_Contatos_Revisao.xlsx
```

Regras da estrutura:

- A subpasta de um grupo é criada **somente** quando o grupo tem contatos; um
  grupo vazio não gera pasta.
- A numeração dos CSVs é **independente por grupo**: cada grupo recomeça em
  `Blip_001.csv` e segue sequencialmente (`Blip_002.csv`...) quando há mais de um
  lote.
- Cada grupo tem a sua **própria** planilha de revisão `Blip_Contatos_Revisao.xlsx`;
  contatos de grupos diferentes nunca se misturam no mesmo CSV nem na mesma
  planilha.
- Os limites e o formato dos CSVs do Blip são preservados (duas colunas na ordem
  `Número ou BSUID` e `Nome`, máximo de 100 contatos e 100 KB reais por arquivo,
  UTF-8 sem BOM, separador vírgula, `+` preservado) e nenhum `Blip_*.csv`
  existente é sobrescrito silenciosamente.

### Duas opções de saída na interface

A interface passa a oferecer **somente duas** opções de saída:

- **Opção A — Planilha completa + números convertidos.** Gera a planilha
  completa consolidada (`<nome>_tratado.xlsx`) e a lista de telefones convertidos
  (`<nome>_telefones.xlsx`). Corresponde exatamente ao comportamento das "duas
  saídas" já existente, sem mudança de campos, colunas, formatação ou regras.
- **Opção B — Blip — listas de contatos (separadas por DDD).** Opção principal da
  operação comercial: prepara os contatos, separa automaticamente nos quatro
  grupos e gera, para cada grupo com contatos, os CSVs e a planilha de revisão na
  estrutura `Blip/DDD_x/`.

As opções intermediárias ou redundantes deixaram de aparecer na interface. As
funções e constantes internas que ainda sustentam compatibilidade e testes
foram preservadas — a simplificação é da **experiência de uso**, não da base de
código.

### Reimportação de revisão com reclassificação por DDD

Ao gerar os CSVs finais a partir das planilhas revisadas, o usuário pode
selecionar **múltiplos** arquivos de revisão (um por grupo). O sistema não confia
cegamente na pasta ou na aba de origem: cada contato é **reclassificado pelo
telefone atual** com a mesma regra conservadora de DDD. Assim:

- Se o telefone foi alterado e mudou de DDD, o contato é encaminhado ao grupo
  correto.
- Se o telefone deixou de ser identificável com segurança, o contato vai para
  SEM_DDD.
- Os contatos são reagrupados e os limites de 100 contatos e 100 KB são
  reaplicados por grupo.

O resumo da operação informa quantos **contatos mudaram de grupo** após a revisão
(considerando a origem conhecida pela pasta-pai do arquivo). As planilhas de
revisão **originais não são modificadas**.

### Feedback de sucesso

Na preparação Blip, o resumo informa o total de contatos processados, a
quantidade em cada grupo, os CSVs gerados por grupo, a quantidade sem DDD
identificável e a localização da pasta raiz `Blip/`. Na geração dos CSVs
revisados, o resumo informa o total exportado, quantos contatos mudaram de grupo,
os CSVs por grupo, a quantidade em SEM_DDD e a localização. Nenhum sucesso é
exibido antes de concluir a gravação de todos os arquivos; em caso de erro,
nenhuma mensagem de sucesso é mostrada.

## Separação por curso e DDD (V3.2)

A V3.2 acrescenta uma **terceira** modalidade de saída, sem alterar nenhuma das
duas anteriores: **"Blip - listas de contatos separadas por curso e DDD"**. Ela
primeiro separa os contatos por **curso** e, dentro de cada curso, pelos mesmos
quatro grupos de DDD da V3.1 (DDD_47, DDD_49, DDD_48_E_OUTROS e SEM_DDD).

### Coluna de curso obrigatória (somente nesta modalidade)

A obrigatoriedade da coluna de curso vale **exclusivamente** para esta terceira
opção. As opções A (planilha completa + números) e B (Blip por DDD) continuam
funcionando sem exigir curso. Antes de gerar qualquer arquivo, todas as
planilhas selecionadas são validadas: se **alguma** não tiver uma coluna de
curso reconhecível ou houver **ambiguidade**, o processamento é interrompido com
uma única mensagem que lista todas as planilhas problemáticas — e **nenhum**
arquivo parcial é gravado.

### Rótulos de curso reconhecidos e ordem de prioridade

A coluna de curso é identificada pelo **nome** do cabeçalho (tolerante a
maiúsculas/minúsculas, acentos e espaços), nunca pelo conteúdo das células. Os
rótulos conhecidos são avaliados nesta ordem de prioridade — primeiro os rótulos
explícitos de curso/interesse, depois os de produto:

1. `Curso`
2. `Nome do Curso`
3. `Tipo de Curso`
4. `Curso de Interesse`
5. `Interesse` / `Formação de Interesse` / `Graduação de Interesse` / `Área de Interesse`
6. `Nome do Produto`
7. `Produto Primário`
8. `Produto Principal`
9. `Primário (Nome do produto) (Produto)`
10. `Produto`
11. `Produto de Interesse` / `Interesse em Produto`

Quando mais de um rótulo **de produto** conhecido está presente, a prioridade
acima resolve qual usar — isso **não** é ambiguidade. A ambiguidade que
interrompe ocorre apenas quando há duas ou mais categorias claramente distintas
de curso/interesse presentes como colunas separadas (por exemplo, `Curso` **e**
`Curso de Interesse`). Nesse caso a mensagem lista as colunas candidatas
encontradas.

### Tabela de equivalências e `CURSO_NAO_IDENTIFICADO`

Os nomes de curso passam por uma tabela de equivalências **cadastrada
manualmente** em `cursos.py` (`EQUIVALENCIAS_CURSO`). A semente inicial unifica
as três variações de Análise e Desenvolvimento de Sistemas em uma única
denominação padrão. Nenhuma similaridade é inferida automaticamente: cursos sem
equivalência cadastrada **preservam o nome original**. Para **estender**,
acrescente um par `normalizar_rotulo("Nome Variante"): "Denominação Padrão"` ao
dicionário — nenhuma equivalência adicional é cadastrada sem confirmação.

Um contato cuja célula de curso esteja **vazia** (mesmo com a coluna presente)
não interrompe o processamento: ele é classificado como
`CURSO_NAO_IDENTIFICADO`, que também é subdividido pelos quatro grupos de DDD.

### Estrutura de pastas Blip/<Curso>/<DDD>/

```text
Blip/
    Analise e Desenvolvimento de Sistemas/
        DDD_47/
            Blip_001.csv
            Blip_Contatos_Revisao.xlsx
        DDD_49/
            Blip_001.csv
            Blip_Contatos_Revisao.xlsx
    CURSO_NAO_IDENTIFICADO/
        SEM_DDD/
            Blip_001.csv
            Blip_Contatos_Revisao.xlsx
```

O nome da pasta de cada curso é sanitizado para o Windows e colisões entre
nomes diferentes são desambiguadas de forma determinística. Uma combinação
curso+DDD só gera pasta quando tem contatos; a numeração dos CSVs reinicia em
`Blip_001.csv` por combinação.

### Formato dos CSVs e da planilha de revisão

- Os **CSVs** continuam com **exatamente duas colunas** (`Número ou BSUID` e
  `Nome`), respeitando os mesmos limites (100 contatos e 100 KB reais por
  arquivo, UTF-8 sem BOM, separador vírgula, `+` preservado). E-mail e curso
  **não** entram no CSV oficial do Blip.
- A **planilha de revisão** da V3.2 (`Blip_Contatos_Revisao.xlsx`) tem **quatro
  colunas**: `Número ou BSUID`, `Nome`, `E-mail` e `Curso`. A coluna de telefone
  é texto (preservando o `+`) e há autofiltro. E-mail e curso ficam na revisão
  para conferência manual, mas não são exportados para o CSV.

### Reimportação com autodetecção (2 vs 4 colunas)

O botão **"Gerar CSVs a partir da planilha revisada"** é **único** e detecta o
formato automaticamente pelos **cabeçalhos** das abas visíveis (não apenas pela
quantidade de colunas):

- `Número ou BSUID, Nome` → revisão antiga (V3.1), reclassificada apenas por DDD.
- `Número ou BSUID, Nome, E-mail, Curso` → revisão nova (V3.2), reclassificada
  por curso **e** DDD.

As revisões antigas de duas colunas continuam funcionando normalmente. Em cada
reimportação v3.2, o curso é repadronizado (curso apagado →
`CURSO_NAO_IDENTIFICADO`) e o DDD é reclassificado pelo telefone atual (telefone
alterado muda a pasta de DDD; inidentificável → `SEM_DDD`). Os arquivos de
revisão originais **não são modificados**. O resumo informa, além dos totais por
curso e por DDD, quantos contatos mudaram de curso, mudaram de DDD, foram para
`CURSO_NAO_IDENTIFICADO` e foram para `SEM_DDD`.

## Histórico de versões

- **V1 — concluída:** processamento inicial.
- **V2 — concluída:** três formatos de saída.
- **V3.1 — concluída:** robustez e processamento em background.
- **V3.2 — concluída:** UX e identidade visual.
- **V2 (multiformato) — concluída:** múltiplos arquivos (até 7), suporte a um
  segundo CRM e a planilhas genéricas por telefone, consolidação com
  deduplicação global, nova saída "Planilha completa estendida" e ícone próprio
  do executável. O fluxo clássico da V1 permanece inalterado.
- **V3 (Blip) — concluída:** modalidade "Blip — lista de contatos" com geração
  de CSVs (até 100 contatos / 100 KB), Excel de conferência por lote e geração
  dos CSVs definitivos a partir do Excel revisado. Fluxos V1/V2 inalterados.
- **V3.1 (separação por DDD) — concluída:** classificação automática dos
  contatos Blip em quatro grupos (DDD_47, DDD_49, DDD_48_E_OUTROS e SEM_DDD) com
  regra conservadora de identificação do DDD, exportação na estrutura
  `Blip/DDD_x/` (subpasta e numeração independentes por grupo, um Excel de
  revisão por grupo), simplificação da interface para duas opções de saída
  (A: planilha completa + números convertidos; B: Blip separado por DDD) e
  reimportação de múltiplas planilhas de revisão com reclassificação pelo
  telefone atual e contador de contatos que mudaram de grupo. Os originais de
  revisão não são modificados e os fluxos V1/V2/V3 permanecem inalterados.
- **V3.2 (separação por curso e DDD) — concluída:** terceira modalidade de saída
  "Blip - listas de contatos separadas por curso e DDD", que separa por curso e,
  dentro de cada curso, pelos quatro grupos de DDD. Coluna de curso **obrigatória
  apenas nesta modalidade**, com validação prévia que interrompe sem gerar
  parciais em caso de ausência ou ambiguidade. Reconhecimento tolerante da coluna
  de curso por ordem de prioridade (curso/interesse antes de produto), tabela de
  equivalências cadastrada manualmente (semente ADS, extensível) com preservação
  do nome original quando não há equivalência, marcador `CURSO_NAO_IDENTIFICADO`,
  estrutura `Blip/<Curso>/<DDD>/`, CSV mantido em duas colunas e planilha de
  revisão de quatro colunas (`Número ou BSUID`, `Nome`, `E-mail`, `Curso`). O
  botão de reimportação é único e autodetecta revisões de 2 colunas (V3.1) e de
  4 colunas (V3.2) pelos cabeçalhos. As duas modalidades anteriores e os fluxos
  V1/V2/V3/V3.1 permanecem inalterados.
