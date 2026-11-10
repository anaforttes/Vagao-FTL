# Automação de Certificados de Vagões-Tanque FTL (Locarga CSN)

Este projeto automatiza o processo de consulta, download, leitura (OCR) dos **Certificados de Verificação para Vagão-Tanque Ferroviário (Inmetro)** no portal **FTL-Locarga** (`https://locarga-ftl.csn.com.br/locarga`) e o preenchimento completo da planilha **[VTs FTL NEW.xlsx](./VTs%20FTL%20NEW.xlsx)**.



## 1. Qual era o problema?
Anteriormente, o processo era totalmente manual:
1. Pegar o código de cada um dos **369 vagões** na planilha.
2. Acessar o site do Locarga FTL, pesquisar vagão por vagão e baixar o PDF do certificado.
3. Abrir cada PDF (que é um documento digitalizado em formato de imagem, sem texto selecionável) e digitar manualmente na planilha:
   - Datas de verificação e validade;
   - Régua mínima e régua máxima (`mm`);
   - Tabela de calibração das **Medidas no Domo** (níveis `4`, `3`, `2`, `1` em `L` e `mm`);
   - Tabela de calibração das **Medidas no Tanque** (níveis `0` a `10` em `L` e `mm`).



## 2. O que foi feito para automatizar?

A automação foi dividida em **4 etapas integradas**, sem necessidade de instalar programas externos (utilizando recursos nativos do Windows, PowerShell e C#):

### Etapa 1 — Consulta direta na API do Portal Locarga FTL
Em vez de pesquisar placa por placa na tela do site, identificamos o endpoint oficial da API do sistema:
- **Endpoint:** `GET https://locarga-ftl.csn.com.br/locarga/api/vagao/todos?pag=0&pgsize=2000&ord=&dir=`
- Com uma única requisição autenticada, o script obtém a lista completa de vagões cadastrados na FTL, incluindo a `data_insercao` e o caminho direto do arquivo PDF (`caminho_arquivo`) no servidor.

### Etapa 2 — Cruzamento e Download em Lote dos PDFs
- O script lê todos os vagões listados na planilha **`VTs FTL NEW.xlsx`** (369 vagões).
- Cruza automaticamente os códigos dos vagões com o retorno da API (tratando zeros à esquerda quando necessário).
- Realiza o download paralelo de todos os **369 certificados em PDF** para a pasta **`pdfs_ftl/`**, salvando cada arquivo no padrão `<codigo_vagao>_<nome_original>.pdf` (ex.: `6089020_TCC608902-0-AF-20-10-22.pdf`).

### Etapa 3 — Leitura Inteligente dos PDFs via OCR (Nativo do Windows)
Como os certificados do Inmetro são imagens digitalizadas, o script utiliza as bibliotecas nativas do Windows (`Windows.Data.Pdf.PdfDocument` + `Windows.Media.Ocr.OcrEngine`) com processamento paralelo em múltiplos núcleos do processador:
- **Mapeamento por coordenadas `(X, Y)`:** Cada palavra/número reconhecido no PDF possui sua posição exata na página, permitindo separar com precisão:
  - **Data de Validade (`Fim Validade`):** canto superior direito do certificado;
  - **Data da Verificação (`Início Validade`):** canto inferior esquerdo do certificado;
  - **Tabela `MEDIDAS NO TANQUE`:** colunas `VOLUME L` e `ESPAÇO VAZIO mm` (níveis `0` a `10`) à esquerda;
  - **Tabela `MEDIDAS NO DOMO`:** colunas `VOLUME L` e `ESPAÇO VAZIO mm` (níveis `1` a `4`) à direita.
- **Validação e Correção Física/Geométrica:**
  - Os volumes do Domo seguem incrementos de `+50 L` acima da capacidade nominal (`CapNominal + 50`, `+ 100`, `+ 150`, `+ 200`).
  - Os volumes do Tanque iniciam na capacidade nominal (nível `0`) e decrescem de `100 L` em `100 L` a partir da centena imediatamente inferior.
  - Todos os casos em que o carimbo, assinatura ou linha de grade do formulário tocavam os números foram auditados e validados para garantir **100% de preenchimento (369/369 vagões sem campos vazios)**.

### Etapa 4 — Preenchimento Direto na Planilha (`VTs FTL NEW.xlsx`)
Um compilador rápido em C# (`pdfs_ftl/build_excel.ps1`) atualiza diretamente a estrutura XML interna da planilha **`VTs FTL NEW.xlsx`**, preservando a formatação, filtros e colunas originais.



## 3. Estrutura das Colunas na Planilha (`VTs FTL NEW.xlsx`)

| Coluna(s) | Nome da Coluna | Origem do Dado |
| :--- | :--- | :--- |
| **A – I** | `Ser`, `Veic`, `Nivel`, `Volume`, `Capacidade Volumétrica`, `Capacidade Máxima (T)`, `Comprimento Total`, `Nro Certificado Arqueacao`, `Validade Certificado` | Mantidas exatamente como na planilha original `VTs FTL NEW.xlsx` |
| **J** | `Início Validade` | Extraído do PDF (*Data da verificação* do Inmetro) |
| **K** | `Régua Mínima (mm)` | Menor `ESPAÇO VAZIO mm` medido no tanque (nível `0`) |
| **L** | `Régua Máxima (mm)` | Maior `ESPAÇO VAZIO mm` medido no tanque (último nível do tanque) |
| **M** | `Situação` | Calculado automaticamente: `Regular` (dentro da validade) ou `Vencido` |
| **N** | `Data de Inserção` | Extraído da API do portal Locarga (`data_insercao`) |
| **O – V** | `Domo 4 (L)`, `Domo 4 (mm)` até `Domo 1 (L)`, `Domo 1 (mm)` | Extraído da tabela **MEDIDAS NO DOMO** do certificado PDF |
| **W – AR** | `Tanque 0 (L)`, `Tanque 0 (mm)` até `Tanque 10 (L)`, `Tanque 10 (mm)` | Extraído da tabela **MEDIDAS NO TANQUE** do certificado PDF |



## 4. Arquivos do Projeto

- **`VTs FTL NEW.xlsx`**: Planilha principal já preenchida com todos os 369 vagões e todas as medidas de Domo e Tanque.
- **`Atualizar_Vagoes_FTL.bat`**: Atalho executável de **1 clique** para rodar toda a automação.
- **`Atualizar_Vagoes_FTL.ps1`**: Script principal que conecta na API do Locarga, baixa novos PDFs, executa o OCR e chama o gerador do Excel.
- **`pdfs_ftl/`**:
  - Contém todos os **369 certificados em PDF** baixados do site;
  - `vagoes_api.json`: Dados brutos retornados pela API do Locarga;
  - `ocr_words_368.json`: Cache das leituras de OCR de todos os certificados (evita ter que reler PDFs que não mudaram);
  - `build_excel.ps1`: Motor em C# que monta e atualiza as colunas na planilha `VTs FTL NEW.xlsx`;
  - `VTs FTL NEW_backup.xlsx`: Cópia de segurança da planilha original antes do preenchimento.



## 5. Como atualizar a planilha no futuro?

1. **Feche a planilha `VTs FTL NEW.xlsx`** no Excel (para que o arquivo não fique bloqueado para gravação).
2. Caso queira incluir novos vagões, basta adicioná-los nas linhas finais da `VTs FTL NEW.xlsx` (preenchendo ao menos `Ser`, `Veic` e `Capacidade Volumétrica`) e salvar.
3. Dê dois cliques no arquivo **`Atualizar_Vagoes_FTL.bat`**.
4. O script irá consultar o site da FTL, baixar apenas os certificados novos, ler os dados via OCR e atualizar a **`VTs FTL NEW.xlsx`** automaticamente em poucos segundos!

