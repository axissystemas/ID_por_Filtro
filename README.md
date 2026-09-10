# ID por Filtro

Script de automação para extração e agrupamento de IDs de chamados por categoria a partir de planilhas Excel (.xlsx), gerando arquivos CSV formatados para filtros em sistemas de chamados (Jira, ServiceNow, SQL, etc.).

## 📌 Funcionalidades

- **Leitura dinâmica de planilhas:** Identifica os IDs (Coluna A) e as Categorias (Coluna C).
- **Limpeza de IDs:** Remove espaços internos (ex: `15 653` -> `15653`) e caracteres indesejados.
- **Geração de CSV Agrupado:** Lista as categorias com o total de chamados e a lista de IDs separados por vírgula (`IDs_agrupados_por_categoria.csv`).
- **Geração de CSV Detalhado:** Mapeamento 1 a 1 de cada registro com `Categoria,ID` (`IDs_por_categoria_detalhado.csv`).
- **Suporte a Acentuação:** Arquivos exportados com codificação `UTF-8 com BOM (utf-8-sig)` para abertura direta e correta no Microsoft Excel.

## 🚀 Como Usar

### Pré-requisitos
Certifique-se de ter o Python instalado e as dependências necessárias:

```bash
pip install pandas openpyxl
```

### Execução
Coloque a planilha no mesmo diretório (ex: `Chamados N3 SGBOM.xlsx`) e execute:

```bash
python extrair_ids.py
```
