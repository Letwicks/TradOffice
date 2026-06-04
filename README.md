# Tradutor de Arquivos Office

Aplicacao desktop para traducao em lote de arquivos Excel, Word e PDF mantendo a formatacao original (cores, layout, fontes, celulas mescladas).

## Funcionalidades

- Traducao de conteudo de arquivos **Excel** (.xls/.xlsx), **Word** (.doc/.docx) e **PDF**
- Preservacao completa da formatacao original apos traducao
- Cache de traducoes para evitar retraducao e economizar chamadas a API
- Traducao dos nomes dos arquivos (opcional)
- Seletor de idioma de destino (16 idiomas suportados)
- Escaneamento recursivo de pastas com filtro por tipo de arquivo
- Barra de progresso e log em tempo real
- Interrupcao segura a qualquer momento
- Interface grafica com padroes de acessibilidade (WCAG AA, ISO 9241)

## Requisitos

- Windows (necessario Microsoft Office instalado para arquivos .xls e .doc)
- Python 3.8+

### Dependencias

```
pip install openpyxl deep-translator pywin32
```

- `openpyxl` — leitura/escrita de .xlsx nativo
- `deep-translator` — API Google Tradutor (gratuita, sem chave)
- `pywin32` — interoperabilidade COM com Excel e Word para formatacao completa

## Instalacao

```bash
git clone https://github.com/seu-usuario/tradutor-office.git
cd tradutor-office
pip install openpyxl deep-translator pywin32
python traduzir_gui.py
```

## Como usar

1. Selecione a pasta raiz com os arquivos
2. Escolha os tipos de arquivo (Excel, Word, PDF)
3. Selecione o idioma de destino
4. (Opcional) Marque "Trad. nomes" para traduzir tambem os nomes dos arquivos
5. Clique em **Escanear** para listar os arquivos
6. Marque/desmarque os arquivos desejados
7. Clique em **Iniciar Traducao**

Os arquivos traduzidos serao salvos com sufixo do idioma (ex.: `_PT_BR`) ao lado do original. Arquivos com `_traduzido` no nome sao ignorados automaticamente.

## Estrutura de dados

```
CHECKLIST_CANTEIRO/
├── obra1/
│   ├── M55.780435_Checklist.xlsx
│   └── M55.780435_Checklist_PT_BR.xlsx  (traduzido)
├── obra2/
│   └── relatorio.doc
└── ...
```

## Cache e config

- `cache_traducao.json` — armazena textos bilingues para reaproveitamento entre sessoes
- `config_tradutor.json` — salva o ultimo diretorio utilizado
- `icone_app.ico` — gerado automaticamente na primeira execucao

## Licenca

MIT
