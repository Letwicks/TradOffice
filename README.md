# Tradutor de Arquivos Office

Aplicação desktop para tradução em lote de arquivos Excel e Word mantendo a formatação original (cores, layout, fontes, células mescladas).

Desenvolvido por **Autolinx Automação** — http://autolinx.me

## Funcionalidades

- Tradução de conteúdo de arquivos **Excel** (.xls/.xlsx) e **Word** (.doc/.docx)
- Preservação completa da formatação original após tradução
- Cache de traduções para evitar retradução e economizar chamadas à API
- Tradução dos nomes dos arquivos (opcional)
- Seletor de idioma de destino (16 idiomas suportados)
- Escaneamento recursivo de pastas com filtro por tipo de arquivo
- Barra de progresso e log em tempo real
- Interrupção segura a qualquer momento
- Interface gráfica com padrões de acessibilidade (WCAG AA, ISO 9241)

## Requisitos

- Windows (necessário Microsoft Office instalado para arquivos .xls e .doc)
- Python 3.8+

### Dependências

```
pip install openpyxl deep-translator pywin32
```

- `openpyxl` — leitura/escrita de .xlsx nativo
- `deep-translator` — API Google Tradutor (gratuita, sem chave)
- `pywin32` — interoperabilidade COM com Excel e Word para formatação completa

## Instalação

```bash
git clone https://github.com/seu-usuario/tradutor-office.git
cd tradutor-office
pip install openpyxl deep-translator pywin32
python traduzir_gui.py
```

## Como usar

1. Selecione a pasta raiz com os arquivos
2. Escolha os tipos de arquivo (Excel, Word)
3. Selecione o idioma de destino
4. (Opcional) Marque "Trad. nomes" para traduzir também os nomes dos arquivos
5. Clique em **Escanear** para listar os arquivos
6. Marque/desmarque os arquivos desejados
7. Clique em **Iniciar Tradução**

Os arquivos traduzidos serão salvos com sufixo do idioma (ex.: `_PT_BR`) ao lado do original. Arquivos com `_traduzido` no nome são ignorados automaticamente.

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

- `cache_traducao.json` — armazena textos bilíngues para reaproveitamento entre sessões
- `config_tradutor.json` — salva o último diretório utilizado
- `icone_app.ico` — gerado automaticamente na primeira execução

## Licença

MIT — Autolinx Automação
