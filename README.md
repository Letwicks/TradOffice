# Tradutor de Arquivos Office

Aplicação desktop para tradução em lote de arquivos Excel e Word mantendo a formatação original (cores, layout, fontes, células mescladas).

Desenvolvido por **Autolinx Automação** — http://autolinx.me

## Funcionalidades

- Tradução de conteúdo de arquivos **Excel** (.xls/.xlsx) e **Word** (.doc/.docx)
- Preservação da formatação original: fórmulas, imagens, campos e notas de rodapé não são tocados
- Word: corpo, tabelas, cabeçalhos, rodapés e caixas de texto; Excel: células, nomes de abas e caixas de texto
- Cache de traduções **por idioma** para evitar retradução e economizar chamadas à API
- Textos curtos agrupados em uma única requisição (bem menos chamadas ao Google)
- Tradução dos nomes dos arquivos (opcional)
- Idioma de origem (ou detecção automática) e de destino (16 idiomas suportados)
- Escaneamento recursivo de pastas com filtro por tipo de arquivo
- Barra de progresso e log em tempo real, resumo de sucesso/falhas ao final
- Interrupção segura a qualquer momento e opção "Pular já traduzidos" para retomar
- Interface gráfica com padrões de acessibilidade (WCAG AA, ISO 9241)

## Requisitos

- Windows com Microsoft Office instalado para arquivos .xls, .doc e .docx
  (sem Office, apenas .xlsx funciona, via openpyxl, com perda de imagens/gráficos)
- Python 3.8+

### Dependências

```
pip install openpyxl deep-translator pywin32
```

- `openpyxl` — leitura/escrita de .xlsx nativo
- `deep-translator` — API Google Tradutor (gratuita, sem chave)
- `pywin32` — interoperabilidade COM com Excel e Word para formatação completa.
  O programa abre instâncias **próprias e invisíveis** do Excel/Word: as janelas do
  Office que você estiver usando não são afetadas.

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
4. (Opcional) Marque "Traduzir nomes" para traduzir também os nomes dos arquivos
5. Clique em **Escanear** para listar os arquivos
6. Marque/desmarque os arquivos desejados
7. Clique em **Iniciar Tradução**

Os arquivos traduzidos serão salvos com sufixo do idioma (ex.: `_PT_BR`) ao lado do original; `.xls` vira `.xlsx`. Arquivos gerados pelo programa (sufixo de idioma ou `_traduzido` no nome) e arquivos temporários do Office (`~$...`) são ignorados no escaneamento.

Parágrafos do Word que contêm imagens, campos ou notas de rodapé são mantidos no original para não perder esses elementos — o log informa quantos.

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

- `cache_traducao.json` — traduções separadas por idioma de destino, reaproveitadas entre sessões
  (o formato antigo, da v1.1, é migrado automaticamente para o idioma `pt`)
- `config_tradutor.json` — última pasta, idiomas e opções escolhidas
- `icone_app.ico` — gerado automaticamente na primeira execução

Ficam ao lado do `traduzir_gui.py` (ou do `.exe`; se a pasta do `.exe` não for gravável, em `%LOCALAPPDATA%\TradOffice`).

## Testes

Rodam em qualquer sistema, sem Office e sem internet (o tradutor e o COM são simulados):

```bash
python -m unittest discover -s tests
```

## Licença

MIT — Autolinx Automação
