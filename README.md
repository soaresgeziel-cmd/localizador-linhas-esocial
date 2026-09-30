# 🔎 Localizador de Linhas - eSocial

Aplicação desktop leve e portátil desenvolvida em Python para busca, filtragem e extração rápida de registros do eSocial (CPFs ou Matrículas) a partir de arquivos texto e CSV locais.

## 🚀 Funcionalidades

- **Varredura Recursiva:** Busca em pastas mãe e todas as suas subpastas.
- **Performance Otimizada:** Motor em Regex unificado e pré-filtro ultrarrápido em C.
- **Validação de Caminho:** Checagem instantânea de consistência entre o evento selecionado e o nome da pasta de origem.
- **Filtragem por Operação e Competência:** Separação entre cargas Normal e Retificação, além de suporte a múltiplas competências separadas por vírgula.
- **Carregamento em Lote:** Leitura de planilhas Excel com ajuste automático de zeros à esquerda para CPFs (`zfill(11)`).
- **Sem Instalação:** Distribuído como um arquivo executável autônomo (`.exe`).

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+**
- **CustomTkinter** (Interface Gráfica Windows Light)
- **Pandas** & **Openpyxl** (Processamento de dados e Excel)
- **PyInstaller** (Empacotamento)

## 📦 Como Usar o Executável

1. Acesse a aba [Releases](../../releases) deste repositório.
2. Baixe o arquivo `Localizador_de_Linhas.exe`.
3. Dê duplo clique para executar (não requer Python ou direitos de administrador).