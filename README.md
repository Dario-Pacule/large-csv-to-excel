# Conversor CSV para Excel

Script Python otimizado para converter arquivos CSV para Excel, especialmente projetado para lidar com arquivos grandes (mais de 500MB).

## Características

- ✅ **Processamento em chunks**: Processa arquivos grandes sem sobrecarregar a memória
- ✅ **Múltiplos encodings**: Suporte automático para diferentes encodings (UTF-8, Latin-1, CP1252, etc.)
- ✅ **Logging detalhado**: Acompanhe o progresso da conversão em tempo real
- ✅ **Tratamento de erros**: Validações e tratamento robusto de erros
- ✅ **Interface de linha de comando**: Fácil de usar via terminal
- ✅ **Estatísticas de performance**: Relatório detalhado do tempo e velocidade de processamento

## Instalação

1. **Instalar Python** (versão 3.7 ou superior)

2. **Instalar dependências**:

   ```bash
   pip install -r requirements.txt
   ```

   Ou instalar manualmente:

   ```bash
   pip install pandas openpyxl
   ```

## Uso

### Uso Básico

```bash
# Converter arquivo CSV para Excel (mesmo nome, extensão .xlsx)
python csv_to_excel_converter.py arquivo.csv

# Especificar arquivo de saída
python csv_to_excel_converter.py arquivo.csv -o resultado.xlsx
```

### Opções Avançadas

```bash
# Especificar nome da planilha
python csv_to_excel_converter.py arquivo.csv -o resultado.xlsx -s "Dados"

# Ajustar tamanho do chunk (para arquivos muito grandes)
python csv_to_excel_converter.py arquivo.csv -c 5000

# Usar delimitador diferente (ponto e vírgula)
python csv_to_excel_converter.py arquivo.csv -d ";"

# Especificar encoding manualmente
python csv_to_excel_converter.py arquivo.csv -e "latin-1"

# Pular linhas no início do arquivo
python csv_to_excel_converter.py arquivo.csv --skip-rows 2
```

### Parâmetros Disponíveis

| Parâmetro          | Descrição                            | Padrão                      |
| ------------------ | ------------------------------------ | --------------------------- |
| `csv_file`         | Caminho do arquivo CSV (obrigatório) | -                           |
| `-o, --output`     | Caminho do arquivo Excel de saída    | Mesmo nome do CSV com .xlsx |
| `-s, --sheet`      | Nome da planilha                     | "Sheet1"                    |
| `-c, --chunk-size` | Tamanho do chunk para processamento  | 10000                       |
| `-d, --delimiter`  | Delimitador do CSV                   | ","                         |
| `-e, --encoding`   | Encoding do arquivo CSV              | "auto"                      |
| `--skip-rows`      | Linhas para pular no início          | 0                           |

## Exemplos de Uso

### Arquivo CSV simples

```bash
python csv_to_excel_converter.py dados.csv
```

### Arquivo CSV com ponto e vírgula

```bash
python csv_to_excel_converter.py dados.csv -d ";"
```

### Arquivo grande com chunk menor

```bash
python csv_to_excel_converter.py arquivo_grande.csv -c 5000 -o resultado.xlsx
```

### Arquivo com encoding específico

```bash
python csv_to_excel_converter.py dados.csv -e "latin-1" -o resultado.xlsx
```

## Otimizações para Arquivos Grandes

O script foi otimizado para lidar com arquivos grandes através de:

1. **Processamento em chunks**: Divide o arquivo em pedaços menores para processar
2. **Gerenciamento de memória**: Libera memória após cada chunk
3. **Encoding automático**: Tenta diferentes encodings automaticamente
4. **Logging de progresso**: Mostra o progresso em tempo real

### Recomendações para arquivos muito grandes (>1GB):

- Use chunks menores: `-c 5000` ou `-c 2000`
- Monitore o uso de memória do sistema
- Certifique-se de ter espaço em disco suficiente (arquivos Excel são maiores que CSV)

## Logs

O script gera logs detalhados que incluem:

- Progresso da conversão
- Estatísticas de performance
- Informações sobre o arquivo (tamanho, número de linhas)
- Tempo total de processamento
- Velocidade de processamento

Os logs são salvos em:

- **Arquivo**: `conversao.log`
- **Console**: Saída em tempo real

## Tratamento de Erros

O script inclui tratamento robusto de erros para:

- Arquivos não encontrados
- Problemas de encoding
- Arquivos corrompidos
- Falta de espaço em disco
- Problemas de permissão

## Limitações

- **Tamanho máximo**: Limitado pela memória disponível e espaço em disco
- **Formato Excel**: Gera arquivos .xlsx (Excel 2007+)
- **Tipos de dados**: Todos os dados são convertidos para string para evitar problemas de tipo

## Solução de Problemas

### Erro de encoding

```bash
# Tente com encoding específico
python csv_to_excel_converter.py arquivo.csv -e "latin-1"
```

### Arquivo muito grande

```bash
# Use chunks menores
python csv_to_excel_converter.py arquivo.csv -c 2000
```

### Falta de memória

- Reduza o tamanho do chunk (`-c 1000`)
- Feche outros programas
- Use um computador com mais RAM

## Suporte

Para problemas ou dúvidas:

1. Verifique os logs em `conversao.log`
2. Teste com um arquivo menor primeiro
3. Verifique se todas as dependências estão instaladas

## Licença

Este script é fornecido "como está" para uso pessoal e comercial.
