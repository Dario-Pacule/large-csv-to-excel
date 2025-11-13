#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor CSV para Excel - Arquivos Separados
Gera arquivos Excel separados para arquivos grandes (ex: May_P1.xlsx, May_P2.xlsx)
Autor: Assistente IA
Data: 2024
"""

import pandas as pd
import argparse
import os
import sys
from pathlib import Path
import time
import math
from typing import Optional, Tuple
import logging

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('conversao_arquivos_separados.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class CSVToExcelSeparateFiles:
    """Conversor de CSV para Excel que gera arquivos separados."""
    
    def __init__(self, chunk_size: int = 10000):
        """
        Inicializa o conversor.
        
        Args:
            chunk_size: Tamanho do chunk para processamento (padrão: 10000 linhas)
        """
        self.chunk_size = chunk_size
        self.total_rows = 0
        self.processed_rows = 0
        self.excel_max_rows = 1048576  # Limite do Excel
        
    def get_file_size_mb(self, file_path: str) -> float:
        """Retorna o tamanho do arquivo em MB."""
        return os.path.getsize(file_path) / (1024 * 1024)
    
    def estimate_total_rows(self, csv_path: str) -> int:
        """Estima o número total de linhas no arquivo CSV."""
        try:
            with open(csv_path, 'r', encoding='utf-8') as f:
                # Conta apenas as linhas que não são vazias
                return sum(1 for line in f if line.strip())
        except UnicodeDecodeError:
            # Tenta com encoding diferente
            with open(csv_path, 'r', encoding='latin-1') as f:
                return sum(1 for line in f if line.strip())
    
    def convert_csv_to_separate_excel_files(
        self, 
        csv_path: str, 
        output_prefix: str,
        sheet_name: str = 'Sheet1',
        encoding: str = 'utf-8',
        delimiter: str = ',',
        skip_rows: int = 0,
        max_rows_per_file: int = None
    ) -> bool:
        """
        Converte arquivo CSV para múltiplos arquivos Excel.
        
        Args:
            csv_path: Caminho do arquivo CSV
            output_prefix: Prefixo dos arquivos de saída (ex: "May" para May_P1.xlsx, May_P2.xlsx)
            sheet_name: Nome da planilha
            encoding: Encoding do arquivo CSV
            delimiter: Delimitador do CSV
            skip_rows: Número de linhas para pular no início
            max_rows_per_file: Máximo de linhas por arquivo (padrão: limite do Excel)
            
        Returns:
            bool: True se conversão foi bem-sucedida
        """
        try:
            # Verificar se arquivo CSV existe
            if not os.path.exists(csv_path):
                logger.error(f"Arquivo CSV não encontrado: {csv_path}")
                return False
            
            # Definir limite por arquivo
            if max_rows_per_file is None:
                max_rows_per_file = self.excel_max_rows
            
            # Obter informações do arquivo
            file_size_mb = self.get_file_size_mb(csv_path)
            logger.info(f"Tamanho do arquivo: {file_size_mb:.2f} MB")
            
            # Estimar número de linhas
            logger.info("Estimando número de linhas...")
            self.total_rows = self.estimate_total_rows(csv_path)
            logger.info(f"Estimativa de linhas: {self.total_rows:,}")
            
            # Calcular número de arquivos necessários
            num_files_needed = math.ceil(self.total_rows / max_rows_per_file)
            
            if num_files_needed > 1:
                logger.info("=" * 60)
                logger.info("ARQUIVO MUITO GRANDE PARA UM ARQUIVO EXCEL!")
                logger.info(f"Serão criados {num_files_needed} arquivos separados")
                logger.info(f"Cada arquivo terá no máximo {max_rows_per_file:,} linhas")
                logger.info("=" * 60)
            else:
                logger.info("Arquivo cabe em um único arquivo Excel")
            
            # Processar arquivo em chunks
            logger.info(f"Iniciando conversão com chunks de {self.chunk_size:,} linhas...")
            start_time = time.time()
            
            current_file = 1
            rows_in_current_file = 0
            first_chunk_in_file = True
            chunk_number = 0
            current_excel_writer = None
            current_excel_path = None
            
            for chunk in pd.read_csv(
                csv_path,
                chunksize=self.chunk_size,
                encoding=encoding,
                delimiter=delimiter,
                skiprows=skip_rows,
                low_memory=False,
                dtype=str  # Usar string para evitar problemas de tipo
            ):
                chunk_number += 1
                chunk_rows = len(chunk)
                self.processed_rows += chunk_rows
                
                # Verificar se precisa de novo arquivo
                if rows_in_current_file + chunk_rows > max_rows_per_file:
                    # Fechar arquivo atual
                    if current_excel_writer:
                        logger.info(f"Salvando arquivo: {current_excel_path}")
                        current_excel_writer.close()
                    
                    # Iniciar novo arquivo
                    current_file += 1
                    rows_in_current_file = 0
                    first_chunk_in_file = True
                    
                    # Criar novo arquivo Excel
                    current_excel_path = f"{output_prefix}_P{current_file}.xlsx"
                    logger.info(f"Criando novo arquivo: {current_excel_path}")
                    
                    current_excel_writer = pd.ExcelWriter(
                        current_excel_path, 
                        engine='openpyxl'
                    )
                
                # Se é o primeiro chunk, criar o primeiro arquivo
                if current_excel_writer is None:
                    current_excel_path = f"{output_prefix}_P{current_file}.xlsx"
                    logger.info(f"Criando primeiro arquivo: {current_excel_path}")
                    
                    current_excel_writer = pd.ExcelWriter(
                        current_excel_path, 
                        engine='openpyxl'
                    )
                
                # Log de progresso
                progress = (self.processed_rows / self.total_rows) * 100
                logger.info(f"Processando chunk {chunk_number} - {progress:.1f}% concluído "
                          f"({self.processed_rows:,}/{self.total_rows:,} linhas) - "
                          f"Arquivo {current_file}")
                
                # Escrever chunk no Excel
                if first_chunk_in_file:
                    # Primeiro chunk do arquivo - escrever com cabeçalho
                    chunk.to_excel(
                        current_excel_writer, 
                        sheet_name=sheet_name, 
                        index=False,
                        startrow=0
                    )
                    first_chunk_in_file = False
                else:
                    # Chunks subsequentes - sem cabeçalho
                    chunk.to_excel(
                        current_excel_writer, 
                        sheet_name=sheet_name, 
                        index=False,
                        header=False,
                        startrow=rows_in_current_file
                    )
                
                rows_in_current_file += chunk_rows
                
                # Limpar memória
                del chunk
            
            # Fechar último arquivo
            if current_excel_writer:
                logger.info(f"Salvando último arquivo: {current_excel_path}")
                current_excel_writer.close()
            
            # Estatísticas finais
            end_time = time.time()
            duration = end_time - start_time
            
            # Calcular tamanho total dos arquivos gerados
            total_output_size = 0
            for i in range(1, current_file + 1):
                file_path = f"{output_prefix}_P{i}.xlsx"
                if os.path.exists(file_path):
                    total_output_size += self.get_file_size_mb(file_path)
            
            logger.info("=" * 60)
            logger.info("CONVERSÃO CONCLUÍDA COM SUCESSO!")
            logger.info(f"Arquivo de entrada: {csv_path}")
            logger.info(f"Arquivos gerados: {current_file}")
            for i in range(1, current_file + 1):
                file_path = f"{output_prefix}_P{i}.xlsx"
                if os.path.exists(file_path):
                    size_mb = self.get_file_size_mb(file_path)
                    logger.info(f"  - {file_path}: {size_mb:.2f} MB")
            logger.info(f"Linhas processadas: {self.processed_rows:,}")
            logger.info(f"Tempo total: {duration:.2f} segundos")
            logger.info(f"Tamanho original: {file_size_mb:.2f} MB")
            logger.info(f"Tamanho total final: {total_output_size:.2f} MB")
            logger.info(f"Velocidade: {self.processed_rows/duration:.0f} linhas/segundo")
            logger.info("=" * 60)
            
            return True
            
        except Exception as e:
            logger.error(f"Erro durante conversão: {str(e)}")
            return False
    
    def convert_with_auto_encoding(self, csv_path: str, output_prefix: str, **kwargs) -> bool:
        """
        Tenta converter com diferentes encodings automaticamente.
        
        Args:
            csv_path: Caminho do arquivo CSV
            output_prefix: Prefixo dos arquivos de saída
            **kwargs: Argumentos adicionais para conversão
            
        Returns:
            bool: True se conversão foi bem-sucedida
        """
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                logger.info(f"Tentando encoding: {encoding}")
                if self.convert_csv_to_separate_excel_files(csv_path, output_prefix, encoding=encoding, **kwargs):
                    return True
            except Exception as e:
                logger.warning(f"Falha com encoding {encoding}: {str(e)}")
                continue
        
        logger.error("Falha ao converter com todos os encodings testados")
        return False


def main():
    """Função principal do script."""
    parser = argparse.ArgumentParser(
        description="Converte arquivos CSV para múltiplos arquivos Excel separados",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos de uso:
  python csv_to_excel_separate_files.py arquivo.csv
  python csv_to_excel_separate_files.py arquivo.csv -p "May"
  python csv_to_excel_separate_files.py arquivo.csv -p "Data" -s "Dados" -c 5000
  python csv_to_excel_separate_files.py arquivo.csv -d ";" -e "latin-1"
  python csv_to_excel_separate_files.py arquivo.csv -r 1000000
        """
    )
    
    parser.add_argument('csv_file', help='Caminho do arquivo CSV de entrada')
    parser.add_argument('-p', '--prefix', help='Prefixo dos arquivos de saída (padrão: nome do CSV)')
    parser.add_argument('-s', '--sheet', default='Sheet1', help='Nome da planilha (padrão: Sheet1)')
    parser.add_argument('-c', '--chunk-size', type=int, default=10000, 
                       help='Tamanho do chunk para processamento (padrão: 10000)')
    parser.add_argument('-d', '--delimiter', default=',', 
                       help='Delimitador do CSV (padrão: vírgula)')
    parser.add_argument('-e', '--encoding', default='auto', 
                       help='Encoding do arquivo CSV (padrão: auto)')
    parser.add_argument('--skip-rows', type=int, default=0, 
                       help='Número de linhas para pular no início (padrão: 0)')
    parser.add_argument('-r', '--max-rows', type=int, default=1048576,
                       help='Máximo de linhas por arquivo (padrão: 1048576)')
    
    args = parser.parse_args()
    
    # Validar arquivo de entrada
    if not os.path.exists(args.csv_file):
        logger.error(f"Arquivo não encontrado: {args.csv_file}")
        sys.exit(1)
    
    # Definir prefixo se não especificado
    if args.prefix:
        output_prefix = args.prefix
    else:
        csv_path = Path(args.csv_file)
        output_prefix = csv_path.stem  # Nome do arquivo sem extensão
    
    # Verificar se arquivos de saída já existem
    existing_files = []
    for i in range(1, 10):  # Verificar até 10 arquivos
        file_path = f"{output_prefix}_P{i}.xlsx"
        if os.path.exists(file_path):
            existing_files.append(file_path)
    
    if existing_files:
        logger.warning(f"Arquivos já existem: {', '.join(existing_files)}. Serão sobrescritos.")
        # Remover arquivos existentes
        for file_path in existing_files:
            try:
                os.remove(file_path)
                logger.info(f"Arquivo removido: {file_path}")
            except Exception as e:
                logger.warning(f"Erro ao remover {file_path}: {e}")
    
    # Criar conversor
    converter = CSVToExcelSeparateFiles(chunk_size=args.chunk_size)
    
    # Executar conversão
    logger.info("Iniciando conversão CSV para arquivos Excel separados...")
    logger.info(f"Arquivo de entrada: {args.csv_file}")
    logger.info(f"Prefixo de saída: {output_prefix}")
    
    success = False
    if args.encoding == 'auto':
        success = converter.convert_with_auto_encoding(
            args.csv_file, 
            output_prefix,
            sheet_name=args.sheet,
            delimiter=args.delimiter,
            skip_rows=args.skip_rows,
            max_rows_per_file=args.max_rows
        )
    else:
        success = converter.convert_csv_to_separate_excel_files(
            args.csv_file, 
            output_prefix,
            sheet_name=args.sheet,
            encoding=args.encoding,
            delimiter=args.delimiter,
            skip_rows=args.skip_rows,
            max_rows_per_file=args.max_rows
        )
    
    if success:
        logger.info("Conversão concluída com sucesso!")
        sys.exit(0)
    else:
        logger.error("Falha na conversão!")
        sys.exit(1)


if __name__ == "__main__":
    main()
