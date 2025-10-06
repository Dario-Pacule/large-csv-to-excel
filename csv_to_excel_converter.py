#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor CSV para Excel - Otimizado para arquivos grandes (>500MB)
Autor: Assistente IA
Data: 2024
"""

import pandas as pd
import argparse
import os
import sys
from pathlib import Path
import time
from typing import Optional, Tuple
import logging

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('conversao.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class CSVToExcelConverter:
    """Conversor de CSV para Excel otimizado para arquivos grandes."""
    
    def __init__(self, chunk_size: int = 10000):
        """
        Inicializa o conversor.
        
        Args:
            chunk_size: Tamanho do chunk para processamento (padrão: 10000 linhas)
        """
        self.chunk_size = chunk_size
        self.total_rows = 0
        self.processed_rows = 0
        
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
    
    def convert_csv_to_excel(
        self, 
        csv_path: str, 
        excel_path: str,
        sheet_name: str = 'Sheet1',
        encoding: str = 'utf-8',
        delimiter: str = ',',
        skip_rows: int = 0
    ) -> bool:
        """
        Converte arquivo CSV para Excel processando em chunks.
        
        Args:
            csv_path: Caminho do arquivo CSV
            excel_path: Caminho do arquivo Excel de saída
            sheet_name: Nome da planilha
            encoding: Encoding do arquivo CSV
            delimiter: Delimitador do CSV
            skip_rows: Número de linhas para pular no início
            
        Returns:
            bool: True se conversão foi bem-sucedida
        """
        try:
            # Verificar se arquivo CSV existe
            if not os.path.exists(csv_path):
                logger.error(f"Arquivo CSV não encontrado: {csv_path}")
                return False
            
            # Obter informações do arquivo
            file_size_mb = self.get_file_size_mb(csv_path)
            logger.info(f"Tamanho do arquivo: {file_size_mb:.2f} MB")
            
            # Estimar número de linhas
            logger.info("Estimando número de linhas...")
            self.total_rows = self.estimate_total_rows(csv_path)
            logger.info(f"Estimativa de linhas: {self.total_rows:,}")
            
            # Configurar engine de escrita Excel
            excel_writer = pd.ExcelWriter(
                excel_path, 
                engine='openpyxl'
            )
            
            # Processar arquivo em chunks
            logger.info(f"Iniciando conversão com chunks de {self.chunk_size:,} linhas...")
            start_time = time.time()
            
            first_chunk = True
            chunk_number = 0
            
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
                self.processed_rows += len(chunk)
                
                # Log de progresso
                progress = (self.processed_rows / self.total_rows) * 100
                logger.info(f"Processando chunk {chunk_number} - {progress:.1f}% concluído "
                          f"({self.processed_rows:,}/{self.total_rows:,} linhas)")
                
                # Escrever chunk no Excel
                if first_chunk:
                    # Primeiro chunk - escrever com cabeçalho
                    chunk.to_excel(
                        excel_writer, 
                        sheet_name=sheet_name, 
                        index=False,
                        startrow=0
                    )
                    first_chunk = False
                else:
                    # Chunks subsequentes - sem cabeçalho
                    chunk.to_excel(
                        excel_writer, 
                        sheet_name=sheet_name, 
                        index=False,
                        header=False,
                        startrow=excel_writer.sheets[sheet_name].max_row
                    )
                
                # Limpar memória
                del chunk
            
            # Salvar arquivo Excel
            logger.info("Salvando arquivo Excel...")
            excel_writer.close()
            
            # Estatísticas finais
            end_time = time.time()
            duration = end_time - start_time
            output_size_mb = self.get_file_size_mb(excel_path)
            
            logger.info("=" * 50)
            logger.info("CONVERSÃO CONCLUÍDA COM SUCESSO!")
            logger.info(f"Arquivo de entrada: {csv_path}")
            logger.info(f"Arquivo de saída: {excel_path}")
            logger.info(f"Linhas processadas: {self.processed_rows:,}")
            logger.info(f"Tempo total: {duration:.2f} segundos")
            logger.info(f"Tamanho original: {file_size_mb:.2f} MB")
            logger.info(f"Tamanho final: {output_size_mb:.2f} MB")
            logger.info(f"Velocidade: {self.processed_rows/duration:.0f} linhas/segundo")
            logger.info("=" * 50)
            
            return True
            
        except Exception as e:
            logger.error(f"Erro durante conversão: {str(e)}")
            return False
    
    def convert_with_auto_encoding(self, csv_path: str, excel_path: str, **kwargs) -> bool:
        """
        Tenta converter com diferentes encodings automaticamente.
        
        Args:
            csv_path: Caminho do arquivo CSV
            excel_path: Caminho do arquivo Excel de saída
            **kwargs: Argumentos adicionais para convert_csv_to_excel
            
        Returns:
            bool: True se conversão foi bem-sucedida
        """
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                logger.info(f"Tentando encoding: {encoding}")
                if self.convert_csv_to_excel(csv_path, excel_path, encoding=encoding, **kwargs):
                    return True
            except Exception as e:
                logger.warning(f"Falha com encoding {encoding}: {str(e)}")
                continue
        
        logger.error("Falha ao converter com todos os encodings testados")
        return False


def main():
    """Função principal do script."""
    parser = argparse.ArgumentParser(
        description="Converte arquivos CSV para Excel (otimizado para arquivos grandes)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos de uso:
  python csv_to_excel_converter.py arquivo.csv
  python csv_to_excel_converter.py arquivo.csv -o resultado.xlsx
  python csv_to_excel_converter.py arquivo.csv -o resultado.xlsx -s "Dados" -c 5000
  python csv_to_excel_converter.py arquivo.csv -d ";" -e "latin-1"
        """
    )
    
    parser.add_argument('csv_file', help='Caminho do arquivo CSV de entrada')
    parser.add_argument('-o', '--output', help='Caminho do arquivo Excel de saída (opcional)')
    parser.add_argument('-s', '--sheet', default='Sheet1', help='Nome da planilha (padrão: Sheet1)')
    parser.add_argument('-c', '--chunk-size', type=int, default=10000, 
                       help='Tamanho do chunk para processamento (padrão: 10000)')
    parser.add_argument('-d', '--delimiter', default=',', 
                       help='Delimitador do CSV (padrão: vírgula)')
    parser.add_argument('-e', '--encoding', default='auto', 
                       help='Encoding do arquivo CSV (padrão: auto)')
    parser.add_argument('--skip-rows', type=int, default=0, 
                       help='Número de linhas para pular no início (padrão: 0)')
    
    args = parser.parse_args()
    
    # Validar arquivo de entrada
    if not os.path.exists(args.csv_file):
        logger.error(f"Arquivo não encontrado: {args.csv_file}")
        sys.exit(1)
    
    # Definir arquivo de saída se não especificado
    if args.output:
        excel_path = args.output
    else:
        csv_path = Path(args.csv_file)
        excel_path = csv_path.with_suffix('.xlsx')
    
    # Verificar se arquivo de saída já existe
    if os.path.exists(excel_path):
        response = input(f"Arquivo {excel_path} já existe. Sobrescrever? (s/N): ")
        if response.lower() not in ['s', 'sim', 'y', 'yes']:
            logger.info("Conversão cancelada pelo usuário.")
            sys.exit(0)
    
    # Criar conversor
    converter = CSVToExcelConverter(chunk_size=args.chunk_size)
    
    # Executar conversão
    logger.info("Iniciando conversão CSV para Excel...")
    logger.info(f"Arquivo de entrada: {args.csv_file}")
    logger.info(f"Arquivo de saída: {excel_path}")
    
    success = False
    if args.encoding == 'auto':
        success = converter.convert_with_auto_encoding(
            args.csv_file, 
            excel_path,
            sheet_name=args.sheet,
            delimiter=args.delimiter,
            skip_rows=args.skip_rows
        )
    else:
        success = converter.convert_csv_to_excel(
            args.csv_file, 
            excel_path,
            sheet_name=args.sheet,
            encoding=args.encoding,
            delimiter=args.delimiter,
            skip_rows=args.skip_rows
        )
    
    if success:
        logger.info("Conversão concluída com sucesso!")
        sys.exit(0)
    else:
        logger.error("Falha na conversão!")
        sys.exit(1)


if __name__ == "__main__":
    main()
