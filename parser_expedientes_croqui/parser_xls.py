import xlrd
import os
import pandas as pd
import pickle
from pathlib import PurePath
import datetime

def crawler_arquivos_tipo(extension, dir_ = os.getcwd()):
    '''Percorre uma árvore de diretórios e obtém os 
    paths absolutos de todos os arquivos de uma determinada extensão'''
    #Essa função está fora da classe porque ela pertence a uma outra classe que abstrai diversas funções de parseamento de arquivos.
    #Coloquei ela aqui por comodidade, para não ter que realizar o import
    
    paths = [(tupla[0], tupla[2]) for tupla in list(os.walk(dir_))]
    path_files = [os.path.abspath(os.path.join(tupla[0], file)) for tupla in paths for file in tupla[1]]
    files = [path for path in path_files if os.path.isfile(path) and extension == PurePath(path).suffix]
    
    return files


class ParserXls():
    '''Implementa o pipeline de parseamento dos arquivos .xls que contém as informações das
    Fichas de Anotação de Expedientes dos Croquis Patrimoniais de CGPATRI'''
    
    def __init__(self, dir_arquivos):
        
        self.dir_arquivos = dir_arquivos
        self.files = crawler_arquivos_tipo('.xls', dir_ = self.dir_arquivos)
        self.maper = dict(
                        croqui = (3, 0),
                        area = (3, 3),
                        info_num = (5, 0),
                        proc = (7, 0),
                        expediente = (11, 0),
                        interessado = (13, 0),
                        assunto = (16, 0),
                        local = (19, 0),
                        anotacao = (21, 1),
                        informacao = (21, 6),
                        despacho = (23, 0),
                        dt_DOM = (24, 0),
                        obs = (26, 0),
                        data = (33,0),
                        autor = (33, 2)
                        )
        self.maper_txt = {
            'Área:' : 'area',
            'ÁREA:' : 'area',
            'ASSUNTO:' : 'assunto',
            'NOME:' : 'autor',
            'CROQUIS:' : 'croqui',
            'Data:' : 'data',
            'DESPACHO:' : 'despacho',
            'PUBLICADO NO DOM EM:' : 'dt_DOM',
            'Nº DE EXPEDIENTE:' : 'expediente',
            'Nº DA INFORMAÇÃO ' : 'info_num',
            'INTERESSADO:' : 'interessado',
            'PROCESSO:' : 'proc',
            'Processo:' : 'proc',
            'LOCAL:' : 'local',
            'DATA:' : 'data',
            'INFORMAÇÕES/OBS.:' : 'obs',
            'Informações/Obs.:' : 'obs',
            'Informações:' : 'obs'
        }
        
        self.files_erro = {}
        
    def format_cel_value(self, cel, item = ''):
        '''Formata o valor, em string, da célula, removendo os nomes de campos
        presentes nas células, para deixar apenas a informação desejada, entre outros
        detalhes'''
        
        limpar = self.maper_txt.keys()
        cel = cel.replace('text:', '')
        cel = cel.replace("'", '')
        cel = cel.replace('empty:', '')
        
        if not item:
            for item in limpar:
                if cel.startswith(item):
                    cel = cel.replace(item, '')
                    break
        else:
            cel = cel.replace(item, '')
        if not len(cel):
            cel = ''
        if pd.isnull(cel):
            cel = ''
        
        return cel
        
    def parser_xls_file(self, file):
        '''Parseia o arquivo .xls a partir da lógica das posições das células
        conforme o atributo self.maper'''
        
        wb = xlrd.open_workbook(file)
        s = wb.sheet_by_name('Plan1')

        result = {}

        for nom, cel in self.maper.items():
            cel = s.cell(*cel)
            cel = str(cel)
            result[nom] = self.format_cel_value(cel)
        result['file'] = file
        return result
    
    def parser_xls_repescagem(self, file):
        '''Parseia o arquivo .xls de forma mais custosa, buscando,
        em cada célula da planilha, a presença dos nomes dos campos.
        Esse método deve ser usado apenas para os arquivos que não puderam
        ser parseados pelo método .parser_xls_file'''
        
        wb = xlrd.open_workbook(file)
        s = wb.sheet_by_name('Plan1')
        
        result = {}
        
        for row in range(s.nrows):
            for col in range(s.ncols):
                cel = str(s.cell(row, col)).replace('text:', '').replace("'", '')
                for item, value in self.maper_txt.items():
                    if value not in result.keys():
                        if cel.startswith(item):
                            result[value] = self.format_cel_value(cel, item = item)
        for col_nom in self.maper.keys():
            if col_nom not in result.keys():
                result[col_nom] = ''
        result['file'] = file
        return result
        
    def parser_todos_xls(self):
        '''Parseia todos os arquivos .xls presentes no atributo self.files, fazendo
        a repescagem daqueles que não puderam ser parseados pelo método self.parser_xls_file por
        meio do método self.parser_xls_repescagem, que é mais custoso'''
        
        result = []
        for file in self.files:
            try:
                result.append(self.parser_xls_file(file))
            except:
                try:
                    result.append(self.parser_xls_repescagem(file))
                except Exception as e:
                    self.files_erro[file] = e

        return pd.DataFrame(result)
    
    def arrumar_dt(self, df):
        '''Formata a coluna de datas do DataFrame.
        Nos casos em que houve problema na conversão do formato de data do excel para o python,
        coloca como valor para a data a data da última modificação do arquivo'''
        
        df = df.copy()
        df['data'] = df['data'].apply(lambda x: str(x).replace('DATA:', ''))
        df['data'] = df['data'].apply(lambda x: str(x).replace('OBS.:', ''))
        df['data'] = df['data'].apply(lambda x: str(x).replace('.', '/'))
        for i in range(len(df)):
            if 'xl' in  str(df.loc[i, 'data']): #indica o erro de conversão
                dtime = os.path.getmtime(df.loc[i, 'file'])
                dtime = datetime.datetime.fromtimestamp(dtime)
                df.loc[i, 'data'] = '/'.join([str(dtime.day), str(dtime.month), str(dtime.year)])
        return df
        
    def format_df(self, df):
        '''Formata o dataframe, arrumando as datas por meio do método anterior, e transformando
        as colunas booleanas, que eram marcadas com um "X" nos formulários em flags 0 e 1'''
        
        df = df.copy()
        df = self.arrumar_dt(df)
        df['anotacao'] = df['anotacao'].apply(lambda x: 1 if 'x' in str(x) else 0)
        df['informacao'] = df['informacao'].apply(lambda x: 1 if 'x' in str(x) else 0)

        for col in df.keys():
            df[col] = df[col].apply(lambda x: str(x).strip())
            
        return df
    
    def main(self):
        
        df = self.parser_todos_xls()
        df = self.format_df(df)
        self.df = df
        