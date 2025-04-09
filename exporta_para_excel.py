import pandas as pd
import numpy as np


def export_to_excel_with_formatting(df, file_path, sheet_name='Sheet1',
                                    header_color='#4472C4', even_row_color='#EAEAEA',
                                    odd_row_color='#FFFFFF', border_color='#4472C4'):
    """
    Exporta um DataFrame para Excel com formatação elegante e estilo zebra.

    Parâmetros:
    -----------
    df : pandas.DataFrame
        DataFrame a ser exportado
    file_path : str
        Caminho completo do arquivo Excel a ser criado
    sheet_name : str, opcional
        Nome da planilha (padrão: 'Sheet1')
    header_color : str, opcional
        Cor do cabeçalho em hex (padrão: '#4472C4' - azul corporativo)
    even_row_color : str, opcional
        Cor das linhas pares em hex (padrão: '#EAEAEA' - cinza claro)
    odd_row_color : str, opcional
        Cor das linhas ímpares em hex (padrão: '#FFFFFF' - branco)
    border_color : str, opcional
        Cor da borda externa em hex (padrão: '#4472C4' - azul corporativo)

    Retorna:
    --------
    None
    """

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Estilos personalizados
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': header_color,
            'font_color': 'white',
            'border': 1,
            'border_color': '#ffffff',
            'align': 'center'
        })

        # Formato para primeira coluna (inteiros)
        format_int = workbook.add_format({
            'num_format': '0',
            'border': 1,
            'border_color': '#d9d9d9',
            'align': 'right',
            'bg_color': odd_row_color
        })

        format_int_even = workbook.add_format({
            'num_format': '0',
            'border': 1,
            'border_color': '#d9d9d9',
            'bg_color': even_row_color,
            'align': 'right'
        })

        # Formato para células numéricas (4 decimais)
        format_4decimals = workbook.add_format({
            'num_format': '0.0000',
            'border': 1,
            'border_color': '#d9d9d9',
            'align': 'right',
            'bg_color': odd_row_color
        })

        format_4decimals_even = workbook.add_format({
            'num_format': '0.0000',
            'border': 1,
            'border_color': '#d9d9d9',
            'bg_color': even_row_color,
            'align': 'right'
        })

        # Formato para células de texto
        text_format = workbook.add_format({
            'border': 1,
            'border_color': '#d9d9d9',
            'align': 'left',
            'bg_color': odd_row_color
        })

        text_even = workbook.add_format({
            'border': 1,
            'border_color': '#d9d9d9',
            'bg_color': even_row_color,
            'align': 'left'
        })

        # Ajusta largura das colunas e aplica formatação
        for col_num, col_name in enumerate(df.columns):
            # Define largura baseada no nome da coluna ou conteúdo
            max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 2
            worksheet.set_column(col_num, col_num, min(max_len, 25))

            # Formata cabeçalho
            worksheet.write(0, col_num, col_name, header_format)

            # Formata células do corpo
            if np.issubdtype(df[col_name].dtype, np.number):
                for row_num in range(1, len(df) + 1):
                    if col_num == 0:  # Primeira coluna - formato inteiro
                        if row_num % 2 == 0:
                            worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], format_int_even)
                        else:
                            worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], format_int)
                    else:  # Demais colunas - 4 decimais
                        if row_num % 2 == 0:
                            worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], format_4decimals_even)
                        else:
                            worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], format_4decimals)
            else:
                for row_num in range(1, len(df) + 1):
                    if row_num % 2 == 0:
                        worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], text_even)
                    else:
                        worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], text_format)

        # Adiciona filtros
        worksheet.autofilter(0, 0, 0, len(df.columns) - 1)

        # Congela a linha do cabeçalho
        worksheet.freeze_panes(1, 0)

        # Adiciona borda externa à tabela
        worksheet.conditional_format(0, 0, len(df), len(df.columns) - 1, {
            'type': 'formula',
            'criteria': 'True',
            'format': workbook.add_format({'border': 2, 'border_color': border_color})
        })