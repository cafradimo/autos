import streamlit as st
import pandas as pd
from fpdf import FPDF
import tempfile
import os
from datetime import datetime

class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.page_width = 210  # A4 width in mm
        self.cell_width = self.page_width - 20  # Margens de 10mm cada lado
    
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Relatório de Autuações', 0, 1, 'C')
        self.set_font('Arial', '', 10)
        self.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", 0, 1)
        self.ln(5)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}/{{nb}}', 0, 0, 'C')

def process_excel_file(uploaded_file):
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(uploaded_file, sheet_name='AIs')
        
        # Verificar se as colunas necessárias existem
        required_columns = ['RF', 'EMISSÃO', 'CODIGO', 'DESCRICAO DA AUTUAÇÃO']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.warning(f"Arquivo {uploaded_file.name} não contém as colunas necessárias: {', '.join(missing_columns)}")
            return None, None
        
        # Selecionar apenas as colunas necessárias
        processed_df = df[required_columns].copy()
        
        # Contar ocorrências de cada descrição de autuação
        descricao_counts = processed_df['DESCRICAO DA AUTUAÇÃO'].value_counts().to_dict()
        
        return processed_df, descricao_counts
        
    except Exception as e:
        st.error(f"Erro ao processar o arquivo {uploaded_file.name}: {str(e)}")
        return None, None

def generate_pdf_report(df, descricao_counts, filename):
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    
    # Seção de Estatísticas
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Estatísticas de Ocorrências:', 0, 1)
    pdf.set_font('Arial', '', 10)
    
    # Ordenar por quantidade (maior primeiro)
    sorted_counts = sorted(descricao_counts.items(), key=lambda item: item[1], reverse=True)
    
    for descricao, count in sorted_counts:
        # Quebra de linha automática para descrições longas
        pdf.multi_cell(pdf.cell_width, 6, f"TOTAL: {descricao} = {count:02d}", 0, 1)
    
    pdf.ln(10)
    
    # Seção de Dados
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Dados das Autuações:', 0, 1)
    
    # Configurar larguras das colunas
    col_widths = {
        'RF': 30,
        'EMISSÃO': 25,
        'CODIGO': 20,
        'DESCRICAO DA AUTUAÇÃO': pdf.cell_width - 75  # O restante do espaço
    }
    
    # Cabeçalho da tabela
    pdf.set_font('Arial', 'B', 10)
    for col in df.columns:
        pdf.cell(col_widths[col], 10, str(col), border=1)
    pdf.ln()
    
    # Dados da tabela
    pdf.set_font('Arial', '', 8)  # Fonte menor para os dados
    for _, row in df.iterrows():
        # Ajustar altura da linha baseado no conteúdo
        max_lines = 1
        for col in df.columns:
            lines = pdf.get_string_width(str(row[col])) // col_widths[col] + 1
            if lines > max_lines:
                max_lines = lines
        
        line_height = 6 * max_lines
        
        for col in df.columns:
            # Usar multi_cell para quebra de texto automática
            x = pdf.get_x()
            y = pdf.get_y()
            pdf.multi_cell(col_widths[col], line_height/max_lines, str(row[col]), border=1, align='L')
            pdf.set_xy(x + col_widths[col], y)
        pdf.ln(line_height)
        
        # Verificar se precisa de nova página
        if pdf.get_y() > 270:  # Margem inferior
            pdf.add_page()
            # Repetir cabeçalho se nova página
            pdf.set_font('Arial', 'B', 10)
            for col in df.columns:
                pdf.cell(col_widths[col], 10, str(col), border=1)
            pdf.ln()
            pdf.set_font('Arial', '', 8)
    
    # Salvar o PDF
    pdf.output(filename)

def main():
    st.set_page_config(page_title="Processador de Autuações", layout="wide")
    
    st.title("📋 Processador de Planilhas de Autuações")
    st.write("""
    Este aplicativo processa planilhas Excel contendo dados de autuações e gera:
    1. Uma planilha Excel simplificada com as colunas RF, EMISSÃO, CODIGO e DESCRICAO DA AUTUAÇÃO
    2. Um relatório PDF com estatísticas das ocorrências de cada tipo de autuação
    """)
    
    with st.expander("ℹ️ Instruções"):
        st.markdown("""
        1. Selecione uma ou mais planilhas Excel no formato padrão
        2. O sistema processará cada arquivo individualmente
        3. Para cada arquivo válido, serão gerados:
           - Uma planilha Excel com os dados filtrados
           - Um relatório PDF com estatísticas
        4. Clique nos botões de download para salvar os arquivos
        """)
    
    uploaded_files = st.file_uploader("Selecione as planilhas Excel para processar", 
                                    type=['xls', 'xlsx'], 
                                    accept_multiple_files=True)
    
    if uploaded_files:
        for uploaded_file in uploaded_files:
            with st.container():
                st.subheader(f"📄 Processando: {uploaded_file.name}")
                st.write("---")
                
                # Processar o arquivo
                processed_df, descricao_counts = process_excel_file(uploaded_file)
                
                if processed_df is not None and descricao_counts is not None:
                    # Criar arquivo Excel temporário
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
                        excel_filename = tmp_excel.name
                        processed_df.to_excel(excel_filename, index=False)
                    
                    # Gerar PDF
                    pdf_filename = f"relatorio_{os.path.splitext(uploaded_file.name)[0]}.pdf"
                    generate_pdf_report(processed_df, descricao_counts, pdf_filename)
                    
                    # Mostrar resultados
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.success("✅ Arquivo processado com sucesso!")
                        st.write(f"📊 Total de registros: {len(processed_df)}")
                        
                        # Exibir estatísticas em uma tabela
                        st.subheader("📈 Estatísticas de Ocorrências")
                        stats_df = pd.DataFrame.from_dict(descricao_counts, orient='index', columns=['Total'])
                        stats_df.index.name = 'Descrição da Autuação'
                        st.dataframe(stats_df.sort_values('Total', ascending=False))
                    
                    with col2:
                        # Exibir prévia dos dados
                        st.subheader("👀 Prévia dos Dados")
                        st.dataframe(processed_df.head())
                    
                    # Botões de download
                    st.write("---")
                    st.subheader("📥 Download dos Arquivos Gerados")
                    
                    dl_col1, dl_col2 = st.columns(2)
                    
                    with dl_col1:
                        # Ler o arquivo Excel novamente para garantir que está fechado
                        with open(excel_filename, 'rb') as f:
                            st.download_button(
                                label="⬇️ Baixar Planilha Processada",
                                data=f,
                                file_name=f"processado_{uploaded_file.name}",
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                key=f"excel_{uploaded_file.name}"
                            )
                    
                    with dl_col2:
                        with open(pdf_filename, 'rb') as f:
                            st.download_button(
                                label="⬇️ Baixar Relatório PDF",
                                data=f,
                                file_name=pdf_filename,
                                mime='application/pdf',
                                key=f"pdf_{uploaded_file.name}"
                            )
                    
                    # Limpar arquivos temporários após o download
                    try:
                        os.unlink(excel_filename)
                        os.unlink(pdf_filename)
                    except PermissionError as e:
                        st.warning(f"Não foi possível excluir os arquivos temporários: {str(e)}")
                        # Tentar novamente mais tarde ou deixar para o sistema limpar
                
                st.write("---")

if __name__ == "__main__":
    main()