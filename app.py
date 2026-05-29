import streamlit as st
from google import genai
from google.genai import types
import json
import io
import zipfile

# Bibliotecas para extração de dados do Microsoft Office
import docx2txt
import openpyxl
from pptx import Presentation

st.set_page_config(page_title="Prêmio Innpulse Bahia - Gestão Documental", page_icon="📁", layout="wide")
st.title("Sistema Inteligente de Governança Documental - Prêmio Innpulse Bahia 📁")
st.write("Organização escalável, padronizada e de alta conscienciosidade para o ecossistema do Grupo Rede+.")

def extrair_texto_office(file, extensao):
    texto = ""
    try:
        if extensao == "docx":
            texto = docx2txt.process(file)
        elif extensao == "xlsx":
            wb = openpyxl.load_workbook(file, data_only=True)
            for sheet in wb.worksheets:
                texto += f"\nAba: {sheet.title}\n"
                for row in sheet.iter_rows(values_only=True):
                    row_text = " | ".join([str(cell) for cell in row if cell is not None])
                    if row_text:
                        texto += row_text + "\n"
        elif extensao == "pptx":
            prs = Presentation(file)
            for i, slide in enumerate(prs.slides):
                texto += f"\nSlide {i+1}:\n"
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        texto += shape.text + "\n"
    except Exception as e:
        texto = f"Erro na extração do arquivo Office: {str(e)}"
    return texto

uploaded_files = st.file_uploader(
    "Arraste ou selecione os arquivos para processamento e triagem estruturada no OneDrive", 
    type=["pdf", "png", "jpg", "jpeg", "docx", "xlsx", "pptx", "txt"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📂 **{len(uploaded_files)} arquivos carregados prontos para organização.**")
    
    if st.button("🚀 Executar Governança Documental"):
        client = genai.Client(api_key=st.secrets.get("GEMINI_API_KEY", "SUA_CHAVE_AQUI"))
        zip_buffer = io.BytesIO()
        arquivos_processados = []
        erros = []
        
        progresso_barra = st.progress(0)
        status_texto = st.empty()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for index, file in enumerate(uploaded_files):
                nome_original = file.name
                extensao = nome_original.split(".")[-1].lower()
                
                progresso_barra.progress((index + 1) / len(uploaded_files))
                status_texto.text(f"Analisando e higienizando ({index + 1}/{len(uploaded_files)}): {nome_original}")
                
                # PROMPT ATUALIZADO COM REGRAS DE DESEMPATE RÍGIDAS
                prompt_sistema = (
                    "Você é um motor de classificação puramente lógico e determinístico para o Prêmio Innpulse Bahia (Grupo Rede+).\n"
                    "Analise o arquivo e tome uma decisão baseada em regras estritas, sem espaço para variações interpretativas.\n\n"
                    "--- 1. TAXONOMIA DE PASTAS (CRITÉRIO DE DESEMPATE SELETO) ---\n"
                    "Escolha a pasta mais específica. Se um documento citar notas e transcrições ao mesmo tempo, a prioridade absoluta é o destino final do processo:\n"
                    "- '00_Gestao_Geral' (Apenas governança macro, OKRs, atribuições corporativas, POPs)\n"
                    "- '01_Planejamento' (Planos de ação, cronogramas estratégicos iniciais)\n"
                    "- '02_Metodologia_e_Regulamento' (Regulamentos estruturados, editais oficiais publicados)\n"
                    "- '03_Comunicacao' (Materiais visuais, e-mails de marketing, réguas de relacionamento)\n"
                    "- '04_Inscricoes' (Dados brutos obtidos logo após o fechamento dos formulários)\n"
                    "- '05_Curadoria_e_Avaliacao' (Listas de curadores, planilhas consolidadas de notas de triagem preliminar)\n"
                    "- '06_Entrevistas_e_Deliberacao' (Transcrições brutas ou editadas de bancas, atas de decisões colegiadas)\n"
                    "- '07_Finalistas_e_Vencedores' (Exclusivo para documentos com status 'Final' nomeando os ganhadores)\n"
                    "- '08_Cerimonia' (Roteiros de palco, checklists operacionais do dia do evento físico)\n"
                    "- '09_Pos_Premiaçao_e_Relatorios' (Relatórios pós-evento de impacto, feedbacks consolidados)\n"
                    "- '10_Memoria_Institucional' (Materiais históricos de edições de anos anteriores)\n"
                    "- '99_Arquivo_Morto' (Rascunhos sem uso, arquivos com a palavra 'Copiar' ou obsoletos)\n\n"
                    "--- 2. REGRAS DE NOMENCLATURA MATEMÁTICA ---\n"
                    "Gere o nome seguindo friamente o modelo: AAAA-MM-DD_Innpulse_Etapa_Tipo_Descricao_Status_VXX.[extensão]\n"
                    "- Data: Use a data explícita do texto. Se houver mais de uma, adote a mais recente. Se não houver, use estritamente: 2026-05-29.\n"
                    "- Descricao: 1 a 3 palavras em CamelCase descrevendo o assunto nuclear (ex: Geral, DeliberacaoEducacao, ListaCuradores).\n"
                    "- Status: Escolha estritamente entre 'Rascunho', 'Aprovado' ou 'Final'. Se o documento tiver marcas de revisão, use 'Rascunho'.\n"
                    "- VXX: Se o nome original continha números de versão (v2, v3), converta para 'V02', 'V03'. Caso contrário, use 'V01'.\n\n"
                    "--- FORMATO DE SAÍDA OBRIGATÓRIO (JSON BRUTO) ---\n"
                    "{\n"
                    "  \"subpasta\": \"NOME_DA_PASTA_SELECIONADA\",\n"
                    "  \"nome_sugerido_arquivo\": \"NOME_PADRONIZADO.ext\",\n"
                    "  \"palavras_chave_banco_dados\": [\"termo1\", \"termo2\"],\n"
                    "  \"resumo_conteudo\": \"Resumo executivo objetivo.\"\n"
                    "}"
                )
                
                if extensao in ["pdf", "png", "jpg", "jpeg"]:
                    bytes_data = file.getvalue()
                    mime_type = f"application/{extensao}" if extensao == "pdf" else f"image/{extensao}"
                    if extensao == "jpg": mime_type = "image/jpeg"
                    conteudo_para_ia = [types.Part.from_bytes(data=bytes_data, mime_type=mime_type), prompt_sistema]
                else:
                    texto_extraido = file.read().decode("utf-8") if extensao == "txt" else extrair_texto_office(file, extensao)
                    conteudo_para_ia = [f"Conteúdo textual para classificação:\n{texto_extraido}\n\n", prompt_sistema]
                
                try:
                    # CHAMADA CONFIGURADA COM TEMPERATURE = 0.0
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=conteudo_para_ia,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.0
                        )
                    )
                    
                    resultado = json.loads(response.text.strip())
                    subpasta_destino = resultado['subpasta'].strip()
                    nome_final = resultado['nome_sugerido_arquivo'].strip()
                    
                    caminho_no_zip = f"{subpasta_destino}/{nome_final}"
                    zip_file.writestr(caminho_no_zip, file.getvalue())
                    
                    arquivos_processados.append({
                        "original": nome_original,
                        "novo": nome_final,
                        "pasta": subpasta_destino,
                        "tags": ", ".join(resultado['palavras_chave_banco_dados']),
                        "resumo": resultado['resumo_conteudo']
                    })
                    
                except Exception as e:
                    zip_file.writestr(f"99_Arquivo_Morto/ERROS_PLUG_AND_PLAY/{nome_original}", file.getvalue())
                    erros.append({"arquivo": nome_original, "erro": str(e)})
        
        status_texto.success("✨ Processamento estável do lote finalizado!")
        progresso_barra.empty()
        
        st.write("---")
        st.subheader("📦 Download da Estrutura de Pastas do OneDrive")
        st.download_button(
            label="📥 BAIXAR REPOSITÓRIO HIGIENIZADO DO PRÊMIO (.ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Premio_Innpulse_Bahia_Organizado.zip",
            mime="application/zip",
            use_container_width=True
        )
        
        st.write("---")
        st.subheader("📊 Diagnóstico Semântico e Metadados de Governança")
        for item in arquivos_processados:
            with st.expander(f"📄 {item['original']} ➔ {item['novo']}"):
                st.markdown(f"**📍 Subpasta de Destino (OneDrive):** `{item['pasta']}`")
                st.markdown(f"**🔑 Indexadores Extraídos:** *{item['tags']}*")
                st.markdown(f"**📝 Resumo Analítico:** {item['resumo']}")
                
        if erros:
            st.warning("⚠️ Documentos direcionados para a contingência (/99_Arquivo_Morto/ERROS_PLUG_AND_PLAY/):")
            for err in erros:
                st.write(f"❌ Arquivo: {err['arquivo']} | Detalhes do erro: {err['erro']}")
