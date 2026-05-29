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
                
                # PROMPT DO MOTOR CONFIGURADO COM A NOVA ESTRUTURA DO PRÊMIO INNPULSE BAHIA
                prompt_sistema = (
                    "Você é o motor de inteligência documental e governança do Prêmio Innpulse Bahia (Grupo Rede+).\n"
                    "Sua tarefa é ler o conteúdo do arquivo submetido, entender seu contexto operacional e retornar estritamente "
                    "um objeto JSON com a classificação correta e o nome perfeitamente padronizado.\n\n"
                    "--- 1. TAXONOMIA DE PASTAS (ESCOLHA UMA SUBPASTA SELETA) ---\n"
                    "Você deve alocar o arquivo em uma das seguintes pastas com base na finalidade descrita:\n"
                    "- '00_Gestao_Geral' (Governança, OKRs, atribuições, POPs, equipe)\n"
                    "- '01_Planejamento' (Definições estratégicas, planos de ação, cronogramas macros)\n"
                    "- '02_Metodologia_e_Regulamento' (Regras oficiais, editais, regulamentos PDF, critérios de avaliação)\n"
                    "- '03_Comunicacao' (Identidade visual, peças, logotipos, templates e réguas de e-mail)\n"
                    "- '04_Inscricoes' (Gestão de dados de inscritos, planilhas de inscritos, formulários)\n"
                    "- '05_Curadoria_e_Avaliacao' (Processo de triagem, notas, baremas, listas de curadores)\n"
                    "- '06_Entrevistas_e_Deliberacao' (Registros de bancas, transcrições, gravações, atas de deliberação)\n"
                    "- '07_Finalistas_e_Vencedores' (Kits de finalistas, declarações de vencedores - REQUER STATUS 'Final')\n"
                    "- '08_Cerimonia' (Produção do evento físico, roteiros, planta, checklists)\n"
                    "- '09_Pos_Premiaçao_e_Relatorios' (Feedbacks enviados, relatórios de impacto, resultados finais)\n"
                    "- '10_Memoria_Institucional' (Arquivo histórico, fotos, vídeos, documentos de anos anteriores - REQUER STATUS 'Final')\n"
                    "- '99_Arquivo_Morto' (Documentos obsoletos, rascunhos descartados, versões desatualizadas)\n\n"
                    "--- 2. REGRAS DE NOMENCLATURA PADRONIZADA ---\n"
                    "O nome gerado deve seguir RIGOROSAMENTE o modelo: AAAA-MM-DD_Innpulse_Etapa_Tipo_Descricao_Status_VXX.[extensão]\n"
                    "Componentes do nome:\n"
                    "- 'AAAA-MM-DD': Data extraída do documento. Caso nenhuma data cronológica clara seja identificada, use a data corrente: 2026-05-29.\n"
                    "- 'Innpulse': Prefixo fixo obrigatório.\n"
                    "- 'Etapa': Nome simplificado da etapa (ex: Gestao, Planejamento, Metodologia, Comunicacao, Inscricoes, Curadoria, Entrevistas, Finalistas, Cerimonia, PosPremio, Memoria, ArquivoMorto).\n"
                    "- 'Tipo': O tipo de documento (ex: Regulamento, Ata, Feedback, OKRs, Cronograma, Planilha, Transcricao, Roteiro, Relatorio, Template).\n"
                    "- 'Descricao': Descrição curta do assunto em formato CamelCase (ex: Geral, DeliberacaoEducacao, EmpresaX, ListaCuradores). Sem espaços.\n"
                    "- 'Status': Status de maturidade (ex: Rascunho, Aprovado, Final). Documentos na pasta '99_Arquivo_Morto' ou inacabados devem ser rotulados como 'Rascunho'. Pastas 07 e 10 aceitam apenas o status 'Final'.\n"
                    "- 'VXX': Controle sequencial de versão com dois dígitos (ex: V01, V02).\n\n"
                    "--- FORMATO DE SAÍDA OBRIGATÓRIO (JSON BRUTO) ---\n"
                    "Sua resposta deve conter exatamente esta estrutura de chaves:\n"
                    "{\n"
                    "  \"subpasta\": \"NOME_DA_PASTA_SELECIONADA\",\n"
                    "  \"nome_sugerido_arquivo\": \"NOME_PADRONIZADO.ext\",\n"
                    "  \"palavras_chave_banco_dados\": [\"termo1\", \"termo2\", \"termo3\"],\n"
                    "  \"resumo_conteudo\": \"Resumo executivo de alta fidelidade para busca semântica.\"\n"
                    "}"
                )
                
                if extensao in ["pdf", "png", "jpg", "jpeg"]:
                    bytes_data = file.getvalue()
                    mime_type = f"application/{extensao}" if extensao == "pdf" else f"image/{extensao}"
                    if extensao == "jpg": mime_type = "image/jpeg"
                    conteudo_para_ia = [types.Part.from_bytes(data=bytes_data, mime_type=mime_type), prompt_sistema]
                else:
                    texto_extraido = file.read().decode("utf-8") if extensao == "txt" else extrair_texto_office(file, extensao)
                    conteudo_para_ia = [f"Conteúdo textual do documento para classificação:\n{texto_extraido}\n\n", prompt_sistema]
                
                try:
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=conteudo_para_ia,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json"
                        )
                    )
                    
                    resultado = json.loads(response.text.strip())
                    subpasta_destino = resultado['subpasta'].strip()
                    nome_final = resultado['nome_sugerido_arquivo'].strip()
                    
                    # Montagem da estrutura física de diretórios dentro do ZIP de saída
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
                    # Rota de contingência para falhas: envia para o Arquivo Morto em subpasta de erros
                    zip_file.writestr(f"99_Arquivo_Morto/ERROS_PLUG_AND_PLAY/{nome_original}", file.getvalue())
                    erros.append({"arquivo": nome_original, "erro": str(e)})
        
        status_texto.success("✨ Processamento e triagem do lote finalizados!")
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
