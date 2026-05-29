import streamlit as st
from google import genai
from google.genai import types
import json
import io
import zipfile

# Importações para extração de arquivos Office
import docx2txt
import openpyxl
from pptx import Presentation

st.set_page_config(page_title="Organizador Inteligente", page_icon="📁", layout="wide")
st.title("Organizador Inteligente de Arquivos em Lote 📦")
st.write("Submeta múltiplos arquivos. A IA vai renomeá-los e organizá-los em uma estrutura de pastas compactada para você.")

# 1. Função para extrair texto de arquivos do Office
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
        texto = f"Erro ao extrair texto: {str(e)}"
    return texto

# 2. Upload de Múltiplos Arquivos (accept_multiple_files=True)
uploaded_files = st.file_uploader(
    "Arraste ou selecione todos os arquivos que deseja organizar", 
    type=["pdf", "png", "jpg", "jpeg", "docx", "xlsx", "pptx", "txt"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📂 **{len(uploaded_files)} arquivos carregados.**")
    
    # Botão para iniciar o processamento em lote
    if st.button("🚀 Iniciar Organização Inteligente"):
        
        # Configuração do cliente Gemini (pegando a chave dos Secrets do Streamlit)
        client = genai.Client(api_key=st.secrets.get("GEMINI_API_KEY", "SUA_CHAVE_AQUI"))
        
        # Criar um buffer na memória para o arquivo ZIP
        zip_buffer = io.BytesIO()
        
        # Listas para exibir o resumo na tela depois
        arquivos_processados = []
        erros = []
        
        # Barra de progresso visual do Streamlit
        progresso_barra = st.progress(0)
        status_texto = st.empty()
        
        # Criando o arquivo ZIP dentro do buffer
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            
            for index, file in enumerate(uploaded_files):
                nome_original = file.name
                extensao = nome_original.split(".")[-1].lower()
                
                # Atualiza os componentes de progresso
                percentual = (index + 1) / len(uploaded_files)
                progresso_barra.progress(percentual)
                status_texto.text(f"Analisando ({index + 1}/{len(uploaded_files)}): {nome_original}")
                
                # Monta o prompt padrão para a IA
                prompt = (
                    f"Analise o conteúdo deste arquivo (Nome original: {nome_original}). "
                    "Defina a subpasta ideal para ele (ex: Financeiro, RH, Contratos, Imagens, Notas_Fiscais) "
                    "e crie um novo nome padronizado no formato AAAAMMDD_NomeDescritivo.[extensão]. "
                    "Gere o nome baseado no assunto principal do documento. "
                    "Responda ESTREITAMENTE em formato JSON, sem marcações markdown de código: "
                    '{"pasta": "NOME_DA_PASTA", "novo_nome": "NOVO_NOME_DO_ARQUIVO"}'
                )
                
                # Decide a estratégia de leitura baseada no formato do arquivo
                if extensao in ["pdf", "png", "jpg", "jpeg"]:
                    bytes_data = file.getvalue()
                    mime_type = f"application/{extensao}" if extensao == "pdf" else f"image/{extensao}"
                    if extensao == "jpg": mime_type = "image/jpeg"
                    
                    conteudo_para_ia = [
                        types.Part.from_bytes(data=bytes_data, mime_type=mime_type),
                        prompt
                    ]
                else:
                    if extensao == "txt":
                        texto_extraido = file.read().decode("utf-8")
                    else:
                        texto_extraido = extrair_texto_office(file, extensao)
                        
                    conteudo_para_ia = [
                        f"Texto extraído do documento:\n{texto_extraido}\n\n",
                        prompt
                    ]
                
                # Chamada da API do Gemini
                try:
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=conteudo_para_ia
                    )
                    
                    # Trata a resposta em JSON
                    resultado = json.loads(response.text.strip())
                    pasta_destino = resultado['pasta'].strip().replace("/", "_") # Evita quebras indesejadas
                    novo_nome = resultado['novo_nome'].strip()
                    
                    # O truque do ZIP: Definir o caminho como "Pasta/Subpasta/Arquivo.ext" 
                    # faz o sistema operacional criar as pastas automaticamente ao extrair!
                    caminho_no_zip = f"{pasta_destino}/{novo_nome}"
                    
                    # Salva o arquivo original com o novo nome dentro da pasta correspondente no ZIP
                    zip_file.writestr(caminho_no_zip, file.getvalue())
                    
                    arquivos_processados.append({
                        "original": nome_original,
                        "novo": novo_nome,
                        "pasta": pasta_destino
                    })
                    
                except Exception as e:
                    # Se um arquivo falhar, ele salva em uma pasta de erros dentro do zip para não travar o lote
                    zip_file.writestr(f"ERROS_PROCESSAMENTO/{nome_original}", file.getvalue())
                    erros.append({"arquivo": nome_original, "erro": str(e)})
        
        # Finaliza os indicadores de progresso
        status_texto.success("✨ Processamento de todos os arquivos concluído!")
        progresso_barra.empty()
        
        # 3. O Botão de Download do ZIP pronto
        st.write("---")
        st.subheader("📦 Baixe sua estrutura organizada")
        
        st.download_button(
            label="📥 BAIXAR PASTA ORGANIZADA (.ZIP)",
            data=zip_buffer.getvalue(),
            file_name="arquivos_organizados.zip",
            mime="application/zip",
            use_container_width=True
        )
        
        # 4. Relatório visual na tela para o usuário conferir
        st.write("---")
        st.subheader("📊 Resumo da Organização")
        
        if arquivos_processados:
            st.write("#### ✅ Arquivos Alocados com Sucesso:")
            for item in arquivos_processados:
                st.markdown(f"🔹 **{item['original']}** ➔ Mover para `/{item['pasta']}/` como *`{item['new'] if 'new' in item else item['novo']}`*")
                
        if erros:
            st.warning("#### ⚠️ Alguns arquivos apresentaram problemas e foram enviados para a pasta `/ERROS_PROCESSAMENTO/`:")
            for err in erros:
                st.write(f"❌ {err['arquivo']} (Motivo: {err['erro']})")
