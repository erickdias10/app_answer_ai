# ---------------------------
# Importação de Bibliotecas
# ---------------------------
import re
from PyPDF2 import PdfReader
import unicodedata
from docx import Document
from docx.shared import Pt
import os
import joblib
import streamlit as st
from langchain.chains.question_answering import load_qa_chain
from langchain.llms import OpenAI
from langchain.document_loaders import TextLoader
from langchain.text_splitter import CharacterTextSplitter
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import FAISS

# ---------------------------
# Configuração do LangChain
# ---------------------------
OPENAI_API_KEY = "sua_openai_api_key_aqui"  # Substitua pela sua chave da OpenAI

def load_llm():
    return OpenAI(model_name="gpt-3.5-turbo", temperature=0, openai_api_key=OPENAI_API_KEY)

# ---------------------------
# Funções de Processamento de Texto
# ---------------------------
def normalize_text(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r"\s{2,}", " ", text)  # Remove múltiplos espaços
    return text.strip()

def corrigir_texto(texto):
    substituicoes = {
        'Ã©': 'é',
        'Ã§Ã£o': 'ção',
        'Ã³': 'ó',
        'Ã': 'à',
    }
    for errado, correto in substituicoes.items():
        texto = texto.replace(errado, correto)
    return texto

def extract_text_with_pypdf2(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        text = corrigir_texto(normalize_text(text))
        return text.strip()
    except Exception as e:
        print(f"Erro ao processar PDF {pdf_path}: {e}")
        return ''

# ---------------------------
# Funções de Extração de Dados
# ---------------------------
def extract_process_number(file_name):
    """
    Extrai o número do processo a partir do nome do arquivo, removendo "SEI" e preservando o restante.
    """
    base_name = os.path.splitext(file_name)[0]  # Remove a extensão
    if base_name.startswith("SEI"):
        base_name = base_name[3:].strip()  # Remove "SEI"
    return base_name

# ---------------------------
# Função de Perguntas e Respostas
# ---------------------------
def create_qa_chain(text):
    """
    Cria um pipeline de perguntas e respostas com LangChain.
    """
    llm = load_llm()
    text_splitter = CharacterTextSplitter(chunk_size=1000, chunk_overlap=0)
    documents = text_splitter.split_text(text)
    
    # Criar embeddings para busca
    embeddings = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)
    vectorstore = FAISS.from_texts(documents, embeddings)
    
    return vectorstore, llm

def answer_question(vectorstore, llm, question):
    """
    Responde a uma pergunta usando LangChain e busca nos embeddings.
    """
    retriever = vectorstore.as_retriever()
    docs = retriever.get_relevant_documents(question)
    
    chain = load_qa_chain(llm, chain_type="stuff")
    answer = chain.run(input_documents=docs, question=question)
    return answer

# ---------------------------
# Interface Streamlit
# ---------------------------
st.title("Sistema de Perguntas e Respostas Baseado em Arquivos")

uploaded_file = st.file_uploader("Envie um arquivo PDF", type="pdf")

if uploaded_file:
    try:
        # Extrai o número do processo a partir do nome do arquivo
        file_name = uploaded_file.name
        numero_processo = extract_process_number(file_name)

        # Extrai o texto do PDF
        text = extract_text_with_pypdf2(uploaded_file)
        if text:
            st.success(f"Texto extraído com sucesso! Número do processo: {numero_processo}")
            
            # Configura o sistema de perguntas e respostas
            st.write("Configurando o sistema de perguntas e respostas...")
            vectorstore, llm = create_qa_chain(text)
            st.success("Sistema configurado! Você pode começar a fazer perguntas.")
            
            # Campo para perguntas do usuário
            question = st.text_input("Faça sua pergunta sobre o conteúdo do arquivo:")
            if question:
                with st.spinner("Processando a resposta..."):
                    answer = answer_question(vectorstore, llm, question)
                    st.write("Resposta:")
                    st.success(answer)
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
