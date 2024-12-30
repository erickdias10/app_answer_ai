import streamlit as st
import logging
import spacy
import re
import os
from PyPDF2 import PdfReader
from transformers import pipeline
from pdf2image import convert_from_bytes  # para converter PDF em imagens
import pytesseract  # para OCR

# Para embeddings e busca vetorial
import faiss
import numpy as np
from sentence_transformers import SentenceTransformer

# ---------------------------------------------------------------------
# DICAS:
# 1) Se o PDF for digital, PyPDF2 extrai o texto diretamente.
# 2) Se não extrair nada (ou extrair texto vazio), partimos para OCR.
# 3) Para OCR funcionar, instale tesseract-ocr no SISTEMA (não é só pip).
# ---------------------------------------------------------------------

# Configuração de logs
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Carregar o modelo spaCy (português)
try:
    nlp = spacy.load("pt_core_news_lg")
    logging.info("Modelo spaCy carregado com sucesso!")
except OSError:
    st.error(
        "Modelo spaCy não encontrado. "
        "Por favor, instale o modelo 'pt_core_news_lg' com:\n\n"
        "pip install spacy\n"
        "python -m spacy download pt_core_news_lg\n\n"
        "Verifique se está no mesmo ambiente virtual."
    )
    raise

# Carregar o modelo de embeddings (Sentence-Transformers)
try:
    embedder = SentenceTransformer('multi-qa-MiniLM-L6-cos-v1')
    # Ou outro modelo, por ex: 'sentence-transformers/all-mpnet-base-v2'
    logging.info("Modelo de embeddings carregado com sucesso!")
except Exception as e:
    st.error(f"Erro ao carregar modelo de embeddings: {e}")
    raise

def clean_text(text: str) -> str:
    """
    Faz uma limpeza no texto para remover caracteres invisíveis,
    quebras de linha e espaços excessivos.
    """
    # Remove caracteres zero-width (por exemplo, \u200B)
    text = re.sub(r'[\u200B-\u200D\uFEFF]', '', text)
    # Substitui quebras de linha por espaço
    text = text.replace('\n', ' ').replace('\r', ' ')
    # Remove múltiplos espaços
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_text_from_pdf(pdf_file) -> str:
    """
    1) Tenta extrair texto com PyPDF2 (PDF digital).
    2) Se PyPDF2 não extrair texto ou extrair muito pouco, converte cada página
       para imagem e usa OCR (pytesseract) para extrair texto.
    Retorna a string com o texto extraído (ou "" se der erro).
    """
    try:
        # 1) Tenta via PyPDF2
        reader = PdfReader(pdf_file)
        all_text = []
        for page in reader.pages:
            txt = page.extract_text() or ""
            all_text.append(txt)

        extracted_text = "\n".join(all_text).strip()

        # Se extraímos pouco ou nenhum texto, tentamos OCR
        if len(extracted_text) < 30:  # heurística simples
            logging.info("Tentando OCR, pois extraímos pouco texto via PyPDF2...")
            pdf_file.seek(0)  # reposiciona o ponteiro do arquivo
            images = convert_from_bytes(pdf_file.read())
            ocr_text = []
            for image in images:
                text_page = pytesseract.image_to_string(image, lang='por')  # OCR em português
                ocr_text.append(text_page)
            extracted_text = "\n".join(ocr_text)

        return clean_text(extracted_text)

    except Exception as e:
        logging.error(f"Erro ao processar PDF: {e}")
        return ""

def chunk_text(text, chunk_size=500, overlap=50):
    """
    Divide o texto em partes (chunks) de tamanho 'chunk_size',
    com uma sobreposição de 'overlap' caracteres para não perder contexto.
    Retorna uma lista de strings.
    """
    text = text.strip()
    words = text.split()
    chunks = []
    start = 0

    while start < len(words):
        end = start + chunk_size
        chunk = " ".join(words[start:end])
        chunks.append(chunk)
        start += (chunk_size - overlap)

    return chunks

def build_vector_index(chunks):
    """
    Cria e retorna um índice vetorial (FAISS) a partir de uma lista de chunks.
    Também retorna os embeddings e a lista de chunks para consultas posteriores.
    """
    # Gera embeddings para cada chunk
    embeddings = embedder.encode(chunks)
    embeddings = np.array(embeddings, dtype="float32")

    # Cria o índice FAISS
    index = faiss.IndexFlatL2(embeddings.shape[1])  
    index.add(embeddings)

    return index, embeddings

def search_similar_chunks(question, index, embeddings, chunks, top_k=3):
    """
    Faz a busca vetorial (similaridade) com base na pergunta, retornando
    os 'top_k' trechos mais relevantes.
    """
    question_embedding = embedder.encode([question])
    question_embedding = np.array(question_embedding, dtype="float32")

    # Busca no FAISS
    distances, indices = index.search(question_embedding, top_k)

    # Retorna os chunks relevantes
    results = []
    for dist, idx in zip(distances[0], indices[0]):
        results.append((chunks[idx], dist))

    return results

def answer_question_with_embeddings(text, question, index, embeddings, chunks):
    """
    Usa a busca semântica (via embeddings) para encontrar trechos
    relevantes e depois tenta extrair a resposta.
    """
    # 1) Busca trechos relevantes
    top_chunks = search_similar_chunks(question, index, embeddings, chunks, top_k=3)

    # 2) Concatena os trechos em um "contexto"
    context = "\n".join([t[0] for t in top_chunks])

    # 3) Tenta extrair com regex ou alguma heurística específica
    # Exemplo: para "CNPJ", tentamos capturar via regex
    question_lower = question.lower()
    if "cnpj" in question_lower:
        # Regex simples para CNPJ: XX.XXX.XXX/XXXX-XX ou variações
        cnpjs = re.findall(r"\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}", context)
        if cnpjs:
            return "; ".join(cnpjs)

        # Tenta outra variação sem pontuação
        cnpjs_alt = re.findall(r"\d{14}", context)
        if cnpjs_alt:
            return "; ".join(cnpjs_alt)

        return "Nenhum CNPJ encontrado nos trechos relevantes."

    # Se for outra pergunta, podemos usar o spaCy para procurar
    doc = nlp(context)
    sentences = [sent.text for sent in doc.sents if question_lower in sent.text.lower()]

    if sentences:
        # Retorna as sentenças que contêm a pergunta (heurística simples)
        return "\n".join(sentences)
    else:
        return "Não foi possível encontrar uma resposta relevante nos trechos."

def generate_summary(text: str) -> str:
    """
    Gera um resumo do texto usando o modelo T5-small (transformers).
    """
    try:
        summarizer = pipeline("summarization", model="t5-small", tokenizer="t5-small")
        # Limitar o texto para evitar problemas de sequência muito longa
        text = clean_text(text)
        if len(text) > 1024:
            text = text[:1024]
        summary = summarizer(text, max_length=200, min_length=50, do_sample=False)
        if summary:
            return summary[0]['summary_text']
        return "Não foi possível gerar um resumo para o texto fornecido."
    except Exception as e:
        logging.error(f"Erro ao gerar resumo: {e}")
        return f"Erro ao gerar resumo: {e}"

# -----------------------------------------------------------------------------
# Interface Streamlit
# -----------------------------------------------------------------------------
st.title("Pergunte à AnaVisa")
st.markdown("Envie um arquivo PDF do processo (digital ou escaneado) para fazer perguntas ou gerar um resumo.")

uploaded_file = st.file_uploader("Envie um arquivo PDF", type=["pdf"])

if uploaded_file:
    try:
        # Extrai texto
        raw_text = extract_text_from_pdf(uploaded_file)
        if not raw_text or len(raw_text) < 10:
            st.error("Não foi possível extrair texto do arquivo PDF. "
                     "Verifique se é um PDF válido ou tente outro arquivo.")
        else:
            st.success("Arquivo carregado com sucesso!")

            # (Opcional) Exiba o texto extraído para debug
            # st.write("Texto Extraído do PDF (debug):")
            # st.write(raw_text)

            # ---------------------------
            # Construir índice vetorial
            # ---------------------------
            st.info("Construindo índice semântico para buscas...")
            chunks = chunk_text(raw_text, chunk_size=100, overlap=20)
            index, embeddings = build_vector_index(chunks)
            st.success("Índice vetorial criado!")

            # Perguntas interativas
            st.markdown("### Faça sua pergunta:")
            question = st.text_area("Digite sua pergunta:")

            if st.button("Fazer pergunta"):
                # Usa a função de QA com embeddings
                answer = answer_question_with_embeddings(raw_text, question, index, embeddings, chunks)
                st.success(f"Resposta: {answer}")

            # Geração de resumo
            if st.button("Gerar Resumo"):
                summary = generate_summary(raw_text)
                st.markdown("### Resumo do Processo:")
                st.success(summary)

    except Exception as e:
        st.error(f"Erro ao carregar ou processar o arquivo: {e}")
