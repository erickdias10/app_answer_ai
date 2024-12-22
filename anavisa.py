# ---------------------------
# Importação de Bibliotecas
# ---------------------------
import re
from PyPDF2 import PdfReader
import unicodedata
import joblib
import streamlit as st
from sklearn.feature_extraction.text import TfidfVectorizer

# ---------------------------
# Configuração do Modelo Personalizado
# ---------------------------
VECTOR_PATH = r"C:\Users\erick\OneDrive\Área de Trabalho\Jupyter Notebook\ANVISA_Projeto02\vectorizer.pkl"  # Caminho para o vetorizador salvo
MODEL_PATH = r"C:\\Users\\erick\\OneDrive\\Área de Trabalho\\Jupyter Notebook\\ANVISA_Projeto02\\model.pkl"  # Caminho para o modelo salvo

def load_model(vectorizer_path, model_path):
    """
    Carrega o vetorizador e o modelo.
    """
    vectorizer = joblib.load(vectorizer_path)
    model = joblib.load(model_path)
    return vectorizer, model

def predict_answer(question, text, vectorizer, model):
    """
    Faz a predição com base na pergunta e no texto fornecido.
    """
    try:
        # Pré-processar texto e pergunta
        input_text = f"{text} [SEP] {question}"  # Concatenar texto e pergunta
        input_vectorized = vectorizer.transform([input_text])

        # Fazer a predição
        prediction = model.predict(input_vectorized)
        return prediction[0]  # Retorna a resposta
    except Exception as e:
        return f"Erro na predição: {e}"

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
# Interface Streamlit
# ---------------------------
st.title("Olá, eu sou a AnaVisa! \n A IA da Anvisa. \n Como posso lhe ajudar hoje?")

uploaded_file = st.file_uploader("Envie um arquivo PDF", type="pdf")

if uploaded_file:
    try:
        # Extrai o texto do PDF
        text = extract_text_with_pypdf2(uploaded_file)
        if text:
            st.success("Texto extraído com sucesso! Você pode começar a fazer perguntas.")

            # Carrega o modelo e o vetorizador
            st.write("Carregando o modelo...")
            vectorizer, model = load_model(VECTOR_PATH, MODEL_PATH)
            st.success("Modelo carregado com sucesso!")

            # Campo para perguntas do usuário
            question = st.text_area("Faça sua pergunta sobre o conteúdo do arquivo:")
            if question:
                with st.spinner("Processando a resposta..."):
                    answer = predict_answer(question, text, vectorizer, model)
                    st.write("Resposta:")
                    st.success(answer)
    except FileNotFoundError as e:
        st.error(f"Arquivo não encontrado: {e.filename}")
    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar o modelo: {e}")
