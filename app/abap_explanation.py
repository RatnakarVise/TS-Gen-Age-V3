import os
import re
import gc
from dotenv import load_dotenv
from langchain.prompts import ChatPromptTemplate
from langchain_community.vectorstores import Chroma
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.document_loaders import TextLoader
from langchain_core.messages import SystemMessage, HumanMessage

# Load environment variables
env_path = os.path.join(os.path.dirname(__file__), ".env")
if os.path.exists(env_path):
 load_dotenv(dotenv_path=env_path)
else:
    print(f"Warning: .env file not found at {env_path}. Environment variables may not be set correctly.")
    
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_API_KEY"] = os.getenv("LANGCHAIN_API_KEY")
os.environ["OPENAI_API_KEY"] = os.getenv("OPENAI_API_KEY")

# ✅ Step 1: Extract line-by-line explanation
def extract_abap_explanation(abap_code: str) -> str:
# Load RAG knowledge base
        try:
            rag_file_path = os.path.join(os.path.dirname(__file__), "rag_for_abap_explanation.txt")
            loader = TextLoader(file_path=rag_file_path, encoding="utf-8")
            documents = loader.load()

            # Create chunks for vector search
            text_splitter = RecursiveCharacterTextSplitter(chunk_size=20000, chunk_overlap=200)
            docs = text_splitter.split_documents(documents)

            embedding = OpenAIEmbeddings()
            vectorstore = Chroma.from_documents(docs, embedding)
            retriever = vectorstore.as_retriever()
            retrieved_docs = retriever.get_relevant_documents(abap_code)
            retrieved_context = "\n\n".join([doc.page_content for doc in retrieved_docs])
            del loader, documents, docs, vectorstore, embedding
            gc.collect()
            if not retrieved_context.strip():
                return "No relevant context found in RAG knowledge base."
            # Final prompt for TSD generation
            prompt_template = ChatPromptTemplate.from_template(
                                "You are an experienced SAP Techno-Functional Solution Architect. "
                                "Use the RAG context below to explain the ABAP code line-by-line in detail from both a technical and functional perspective.\n"
                                "Explain the given ABAP code line-by-line in detail from both a technical and functional perspective.\n"
                                "Output should be like below:\n"
                                "Selection Screen Parameters:\n(Go through all the selection scrren parameters and select-options in the ABAP code and explain them in detail THIS IS MANDATORY)\n"
                                "DATA SELECTION:Technically explain each and every data select query in the ABAP code.\n"
                                "Technical: Technical Explanation of selection screen PARAMETERS and SELECT-OPTIONS\n"
                                "Technical: Technical Explanation of line n\n"
                                "Functional: Functional Explanation of line n\n"
                                "Ensure to cover all lines in the ABAP code."
                                "RAG Context:\n{context}\n\n"
                                "ABAP Code:\n{abap_code}\n\n"

        )

            messages = prompt_template.format_messages(
            context=retrieved_context,
            abap_code=abap_code,
        )

            llm = ChatOpenAI(model="gpt-4.1", temperature=0)
            response = llm.invoke(messages)
            return response.content if hasattr(response, "content") else str(response)
        finally:
            del abap_code, retrieved_docs, retrieved_context, messages, prompt_template
            try:
                del llm, response  # in case they exist
            except:
                pass
            gc.collect()
