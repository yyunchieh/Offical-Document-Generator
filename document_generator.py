import streamlit as st
import os
from langchain.chains import LLMChain
from langchain.prompts import PromptTemplate
from langchain.chat_models import ChatOpenAI
import openai
from io import BytesIO
from docx import Document

# Load API key
API_path = r"C:\Users\4019-tjyen\Desktop\API.txt"
with open(API_path,"r") as file:
    openapi_key = file.read().strip()
    
os.environ['OPENAI_API_KEY'] = openapi_key
openai.api_key = openapi_key

# Initialize Chat Model
chat_model = ChatOpenAI(model="gpt-4o", temperature=0.5)

# Define the templates
doc_template = """
標題:{title}

聯絡資訊: {contact_info}

受文者: {recipient}

發文字號: {doc_number}

主旨: {subject}

說明: {description}

"""

prompt_template = """
You are a helpful assistant tasked with creating a formal document based on the following description and details:
Request: {description}

Example: {template}

Please generate a clean, professional document in the requested format.
"""

# Initialize LangChain
prompt = PromptTemplate(input_variables=["description", "template"], template=prompt_template)
chain = LLMChain(llm=chat_model, prompt=prompt)

# Streamlit App
st.title("Automated Document Generator")

# Input
title = st.text_input("標題:", value = "ooo函")
contact_info = st.text_area("聯絡資訊:", value="地址：oo市oo區oo路\n承辦人：張小姐\n電話：(02)1234-5678\nEmail：example@example.com")
recipient = st.text_input("受文者:", value="oo部oo處")
doc_number = st.text_input("發文字號:", value="公司簡稱-113-001")
subject = st.text_input("主旨:", value="關於公司年度計畫備查")
description = st.text_area("說明內容:", value="1. 敬請參閱本公司年度計畫。\n2. 如需進一步資訊，請隨時聯繫我們。")

# Upload files as references
#uploaded_file = st.file_uploader("Upload files as references(optional)", type=["jpg", "png", "pdf", "docx"])
#if uploaded_file is not None:
#    try:
#        if 


# Button to generate document
if st.button("Generate Document"):
    formatted_input = {
        "title": title,
        "contact_info": contact_info,
        "recipient": recipient,
        "doc_number": doc_number,
        "subject": subject,
        "description": description
    }

    result = chain.run({"description": description, "template": doc_template})
    st.subheader("Generated Document:")
    st.text_area("Generated Content:", value=result, height=300)
            
    # Save as Word file
    doc = Document()
    doc.add_paragraph(result)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
                
    #Download button for word file
    st.download_button(
        label="Download Document as Word",
        data=buffer,
        file_name="generated_document.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        