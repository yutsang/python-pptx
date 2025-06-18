import json, os, httpx
from openai import AzureOpenAI
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from azure.search.documents.indexes import SearchIndexClient
from azure.search.documents.indexes.models import SearchIndex, SimpleField, SearchableField, ComplexField
from azure.search.documents.models import VectorizedQuery, QueryType, SearchMode
import pandas as pd
from IPython.display import display, Markdown
from PIL import Image

def load_config(file_path):
    with open(file_path) as config_file:
        config_details = json.load(config_file)
    return config_details

def create_workflow_description(doc):
    str_content = f"title: {doc['title']}\nsteps:\n"
    for stp in doc['steps']:
        str_content += f"{stp['step_id']} {stp['step_name']} {stp['actor']} \n"
    return str_content

def initialize_ai_services(config_details):
    httpx_client = httpx.Client(verify=False)
    
    openai_api_base = config_details['OPENAI_API_BASE']
    openai_api_key = config_details['OPENAI_API_KEY']
    
    oai_client = AzureOpenAI(
        api_base=openai_api_base,
        api_key=openai_api_key,
        api_version=config_details['OPENAI_API_VERSION_COMPLETION'],
        client=httpx_client
    )
        
    search_client_configs = {
        "connection_verify": False,
        "headers": {
            "Host": f"{config_details['AZURE_SEARCH_SERVICE_NAME']}.search.windows.net",
        }
    }
    
    
    search_client = SearchClient(
        endpoint=f"https://{config_details['AZURE_AI_SEARCH_SERVICE_ENDPOINT']}/",
        index_name=config_details['AZURE_SEARCH_INDEX_NAME'],
        credential=AzureKeyCredential(config_details['AZURE_SEARCH_API_KEY']),
        **search_client_configs
    )
    
    return oai_client, search_client

def get_search_results(oai_client, search_client, user_query, openai_embedding_model, max_embedding_search_result):
    response = oai_client.embeddings.create(
        input=user_query,
        model=openai_embedding_model
    )
    embedding = response['data'][0]['embedding']
    vector_query = VectorizedQuery(
        vector=embedding,
        k_nearest_neighbors=max_embedding_search_result,
        fields = "embedding"
    )
    results = search_client.search(
        search_text=None,
        vector_queries = [vector_query],
        select=["title", "steps/step_id", "steps/step_name", "steps/actor", "steps/isDecision", "steps/nextStep", "img_file"]
    )
    documents = [doc for doc in results]
    df = pd.json_normalize(documents)
    
    context_content = "\n\n".join(create_workflow_description(doc) for doc in documents)
    
    return df, context_content

def generate_response(user_query, system_prompt, oai_client, context_content, openai_chat_model):
    user_message = {
        "role": "user",
        "content": user_query
    }
    
    conversation = [
        {
            "role": "system",
            "content": system_prompt
        },
        {
            "role": "assistant",
            "content": f"Context data: \n{context_content}"
        },
        {
            "role": "user",
            "content": user_query
        }
    ]
    
    response = oai_client.chat.completions.create(
        model=openai_chat_model,
        messages=conversation,
    )
    
    response_txt = response.choices[0].message.content
    
    return response_txt