import os

async def get_doc_data(embeddings):
    with open(f'{os.getcwd()}/src/indexers/data/Contoso_Electronics_PerkPlus_Program.md', 'r') as file:
        raw_description1 = file.read()
    doc1 = {
        "docId": "1",
        "docTitle": "Contoso_Electronics_PerkPlus_Program",
        "description": raw_description1,
        "descriptionVector": await get_embedding_vector(raw_description1, embeddings=embeddings),
    }
    
    with open(f'{os.getcwd()}/src/indexers/data/Contoso_Electronics_Company_Overview.md', 'r') as file:
        raw_description2 = file.read()
    doc2 = {
        "docId": "2",
        "docTitle": "Contoso_Electronics_Company_Overview",
        "description": raw_description2,
        "descriptionVector": await get_embedding_vector(raw_description2, embeddings=embeddings),
    }
    
    with open(f'{os.getcwd()}/src/indexers/data/Contoso_Electronics_Plan_Benefits.md', 'r') as file:
        raw_description3 = file.read()
    doc3 = {
        "docId": "3",
        "docTitle": "Contoso_Electronics_Plan_Benefits",
        "description": raw_description3,
        "descriptionVector": await get_embedding_vector(raw_description3, embeddings=embeddings),
    }

    return [doc1, doc2, doc3]


async def get_embedding_vector(text: str, embeddings):
    result = await embeddings.create_embeddings(text)
    if (result.status != 'success' or not result.output):
        if result.status == 'error':
            raise Exception(f"Failed to generate embeddings for description: <{text[:200]+'...'}>\n\nError: {result.output}")
        raise Exception(f"Failed to generate embeddings for description: <{text[:200]+'...'}>")
    
    return result.output[0]