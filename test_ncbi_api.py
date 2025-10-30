import requests
import json
import os

api_key = os.getenv('API_KEY')
species = 'GCF_000003055.6'
url = f'https://api.ncbi.nlm.nih.gov/datasets/v2/genome/accession/{species}/annotation_report'
headers = {'Content-Type': 'application/json'}
params = {'api_key': api_key}

print("Fetching NCBI API...")
resp = requests.get(url, headers=headers, params=params, timeout=30)
data = resp.json()

if 'genes' in data and len(data['genes']) > 0:
    print(f"Total genes: {len(data['genes'])}")

    # 첫 번째 유전자 구조 확인
    gene = data['genes'][0]
    print("\n=== First Gene Keys ===")
    print(list(gene.keys()))

    # go_terms 확인
    if 'go_terms' in gene:
        print("\n=== GO Terms Found ===")
        print(f"Number of GO terms: {len(gene['go_terms'])}")
        if len(gene['go_terms']) > 0:
            print("First GO term:", gene['go_terms'][0])
    else:
        print("\n=== NO go_terms field ===")

    # 파일로 저장
    with open('ncbi_gene_sample.json', 'w', encoding='utf-8') as f:
        json.dump(gene, f, indent=2, ensure_ascii=False)
    print("\nSaved to ncbi_gene_sample.json")
