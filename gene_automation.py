import requests
import json
from bs4 import BeautifulSoup
from openpyxl import Workbook
import logging
from datetime import datetime

species = 'bos_taurus'

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('gene_automation.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def get_gene_at_pos(species, chrom, pos):
    url = f'https://rest.ensembl.org/overlap/region/{species}/{chrom}:{pos}-{pos}'
    headers = {'Content-Type': 'application/json'}
    params = {'feature': 'gene'}
    resp = requests.get(url, headers=headers, params=params)
    if resp.status_code != 200 or not resp.json():
        return None
    return resp.json()[0]

def get_go_terms(gene_id):
    url = f'https://rest.ensembl.org/xrefs/id/{gene_id}'
    headers = {'Content-Type': 'application/json'}
    resp = requests.get(url, headers=headers)
    return [ref['primary_id'] for ref in resp.json() if ref['dbname'] == 'GO']

def get_go_description(go_id):
    url = f"https://rest.ensembl.org/ontology/id/{go_id}"
    headers = {'Content-Type': 'application/json'}
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        js = resp.json()
        # 1. label(이름)이 있으면 우선 반환
        label = js.get('label', '')
        # 2. label이 없으면 description(내용) 반환
        if label:
            return label
        desc = js.get('description', '')
        if desc:
            return desc
        # 3. 둘 다 없으면 GO term ID 반환
        return go_id
    return go_id


def load_positions_from_json(file_path):
    """snps.json 파일에서 위치 정보를 읽어옴"""
    with open(file_path, 'r') as f:
        snps = json.load(f).get('snps')

    if snps is None:
        return []
    # "11: 55704515" 형식을 파싱
    parsed_snps = []
    for snp in snps:
        chrom_value, position_value = snp.split(':')
        position_value = position_value.strip()
        parsed_snps.append({
            'chrom': chrom_value,
            'pos': position_value
        })
    return parsed_snps


snps_value = load_positions_from_json('snps.json')

# Excel 워크북 생성
wb = Workbook()
ws = wb.active
ws.title = "Gene Data"

# 헤더 작성
ws['A1'] = 'SNP'
ws['B1'] = 'GeneID'
ws['C1'] = 'Gene'
ws['D1'] = 'Function'

# snps_value가 빈 값이 아니면
if snps_value:
    total_count = len(snps_value)
    logger.info(f'=== 유전자 데이터 처리 시작 ===')
    logger.info(f'총 처리할 SNP 개수: {total_count}')

    row = 2  # 데이터는 2행부터 시작
    for i in range(len(snps_value)):
        chrom = snps_value[i]['chrom']
        pos = snps_value[i]['pos']
        snp_value = chrom + ':' + pos

        # 진행률 계산
        current_index = i + 1
        progress_percent = (current_index / total_count) * 100
        logger.info(f'[{current_index}/{total_count}] ({progress_percent:.1f}%) 처리 중: {snp_value}')

        # 1. 유전자 정보
        gene = get_gene_at_pos(species, chrom, pos)
        if not gene:
            logger.warning(f'  └─ 해당 위치에 유전자 정보 없음')
            # 빈 값으로 채우기
            ws[f'A{row}'] = snp_value
            ws[f'B{row}'] = ''
            ws[f'C{row}'] = ''
            ws[f'D{row}'] = ''
            row += 1
            continue
        # 2. GO term ID 얻기
        go_terms = get_go_terms(gene['id'])
        logger.info(f'  └─ Gene ID: {gene["id"]}, Gene Symbol: {gene.get("external_name", "-")}')

        # NCBI GeneID(Entrez)와 페이지 URL 예시 (소, CTNNA2: 527492)
        if not gene['description']:
            logger.warning(f'  └─ Gene description 없음')
            ws[f'A{row}'] = snp_value
            ws[f'B{row}'] = gene['id']
            ws[f'C{row}'] = gene.get('external_name', '-')
            ws[f'D{row}'] = ''
            row += 1
            continue
        gene_ncbi_id = gene['description'].split('Acc:')[1].replace(']', '')
        url = f'https://www.ncbi.nlm.nih.gov/gene/{gene_ncbi_id}'
        headers = {'User-Agent': 'Mozilla/5.0'}

        r = requests.get(url, headers=headers)
        soup = BeautifulSoup(r.text, 'html.parser')

        # Gene Ontology(Molecular function) 테이블 파싱
        # 테이블 내 각 'Function' (혹은 'Molecular function') 행의 텍스트를 추출
        functions = []
        for tr in soup.find_all('tr'):
            # Function 또는 Molecular function 필드에 해당하는 행 판별
            if tr.text.strip().lower().startswith(
                    'enables') or 'binding' in tr.text.lower() or 'structural molecule activity' in tr.text.lower():
                # 각 셀(열) 분리 (Function label만 추출)
                tds = tr.find_all('td')
                if tds:
                    # 첫 번째 셀 혹은 전체 텍스트로 Function 설명 추출
                    func_text = tds[0].text.strip()
                    functions.append(func_text)
                else:
                    # td가 없으면 전체 텍스트(행)를 넣음
                    functions.append(tr.text.strip())
        if functions:
            functionStr = ""
            for i in range(len(functions)):
                if i < len(functions) - 1:
                    f = functions[i] + ", "
                else:
                    f = functions[i]
                functionStr += f
            logger.info(f'  └─ Function 정보 {len(functions)}개 수집 완료')
            # Excel에 데이터 작성
            ws[f'A{row}'] = snp_value
            ws[f'B{row}'] = gene['id']
            ws[f'C{row}'] = gene.get('external_name', '-')
            ws[f'D{row}'] = functionStr
            row += 1
        else:
            logger.warning(f'  └─ Function 정보 없음')
            ws[f'A{row}'] = snp_value
            ws[f'B{row}'] = gene['id']
            ws[f'C{row}'] = gene.get('external_name', '-')
            ws[f'D{row}'] = ''
            row += 1
    logger.info(f'=== 모든 SNP 처리 완료 ===')
else:
    logger.error('snp.json 파일 형식에 오류가 있습니다.')
    wb.close()
    exit()

# Excel 파일 저장
wb.save('gene_data_output.xlsx')
logger.info(f'Excel 파일 저장 완료: gene_data_output.xlsx')
wb.close() # Excel 파일 비우기

