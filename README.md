# SNP to Function Data Extractor

SNP(Single Nucleotide Polymorphism) 위치 정보를 기반으로 유전자 데이터 및 기능 정보를 자동으로 추출하는 도구입니다.

## 개요

이 프로젝트는 Ensembl REST API와 NCBI Gene 데이터베이스를 활용하여 SNP 위치에 해당하는 유전자의 정보를 수집하고, 해당 유전자의 분자적 기능(Molecular Function)을 추출하여 Excel 파일로 저장합니다.

**대상 종**: Bos taurus (소)

## 주요 기능

- SNP 위치로부터 유전자 정보 자동 추출
- Ensembl API를 통한 유전자 ID 및 Gene Symbol 조회
- NCBI Gene 데이터베이스에서 분자적 기능 정보 수집
- 진행률 및 로그 실시간 표시
- Excel 형식으로 결과 자동 저장

## 설치 방법

### 필수 요구사항

- Python 3.6 이상
- pip (Python 패키지 관리자)

### 의존성 패키지

```bash
pip install requests beautifulsoup4 openpyxl
```

## 사용 방법

### 1. SNP 데이터 준비

`snps.json` 파일에 분석하려는 SNP 위치 정보를 입력합니다:

```json
{
    "snps": [
        "1: 110900379",
        "1: 118361465",
        "11: 55704515"
    ]
}
```

형식: `"염색체번호: 위치"`

### 2. 프로그램 실행

#### Windows (배치 파일 사용)

```bash
run_gene_automation.bat
```

이 방법은 자동으로 필요한 패키지를 설치하고 프로그램을 실행합니다.

#### Windows (PowerShell 사용)

```powershell
.\run_gene_automation.ps1
```

#### 직접 실행

```bash
python gene_automation.py
```

### 3. 결과 확인

프로그램 실행 후 다음 파일들이 생성됩니다:

- `gene_data_output.xlsx`: 추출된 유전자 데이터 및 기능 정보
- `gene_automation.log`: 실행 로그 파일

## 출력 형식

생성되는 Excel 파일은 다음 열을 포함합니다:

| 열 이름 | 설명 |
|---------|------|
| SNP | 입력된 SNP 위치 (염색체:위치) |
| GeneID | Ensembl Gene ID |
| Gene | 유전자 심볼 (Gene Symbol) |
| Function | 분자적 기능 설명 (Molecular Function) |

## 처리 과정

1. `snps.json`에서 SNP 위치 정보 로드
2. 각 SNP 위치에 대해:
   - Ensembl API로 해당 위치의 유전자 정보 조회
   - NCBI Gene ID 추출
   - NCBI Gene 웹페이지에서 분자적 기능 정보 스크래핑
3. 수집된 데이터를 Excel 파일로 저장
4. 진행률 및 결과를 로그에 기록

## 로깅

프로그램 실행 중 다음 정보가 로그로 기록됩니다:

- 전체 처리할 SNP 개수
- 각 SNP별 진행률 (%)
- Gene ID 및 Gene Symbol
- Function 정보 수집 결과
- 오류 및 경고 메시지

로그는 콘솔과 `gene_automation.log` 파일에 동시에 기록됩니다.

## 오류 처리

프로그램은 다음과 같은 상황을 자동으로 처리합니다:

- 유전자 정보가 없는 SNP 위치: 빈 값으로 기록
- Gene description이 없는 경우: Function 정보 없이 기록
- Function 정보를 찾을 수 없는 경우: 빈 값으로 기록

## 데이터 소스

- **Ensembl REST API**: https://rest.ensembl.org/
  - 유전자 위치 정보 및 Gene ID 조회
  - GO term 조회
- **NCBI Gene Database**: https://www.ncbi.nlm.nih.gov/gene/
  - 분자적 기능 정보 수집

## 주의사항

- 대용량의 SNP 데이터 처리 시 실행 시간이 오래 걸릴 수 있습니다
- 네트워크 연결이 필요합니다 (API 및 웹 스크래핑)
- API 요청 제한을 고려하여 대량 처리 시 주의가 필요합니다

## 라이선스

이 프로젝트는 교육 및 연구 목적으로 개발되었습니다.

## 문의

문제가 발생하거나 질문이 있으시면 이슈를 등록해주세요.
