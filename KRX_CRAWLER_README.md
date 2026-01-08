# KRX ESG 포털 크롤러 (로컬 PC용)

## 개요
KRX ESG 포털(https://esg.krx.co.kr)에서 기업별 ESG 등급을 수집하는 크롤러입니다.

**주의**: Google Colab에서는 IP 차단으로 실행 불가. 반드시 로컬 PC에서 실행하세요.

## 수집 데이터
| 평가기관 | 등급 체계 | 수집 항목 |
|---------|----------|----------|
| KCGS | S, A+, A, B+, B, C, D | ESG 통합, E, S, G |
| 서스틴베스트 | AA, A, BB, B, C, D, E | ESG 통합 |
| 한국ESG연구소 | S, A+, A, B+, B, C, D | ESG 통합 |
| MSCI | AAA ~ CCC | ESG 통합 |
| S&P | 0-100 점수 | ESG Score |

## 설치

### 1. Python 패키지 설치
```bash
pip install playwright pandas openpyxl tqdm
```

### 2. Playwright 브라우저 설치
```bash
playwright install chromium
```

## 사용법

### 기본 실행 (테스트 모드)
```bash
python krx_esg_crawler.py
```

### KCGS 파일 기반 전체 크롤링
```bash
python krx_esg_crawler.py --input kcgs_data.xlsx --output krx_esg_result.csv
```

### 옵션

| 옵션 | 설명 | 예시 |
|-----|------|-----|
| `--input`, `-i` | KCGS Excel 파일 경로 | `--input data.xlsx` |
| `--output`, `-o` | 출력 CSV 파일명 | `--output result.csv` |
| `--max`, `-m` | 최대 크롤링 기업 수 | `--max 100` |
| `--headless` | 브라우저 숨김 모드 | `--headless` |
| `--resume` | 체크포인트에서 재개 | `--resume` |

### 예시

```bash
# 처음 10개 기업만 테스트
python krx_esg_crawler.py --input kcgs_data.xlsx --max 10

# 브라우저 보이면서 실행 (디버깅용)
python krx_esg_crawler.py --input kcgs_data.xlsx --max 5

# 헤드리스 + 전체 크롤링
python krx_esg_crawler.py --input kcgs_data.xlsx --headless

# 중단된 곳에서 재개
python krx_esg_crawler.py --input kcgs_data.xlsx --resume
```

## 출력 파일

1. **krx_esg_result.csv**: 크롤링 결과
   - 기업명, 기업코드, 연도, KCGS_ESG, KCGS_E, KCGS_S, KCGS_G, 서스틴베스트_ESG, ...

2. **krx_esg_result_failed.csv**: 실패한 기업 목록
   - 기업명, 기업코드, 사유

3. **krx_checkpoint.json**: 체크포인트 (자동 저장)
   - 10개 기업마다 자동 저장, `--resume`으로 재개 가능

## 주의사항

1. **크롤링 속도**: 차단 방지를 위해 기업당 10-30초 소요
2. **전체 소요 시간**: 1,500개 기업 기준 약 6-12시간
3. **IP 차단 시**: VPN 사용 또는 몇 시간 후 재시도
4. **체크포인트**: 중단되어도 `--resume`으로 이어서 실행 가능

## 문제 해결

### IP 차단 메시지가 나올 때
```
❌ IP가 차단되었습니다!
```
→ VPN을 사용하거나 1-2시간 후 재시도

### 브라우저 설치 오류
```bash
playwright install chromium --with-deps
```

### 한글 인코딩 문제
CSV 파일은 `utf-8-sig` 인코딩으로 저장됩니다. Excel에서 바로 열 수 있습니다.

## 라이선스
학술 연구 목적으로만 사용하세요. KRX 이용약관을 준수해주세요.
