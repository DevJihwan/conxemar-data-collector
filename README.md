# 🏢 Conxemar Data Collector & Ambiente Formatter

Conxemar 박람회 데이터를 수집하고 Ambiente 2025 양식으로 변환하는 완전 자동화된 Node.js 도구입니다.

## 🎯 프로젝트 개요

이 프로젝트는 **3단계 워크플로우**로 구성됩니다:
1. **데이터 수집**: Conxemar 박람회 웹사이트에서 기업 데이터 자동 수집
2. **양식 분석**: Ambiente 2025 기준 양식 구조 분석
3. **데이터 변환**: 수집된 데이터를 기준 양식에 맞춰 변환

## 📋 핵심 기능

### 🔍 1. 데이터 수집기 (conxemar_scraper.js)
- ✅ Conxemar API 호출을 통한 자동 데이터 수집
- ✅ 페이지네이션 처리로 전체 데이터 수집 (677개+ 기업)
- ✅ JSON 및 Excel 형태로 데이터 저장
- ✅ 실시간 수집 진행상황 표시
- ✅ 국가별/업종별 통계 자동 생성

### 📊 2. 양식 분석기 (template_analyzer.js)
- ✅ Ambiente 2025 기준 양식 파일 자동 분석
- ✅ 헤더 구조 및 데이터 타입 분석
- ✅ 필드 매핑 제안 자동 생성
- ✅ 코드 템플릿 자동 생성

### 🔄 3. 양식 변환기 (ambiente_formatter.js)
- ✅ 스마트 필드 매핑 (영문 헤더 ↔ 한글 데이터)
- ✅ 실제 기준 양식에 맞춘 자동 변환
- ✅ 데이터 정규화 및 품질 개선
- ✅ 변환 통계 및 품질 지표 제공

## 🚀 빠른 시작

### 1. 프로젝트 설정
```bash
# 리포지토리 클론
git clone https://github.com/DevJihwan/conxemar-data-collector.git
cd conxemar-data-collector

# 의존성 설치
npm install
```

### 2. 단계별 실행

#### Step 1: 데이터 수집
```bash
npm run collect
```
> 결과: `conxemar_companies.json`, `conxemar_companies.xlsx` 생성

#### Step 2: 기준 양식 분석
```bash
# Ambiente 2025 Exhibitor 리스트.xlsx 파일을 프로젝트 폴더에 복사한 후
npm run analyze
```
> 결과: 기준 양식의 정확한 헤더 구조 확인

#### Step 3: 양식 변환
```bash
npm run format
```
> 결과: `conxemar_ambiente_format.xlsx` 생성

## 📁 파일 구조

```
conxemar-data-collector/
├── conxemar_scraper.js          # 데이터 수집기
├── template_analyzer.js         # 양식 분석기 
├── ambiente_formatter.js        # 양식 변환기
├── package.json                 # 프로젝트 설정
├── README.md                    # 사용 가이드
│
├── 📥 입력 파일 (사용자 제공)
│   └── Ambiente 2025 Exhibitor 리스트.xlsx
│
└── 📤 출력 파일 (자동 생성)
    ├── conxemar_companies.json          # 원본 JSON 데이터
    ├── conxemar_companies.xlsx          # 한글 헤더 엑셀
    └── conxemar_ambiente_format.xlsx    # Ambiente 양식 엑셀
```

## 📊 데이터 구조 매핑

| Ambiente 2025 양식 | Conxemar 수집 데이터 | 설명 |
|-------------------|-------------------|------|
| Company Name | 회사명 | 기업명 |
| Stand Number | 부스번호 | 전시 부스 번호 |
| Country | 국가 | 국가명 |
| City | 도시 | 도시명 |
| Address | 주소 | 회사 주소 |
| Postal Code | 우편번호 | 우편번호 |
| Phone | 전화번호 | 연락처 |
| Email | 이메일 | 이메일 주소 |
| Website | 웹사이트 | 회사 홈페이지 |
| Contact Person | 이메일 추출 | 연락담당자 (이메일에서 추출) |
| Industry Sector | 업종 | 산업 분야 |
| Product Categories | 하위업종 | 제품 카테고리 |
| Social Media | 통합 처리 | SNS 링크 통합 |

## 📈 실행 결과 예시

### 데이터 수집 결과
```
🚀 Conxemar 박람회 데이터 수집을 시작합니다...

페이지 1 요청 중...
페이지 1: 20개 기업 데이터 수집
...
총 677개의 기업 데이터 수집 완료

=== 데이터 수집 통계 ===
총 기업 수: 677

국가별 기업 수:
  Spain: 580개
  Portugal: 45개
  France: 23개
  Italy: 15개
  ...

주요 업종별 기업 수:
  FISH: 234개
  SERVICES: 156개
  MACHINERY: 98개
  ...
```

### 양식 변환 결과
```
🚀 Conxemar → Ambiente 2025 양식 변환을 시작합니다...

✅ 기준 양식 분석 완료
시트명: Exhibitor List
헤더 수: 19

Conxemar 데이터 로드 완료: 677개 기업

=== 변환 통계 ===
총 변환 기업 수: 677

데이터 품질 지표:
  이메일 보유: 645개 (95.3%)
  웹사이트 보유: 589개 (87.0%)
  전화번호 보유: 652개 (96.3%)
  부스번호 보유: 677개 (100.0%)

✅ 변환 완료!
📁 생성된 파일: conxemar_ambiente_format.xlsx
```

## 🛠️ 고급 사용법

### 1. 기준 양식 구조 확인
```bash
# 실제 Ambiente 양식의 정확한 헤더 구조 분석
npm run analyze
```

### 2. 특정 필터링
```javascript
// ambiente_formatter.js 수정 예시
// 특정 국가만 필터링
const filteredData = this.sourceData.filter(company => 
    ['Spain', 'Portugal', 'France'].includes(company['국가'])
);

// 특정 업종만 필터링  
const filteredData = this.sourceData.filter(company => 
    company['업종'].includes('FISH')
);
```

### 3. 커스텀 필드 매핑
```javascript
// ambiente_formatter.js의 createFieldMapping() 메서드 수정
const customMapping = {
    'Custom Field 1': '회사명',
    'Custom Field 2': '부스번호',
    // ... 추가 매핑
};
```

## ⚡ 성능 최적화

- **병렬 처리**: 여러 페이지 동시 요청 (서버 부하 고려)
- **캐싱**: 수집된 데이터 로컬 캐싱
- **메모리 효율**: 대용량 데이터 스트리밍 처리
- **오류 복구**: 네트워크 오류 시 자동 재시도

## ⚠️ 주의사항

1. **기준 양식 파일**: `Ambiente 2025 Exhibitor 리스트.xlsx` 파일이 프로젝트 폴더에 있어야 함
2. **네트워크 연결**: 안정적인 인터넷 연결 필요
3. **서버 정책**: Conxemar 서버의 접근 정책 변경 가능성
4. **데이터 실시간성**: 박람회 데이터는 실시간으로 변경될 수 있음

## 🔧 트러블슈팅

### 파일 관련 오류
```bash
# 파일 권한 확인
ls -la *.xlsx

# 파일명 정확성 확인
file "Ambiente 2025 Exhibitor 리스트.xlsx"
```

### 네트워크 오류
```bash
# Conxemar 웹사이트 접근 확인
curl -I https://conxemar.net

# 프록시 설정 (필요시)
export HTTP_PROXY=http://proxy-server:port
export HTTPS_PROXY=http://proxy-server:port
```

### 메모리 부족
```bash
# Node.js 메모리 증가
node --max-old-space-size=4096 conxemar_scraper.js
```

## 🤝 기여 방법

1. **이슈 리포트**: [GitHub Issues](https://github.com/DevJihwan/conxemar-data-collector/issues)
2. **기능 요청**: Issue를 통한 기능 제안
3. **Pull Request**: 코드 개선 사항 제출

## 📄 라이선스

MIT License - 자유롭게 사용, 수정, 배포 가능합니다.

---

**🧑‍💻 개발자**: DevJihwan  
**📅 최종 업데이트**: 2025-09-01  
**🔗 리포지토리**: https://github.com/DevJihwan/conxemar-data-collector  
**⭐ 도움이 되셨다면 Star를 눌러주세요!**