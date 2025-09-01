# Conxemar Data Collector & Ambiente Formatter

Conxemar 박람회 데이터를 수집하고 Ambiente 2025 양식으로 변환하는 Node.js 도구입니다.

## 🎯 프로젝트 개요

이 도구는 두 가지 주요 기능을 제공합니다:
1. **Conxemar 박람회 웹사이트에서 기업 데이터 자동 수집**
2. **수집된 데이터를 Ambiente 2025 양식으로 변환**

## 📋 기능

### 1. 데이터 수집기 (conxemar_scraper.js)
- ✅ Conxemar 박람회 웹사이트에서 기업 데이터 자동 수집
- ✅ 페이지네이션을 통한 전체 데이터 수집 (약 700개 기업)
- ✅ JSON 및 Excel 형태로 데이터 저장
- ✅ 수집 통계 및 데이터 품질 분석
- ✅ 서버 부하 방지를 위한 요청 간 딜레이

### 2. 양식 변환기 (ambiente_formatter.js)
- ✅ 수집된 Conxemar 데이터를 Ambiente 2025 양식으로 변환
- ✅ 실제 기준 양식 파일 자동 분석
- ✅ 스마트 필드 매핑 (영문 헤더 → 한글 데이터 매칭)
- ✅ 데이터 정규화 및 품질 개선
- ✅ 변환 통계 및 매핑 상태 확인

## 🚀 설치 및 사용법

### 1. 프로젝트 클론
```bash
git clone https://github.com/DevJihwan/conxemar-data-collector.git
cd conxemar-data-collector
```

### 2. 의존성 설치
```bash
npm install
```

### 3. 데이터 수집
```bash
npm run collect
# 또는
node conxemar_scraper.js
```

### 4. 양식 변환
```bash
# 기준 양식 파일(Ambiente 2025 Exhibitor 리스트.xlsx)을 
# 프로젝트 폴더에 복사한 후 실행
npm run format
# 또는
node ambiente_formatter.js
```

## 📁 출력 파일

### 데이터 수집 결과
- `conxemar_companies.json` - 원본 JSON 데이터
- `conxemar_companies.xlsx` - 한글 헤더 엑셀 파일

### 양식 변환 결과
- `conxemar_ambiente_format.xlsx` - Ambiente 2025 양식에 맞춘 엑셀 파일

## 📊 데이터 구조

### Conxemar 원본 필드
- 회사명, 부스번호, 국가, 도시, 주소, 우편번호
- 전화번호, 팩스, 이메일, 웹사이트
- 업종, 하위업종, 제품수, 활동수
- 소셜미디어 링크 (페이스북, 트위터, 링크드인 등)

### Ambiente 2025 양식 필드
- Company Name, Stand Number, Country, City
- Address, Postal Code, Phone, Fax, Email, Website
- Contact Person, Industry Sector, Product Categories
- Company Description, Hall/Pavilion, Region/State
- Social Media, Number of Products, Company ID

## 🔧 설정 옵션

### 데이터 수집 설정
```javascript
// conxemar_scraper.js에서 수정 가능
const pageSize = 20;  // 페이지당 데이터 수
const delay = 1000;   // 요청 간 대기 시간(ms)
```

### 양식 변환 설정
```javascript
// ambiente_formatter.js에서 헤더 수정 가능
this.ambienteHeaders = [
    'Company Name',
    'Stand Number',
    // ... 추가 필드
];
```

## 📈 실행 결과 예시

```
🚀 Conxemar → Ambiente 2025 양식 변환을 시작합니다...

✅ 기준 양식 분석 완료
시트명: Exhibitor List
헤더 수: 19

Conxemar 데이터 로드 완료: 677개 기업

=== 필드 매핑 확인 ===
기준 양식 헤더 수: 19
매핑된 필드 수: 15

=== 변환 통계 ===
총 변환 기업 수: 677

국가별 기업 분포 (상위 10개):
  Spain: 580개
  Portugal: 45개
  France: 23개
  ...

업종별 기업 분포 (상위 10개):
  FISH: 234개
  SERVICES: 156개
  MACHINERY: 98개
  ...

데이터 품질 지표:
  이메일 보유: 645개 (95.3%)
  웹사이트 보유: 589개 (87.0%)
  전화번호 보유: 652개 (96.3%)
  부스번호 보유: 677개 (100.0%)

✅ 변환 완료!
📁 생성된 파일: conxemar_ambiente_format.xlsx
```

## ⚠️ 주의사항

1. **서버 부하**: 요청 간 1초 딜레이로 서버 부하 방지
2. **네트워크**: 안정적인 인터넷 연결 필요
3. **데이터 변경**: 박람회 데이터는 실시간으로 변경될 수 있음
4. **에러 처리**: 네트워크 오류 시 자동 중단 및 에러 로그
5. **기준 양식**: Ambiente 기준 양식 파일이 프로젝트 폴더에 있어야 함

## 🛠️ 커스터마이징

### 새로운 양식 추가
1. `ambiente_formatter.js`를 복사하여 새 파일 생성
2. `ambienteHeaders` 배열을 원하는 양식으로 수정
3. `createFieldMapping()` 메서드에서 필드 매핑 규칙 수정

### 추가 데이터 소스
1. `conxemar_scraper.js`의 URL 및 파라미터 수정
2. 응답 데이터 구조에 맞게 파싱 로직 조정
3. 필요시 헤더 및 인증 정보 추가

### 데이터 필터링
```javascript
// 특정 국가만 필터링
const filteredData = sourceData.filter(company => 
    company['국가'] === 'Spain'
);

// 특정 업종만 필터링
const filteredData = sourceData.filter(company => 
    company['업종'].includes('FISH')
);
```

## 🔍 트러블슈팅

### 1. 파일을 찾을 수 없음
- 기준 양식 파일이 프로젝트 폴더에 있는지 확인
- 파일명이 정확한지 확인 (`Ambiente 2025 Exhibitor 리스트.xlsx`)

### 2. 네트워크 오류
- 인터넷 연결 상태 확인
- Conxemar 웹사이트 접근 가능 여부 확인
- 방화벽 설정 확인

### 3. 데이터 품질 문제
- 수집된 데이터의 통계 정보 확인
- 필요시 데이터 정제 로직 추가

## 📝 라이선스

MIT License - 자유롭게 사용, 수정, 배포 가능합니다.

## 🤝 기여

Issue나 Pull Request를 통해 개선사항을 제안해주세요!

---

**개발자**: DevJihwan  
**버전**: 1.0.0  
**최종 업데이트**: 2025-09-01  
**리포지토리**: https://github.com/DevJihwan/conxemar-data-collector