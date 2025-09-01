const axios = require('axios');
const XLSX = require('xlsx');
const fs = require('fs').promises;

class ConxemarDataCollector {
    constructor() {
        this.baseUrl = 'https://conxemar.net/Conxemar2022/en/Company/Companies_Read';
        this.queryParams = {
            Sector1: 0,
            Sector2: 0,
            Sector3: 0,
            Index: 'System.Collections.Generic.List`1[System.String]',
            IsNew: false,
            IsSpecial: false,
            HasProducts: false,
            hasActivities: false,
            Tags: 'System.Collections.Generic.List`1[System.Web.Mvc.SelectListItem]',
            AvailableSubindustries: 'System.Collections.Generic.List`1[System.Web.Mvc.SelectListItem]',
            SelectedSubindustryID: 0,
            ProductTag1: false,
            ProductTag2: false,
            FilterByCountry: true,
            FilterByState: false,
            AnyFilterAddress: true,
            Feat_Avoid_Sector_Propagation: false,
            IdCompany: 0,
            idEvent: 0,
            iFrameMode: false,
            Features: 'System.Collections.Generic.Dictionary`2[System.String,System.Object]',
            HelpTexts: 'System.Collections.Generic.Dictionary`2[System.String,IventPublicPortal.Models.HelpTextModelView]'
        };
        
        this.headers = {
            'Accept': '*/*',
            'Accept-Encoding': 'gzip, deflate, br, zstd',
            'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Origin': 'https://conxemar.net',
            'Referer': 'https://conxemar.net/Conxemar2022/en/company/search',
            'User-Agent': 'Mozilla/5.0 (Linux; Android 11.0; Surface Duo) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Mobile Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest'
        };
        
        this.allData = [];
    }

    // URL 쿼리 파라미터 생성
    buildQueryString() {
        const params = new URLSearchParams();
        Object.entries(this.queryParams).forEach(([key, value]) => {
            params.append(key, value.toString());
        });
        return params.toString();
    }

    // 폼 데이터 생성
    buildFormData(page = 1, pageSize = 20) {
        const formData = new URLSearchParams();
        formData.append('sort', 'corder-asc~Name-asc');
        formData.append('page', page.toString());
        formData.append('pageSize', pageSize.toString());
        formData.append('group', '');
        formData.append('filter', '');
        return formData.toString();
    }

    // 단일 페이지 데이터 요청
    async fetchPage(page = 1, pageSize = 20) {
        try {
            const url = `${this.baseUrl}?${this.buildQueryString()}`;
            const formData = this.buildFormData(page, pageSize);

            console.log(`페이지 ${page} 요청 중...`);

            const response = await axios.post(url, formData, {
                headers: this.headers,
                timeout: 30000
            });

            if (response.status === 200 && response.data && response.data.Data) {
                console.log(`페이지 ${page}: ${response.data.Data.length}개 기업 데이터 수집`);
                return response.data;
            } else {
                throw new Error(`페이지 ${page}: 유효하지 않은 응답`);
            }
        } catch (error) {
            console.error(`페이지 ${page} 요청 실패:`, error.message);
            throw error;
        }
    }

    // 모든 페이지 데이터 수집
    async fetchAllData() {
        let page = 1;
        const pageSize = 20;
        let hasMoreData = true;
        
        while (hasMoreData) {
            try {
                const response = await this.fetchPage(page, pageSize);
                
                if (response.Data && response.Data.length > 0) {
                    this.allData.push(...response.Data);
                    
                    // 페이지 크기보다 적은 데이터가 반환되면 마지막 페이지
                    if (response.Data.length < pageSize) {
                        hasMoreData = false;
                    } else {
                        page++;
                        // 서버 부하 방지를 위한 딜레이
                        await this.delay(1000);
                    }
                } else {
                    hasMoreData = false;
                }
            } catch (error) {
                console.error(`데이터 수집 중 오류 발생:`, error.message);
                break;
            }
        }

        console.log(`총 ${this.allData.length}개의 기업 데이터 수집 완료`);
        return this.allData;
    }

    // 딜레이 함수
    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // JSON 파일로 저장
    async saveAsJSON(filename = 'conxemar_companies.json') {
        try {
            const jsonData = JSON.stringify(this.allData, null, 2);
            await fs.writeFile(filename, jsonData, 'utf8');
            console.log(`JSON 파일 저장 완료: ${filename}`);
        } catch (error) {
            console.error('JSON 파일 저장 실패:', error.message);
        }
    }

    // 엑셀 파일로 저장
    async saveAsExcel(filename = 'conxemar_companies.xlsx') {
        try {
            // 데이터를 플랫 구조로 변환
            const flatData = this.allData.map(company => ({
                '순번': company.corder,
                '회사ID': company.IdAccount,
                '회사명': company.Name,
                '부제목': company.Subtitle || '',
                '주소': company.Address1,
                '도시': company.Town,
                '지역': company.County,
                '국가': company.Country,
                '구역': company.Area || '',
                '우편번호': company.Postcode,
                '전시관': company.Pavilion || '',
                '부스번호': company.Stand,
                '팩스': company.Fax || '',
                '전화번호': company.Telephone || '',
                '이메일': company.Email || '',
                '웹사이트': company.Web || '',
                '페이스북': company.Facebook || '',
                '트위터': company.Twitter || '',
                '링크드인': company.LinkedIn || '',
                '인스타그램': company.Instagram || '',
                '유튜브': company.Youtube || '',
                '설명': company.Description || '',
                '하위업종': company.SubIndustry || '',
                '업종': company.Sectors || '',
                '제품수': company.NumberOfProducts,
                '활동수': company.NumberOfActivities,
                '뉴스수': company.NumberOfNews,
                '로고보유': company.HasLogo ? 'Y' : 'N',
                '일정보유': company.HasAgenda ? 'Y' : 'N',
                '즐겨찾기': company.Favourite ? 'Y' : 'N'
            }));

            // 워크북 생성
            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.json_to_sheet(flatData);

            // 컬럼 너비 자동 조정
            const columnWidths = [];
            Object.keys(flatData[0] || {}).forEach(key => {
                const maxLength = Math.max(
                    key.length,
                    ...flatData.map(row => (row[key] || '').toString().length)
                );
                columnWidths.push({ wch: Math.min(maxLength + 2, 50) });
            });
            worksheet['!cols'] = columnWidths;

            // 시트 추가
            XLSX.utils.book_append_sheet(workbook, worksheet, '박람회_기업목록');

            // 파일 저장
            XLSX.writeFile(workbook, filename);
            console.log(`엑셀 파일 저장 완료: ${filename}`);
        } catch (error) {
            console.error('엑셀 파일 저장 실패:', error.message);
        }
    }

    // 데이터 통계 출력
    printStatistics() {
        if (this.allData.length === 0) {
            console.log('수집된 데이터가 없습니다.');
            return;
        }

        console.log('\n=== 데이터 수집 통계 ===');
        console.log(`총 기업 수: ${this.allData.length}`);
        
        // 국가별 통계
        const countryStats = {};
        this.allData.forEach(company => {
            const country = company.Country || '미상';
            countryStats[country] = (countryStats[country] || 0) + 1;
        });
        
        console.log('\n국가별 기업 수:');
        Object.entries(countryStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10)
            .forEach(([country, count]) => {
                console.log(`  ${country}: ${count}개`);
            });

        // 업종별 통계
        const sectorStats = {};
        this.allData.forEach(company => {
            const sectors = company.Sectors || '미상';
            sectors.split(',').forEach(sector => {
                const trimmedSector = sector.trim();
                if (trimmedSector) {
                    sectorStats[trimmedSector] = (sectorStats[trimmedSector] || 0) + 1;
                }
            });
        });

        console.log('\n주요 업종별 기업 수:');
        Object.entries(sectorStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10)
            .forEach(([sector, count]) => {
                console.log(`  ${sector}: ${count}개`);
            });
    }

    // 메인 실행 함수
    async run() {
        try {
            console.log('Conxemar 박람회 데이터 수집을 시작합니다...\n');
            
            // 모든 데이터 수집
            await this.fetchAllData();
            
            if (this.allData.length === 0) {
                console.log('수집된 데이터가 없습니다.');
                return;
            }

            // 통계 출력
            this.printStatistics();

            // 파일 저장
            console.log('\n파일 저장 중...');
            await this.saveAsJSON();
            await this.saveAsExcel();

            console.log('\n데이터 수집 및 저장이 완료되었습니다!');
            console.log('생성된 파일:');
            console.log('- conxemar_companies.json');
            console.log('- conxemar_companies.xlsx');

        } catch (error) {
            console.error('프로그램 실행 중 오류 발생:', error.message);
        }
    }
}

// 사용법
async function main() {
    const collector = new ConxemarDataCollector();
    await collector.run();
}

// 스크립트 직접 실행 시
if (require.main === module) {
    main();
}

module.exports = ConxemarDataCollector;