const XLSX = require('xlsx');
const fs = require('fs').promises;

class ConxemarToAmbienteFormatter {
    constructor() {
        // Ambiente 2025 기준 양식 (일반적인 국제 박람회 양식)
        this.ambienteHeaders = [
            'Company Name',
            'Stand Number', 
            'Country',
            'City',
            'Address',
            'Postal Code',
            'Phone',
            'Fax',
            'Email',
            'Website',
            'Contact Person',
            'Industry Sector',
            'Product Categories',
            'Company Description',
            'Hall/Pavilion',
            'Region/State',
            'Social Media',
            'Number of Products',
            'Company ID'
        ];
        
        this.sourceData = null;
    }

    // Conxemar 수집 데이터 로드
    async loadConxemarData(filePath = 'conxemar_companies.xlsx') {
        try {
            const fileBuffer = await fs.readFile(filePath);
            const workbook = XLSX.read(fileBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            this.sourceData = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Conxemar 데이터 로드 완료: ${this.sourceData.length}개 기업`);
            return this.sourceData;
        } catch (error) {
            console.error('Conxemar 데이터 로드 실패:', error.message);
            throw error;
        }
    }

    // JSON에서 로드 (대안)
    async loadConxemarDataFromJSON(filePath = 'conxemar_companies.json') {
        try {
            const jsonData = await fs.readFile(filePath, 'utf8');
            this.sourceData = JSON.parse(jsonData);
            
            console.log(`Conxemar JSON 데이터 로드 완료: ${this.sourceData.length}개 기업`);
            return this.sourceData;
        } catch (error) {
            console.error('Conxemar JSON 데이터 로드 실패:', error.message);
            throw error;
        }
    }

    // 연락처 정보 추출 (이메일에서 연락담당자 추정)
    extractContactPerson(email) {
        if (!email) return '';
        
        // 첫 번째 이메일에서 @ 앞부분을 연락담당자로 추정
        const firstEmail = email.split(';')[0].trim();
        const localPart = firstEmail.split('@')[0];
        
        // 일반적인 패턴에서 이름 추출
        if (localPart.includes('.')) {
            const parts = localPart.split('.');
            return parts.map(part => 
                part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
            ).join(' ');
        }
        
        return localPart.charAt(0).toUpperCase() + localPart.slice(1).toLowerCase();
    }

    // 소셜미디어 정보 통합
    combineSocialMedia(company) {
        const socialFields = ['페이스북', '트위터', '링크드인', '인스타그램', '유튜브'];
        const socialLinks = [];
        
        socialFields.forEach(field => {
            if (company[field] && company[field].trim()) {
                socialLinks.push(company[field].trim());
            }
        });
        
        return socialLinks.join('; ');
    }

    // 웹사이트 URL 정규화
    normalizeWebsite(website) {
        if (!website || !website.trim()) return '';
        
        const url = website.trim();
        if (!url.startsWith('http://') && !url.startsWith('https://')) {
            return `https://${url}`;
        }
        return url;
    }

    // Conxemar 데이터를 Ambiente 양식으로 변환
    convertToAmbienteFormat() {
        if (!this.sourceData || this.sourceData.length === 0) {
            throw new Error('변환할 Conxemar 데이터가 없습니다.');
        }

        const convertedData = this.sourceData.map((company, index) => {
            return {
                'Company Name': company['회사명'] || '',
                'Stand Number': company['부스번호'] || '',
                'Country': company['국가'] || '',
                'City': company['도시'] || '',
                'Address': company['주소'] || '',
                'Postal Code': company['우편번호'] || '',
                'Phone': company['전화번호'] || '',
                'Fax': company['팩스'] || '',
                'Email': company['이메일'] || '',
                'Website': this.normalizeWebsite(company['웹사이트']),
                'Contact Person': this.extractContactPerson(company['이메일']),
                'Industry Sector': company['업종'] || '',
                'Product Categories': company['하위업종'] || company['업종'] || '',
                'Company Description': company['설명'] || '',
                'Hall/Pavilion': company['전시관'] || '',
                'Region/State': company['지역'] || company['구역'] || '',
                'Social Media': this.combineSocialMedia(company),
                'Number of Products': company['제품수'] || 0,
                'Company ID': company['회사ID'] || ''
            };
        });

        console.log(`${convertedData.length}개 기업 데이터를 Ambiente 양식으로 변환 완료`);
        return convertedData;
    }

    // Ambiente 양식에 맞춘 엑셀 파일 생성
    async generateAmbienteExcel(outputPath = 'conxemar_ambiente_format.xlsx') {
        try {
            const convertedData = this.convertToAmbienteFormat();
            
            // 새 워크북 생성
            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.json_to_sheet(convertedData);

            // 컬럼 너비 최적화
            const columnWidths = this.ambienteHeaders.map(header => {
                const maxLength = Math.max(
                    header.length,
                    ...convertedData.map(row => (row[header] || '').toString().length)
                );
                return { wch: Math.min(maxLength + 2, 50) };
            });
            worksheet['!cols'] = columnWidths;

            // 시트 추가
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Exhibitor List');

            // 파일 저장
            XLSX.writeFile(workbook, outputPath);
            console.log(`Ambiente 양식 엑셀 파일 저장 완료: ${outputPath}`);
            
            return outputPath;
        } catch (error) {
            console.error('Ambiente 양식 변환 중 오류:', error.message);
            throw error;
        }
    }

    // 변환 통계 출력
    printConversionStats() {
        if (!this.sourceData) {
            console.log('변환할 데이터가 없습니다.');
            return;
        }

        console.log('\n=== 변환 통계 ===');
        
        // 국가별 분포
        const countryStats = {};
        this.sourceData.forEach(company => {
            const country = company['국가'] || '미상';
            countryStats[country] = (countryStats[country] || 0) + 1;
        });
        
        console.log('\n국가별 기업 분포:');
        Object.entries(countryStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10)
            .forEach(([country, count]) => {
                console.log(`  ${country}: ${count}개`);
            });

        // 업종별 분포
        const sectorStats = {};
        this.sourceData.forEach(company => {
            const sector = company['업종'] || '미상';
            if (sector !== '미상') {
                sector.split(',').forEach(s => {
                    const trimmed = s.trim();
                    if (trimmed) {
                        sectorStats[trimmed] = (sectorStats[trimmed] || 0) + 1;
                    }
                });
            } else {
                sectorStats[sector] = (sectorStats[sector] || 0) + 1;
            }
        });

        console.log('\n업종별 기업 분포:');
        Object.entries(sectorStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10)
            .forEach(([sector, count]) => {
                console.log(`  ${sector}: ${count}개`);
            });

        // 데이터 품질 확인
        const qualityStats = {
            withEmail: this.sourceData.filter(c => c['이메일'] && c['이메일'].trim()).length,
            withWebsite: this.sourceData.filter(c => c['웹사이트'] && c['웹사이트'].trim()).length,
            withPhone: this.sourceData.filter(c => c['전화번호'] && c['전화번호'].trim()).length,
            withStand: this.sourceData.filter(c => c['부스번호'] && c['부스번호'].trim()).length
        };

        console.log('\n데이터 품질:');
        console.log(`  이메일 보유: ${qualityStats.withEmail}개 (${(qualityStats.withEmail/this.sourceData.length*100).toFixed(1)}%)`);
        console.log(`  웹사이트 보유: ${qualityStats.withWebsite}개 (${(qualityStats.withWebsite/this.sourceData.length*100).toFixed(1)}%)`);
        console.log(`  전화번호 보유: ${qualityStats.withPhone}개 (${(qualityStats.withPhone/this.sourceData.length*100).toFixed(1)}%)`);
        console.log(`  부스번호 보유: ${qualityStats.withStand}개 (${(qualityStats.withStand/this.sourceData.length*100).toFixed(1)}%)`);
    }

    // 메인 실행 함수
    async run() {
        try {
            console.log('Conxemar → Ambiente 양식 변환을 시작합니다...\n');
            
            // 데이터 로드 (엑셀 우선, 실패시 JSON)
            try {
                await this.loadConxemarData();
            } catch {
                console.log('엑셀 파일 로드 실패, JSON 파일 시도...');
                await this.loadConxemarDataFromJSON();
            }
            
            // 변환 통계 출력
            this.printConversionStats();
            
            // Ambiente 양식으로 변환 및 저장
            await this.generateAmbienteExcel();
            
            console.log('\n✅ 변환 완료!');
            console.log('생성된 파일: conxemar_ambiente_format.xlsx');
            
        } catch (error) {
            console.error('변환 과정에서 오류 발생:', error.message);
        }
    }
}

// 사용법
async function convertToAmbiente() {
    const formatter = new ConxemarToAmbienteFormatter();
    await formatter.run();
}

// 직접 실행
if (require.main === module) {
    convertToAmbiente();
}

module.exports = ConxemarToAmbienteFormatter;