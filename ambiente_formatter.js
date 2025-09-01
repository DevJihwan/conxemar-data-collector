const XLSX = require('xlsx');
const fs = require('fs').promises;

class ConxemarToAmbienteFormatter {
    constructor() {
        // Ambiente 2025 기준 양식 (실제 파일 분석 후 업데이트 예정)
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
        this.actualTemplateHeaders = null;
    }

    // 실제 기준 양식 파일 분석
    async analyzeTemplateFile(templatePath = 'Ambiente 2025 Exhibitor 리스트.xlsx') {
        try {
            console.log('기준 양식 파일 분석 중...');
            const templateBuffer = await fs.readFile(templatePath);
            const templateWorkbook = XLSX.read(templateBuffer, {
                cellStyles: true,
                cellFormulas: true,
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });

            const sheetName = templateWorkbook.SheetNames[0];
            const worksheet = templateWorkbook.Sheets[sheetName];
            const templateData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            // 실제 헤더로 업데이트
            this.actualTemplateHeaders = templateData[0] || [];
            this.ambienteHeaders = this.actualTemplateHeaders;
            
            console.log(`✅ 기준 양식 분석 완료`);
            console.log(`시트명: ${sheetName}`);
            console.log(`헤더 수: ${this.actualTemplateHeaders.length}`);
            console.log('실제 헤더:', this.actualTemplateHeaders);
            
            return {
                sheetName,
                headers: this.actualTemplateHeaders,
                sampleData: templateData.slice(1, 3),
                totalRows: templateData.length
            };
        } catch (error) {
            console.warn('⚠️  기준 양식 파일을 찾을 수 없습니다. 기본 양식을 사용합니다.');
            console.warn('오류:', error.message);
            return null;
        }
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
            
            // 수집된 데이터의 구조 출력
            if (this.sourceData.length > 0) {
                console.log('수집된 데이터 필드:', Object.keys(this.sourceData[0]));
            }
            
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

    // 스마트 필드 매핑 (기준 양식 헤더에 따라 동적 매핑)
    createFieldMapping() {
        const mapping = {};
        
        this.ambienteHeaders.forEach(header => {
            if (!header) return;
            
            const lowerHeader = header.toLowerCase();
            
            // 회사명 관련
            if (lowerHeader.includes('company') && lowerHeader.includes('name')) {
                mapping[header] = '회사명';
            }
            // 부스번호 관련
            else if (lowerHeader.includes('stand') || lowerHeader.includes('booth')) {
                mapping[header] = '부스번호';
            }
            // 국가 관련
            else if (lowerHeader.includes('country')) {
                mapping[header] = '국가';
            }
            // 도시 관련
            else if (lowerHeader.includes('city')) {
                mapping[header] = '도시';
            }
            // 주소 관련
            else if (lowerHeader.includes('address')) {
                mapping[header] = '주소';
            }
            // 우편번호 관련
            else if (lowerHeader.includes('postal') || lowerHeader.includes('zip')) {
                mapping[header] = '우편번호';
            }
            // 전화번호 관련
            else if (lowerHeader.includes('phone') || lowerHeader.includes('tel')) {
                mapping[header] = '전화번호';
            }
            // 팩스 관련
            else if (lowerHeader.includes('fax')) {
                mapping[header] = '팩스';
            }
            // 이메일 관련
            else if (lowerHeader.includes('email') || lowerHeader.includes('mail')) {
                mapping[header] = '이메일';
            }
            // 웹사이트 관련
            else if (lowerHeader.includes('website') || lowerHeader.includes('web')) {
                mapping[header] = '웹사이트';
            }
            // 담당자 관련
            else if (lowerHeader.includes('contact') && lowerHeader.includes('person')) {
                mapping[header] = 'CONTACT_PERSON'; // 특별 처리
            }
            // 업종 관련
            else if (lowerHeader.includes('industry') || lowerHeader.includes('sector')) {
                mapping[header] = '업종';
            }
            // 제품 카테고리 관련
            else if (lowerHeader.includes('product') && lowerHeader.includes('categories')) {
                mapping[header] = '하위업종';
            }
            // 설명 관련
            else if (lowerHeader.includes('description')) {
                mapping[header] = '설명';
            }
            // 전시관 관련
            else if (lowerHeader.includes('hall') || lowerHeader.includes('pavilion')) {
                mapping[header] = '전시관';
            }
            // 지역/주 관련
            else if (lowerHeader.includes('region') || lowerHeader.includes('state')) {
                mapping[header] = '지역';
            }
            // 소셜미디어 관련
            else if (lowerHeader.includes('social') && lowerHeader.includes('media')) {
                mapping[header] = 'SOCIAL_MEDIA'; // 특별 처리
            }
            // 제품수 관련
            else if (lowerHeader.includes('number') && lowerHeader.includes('products')) {
                mapping[header] = '제품수';
            }
            // 회사ID 관련
            else if (lowerHeader.includes('company') && lowerHeader.includes('id')) {
                mapping[header] = '회사ID';
            }
        });
        
        console.log('\n필드 매핑 결과:');
        Object.entries(mapping).forEach(([ambienteField, conxemarField]) => {
            console.log(`  "${ambienteField}" ← "${conxemarField}"`);
        });
        
        return mapping;
    }

    // Conxemar 데이터를 Ambiente 양식으로 변환
    convertToAmbienteFormat() {
        if (!this.sourceData || this.sourceData.length === 0) {
            throw new Error('변환할 Conxemar 데이터가 없습니다.');
        }

        const fieldMapping = this.createFieldMapping();
        const convertedData = this.sourceData.map((company, index) => {
            const mappedCompany = {};
            
            this.ambienteHeaders.forEach(header => {
                if (!header) {
                    mappedCompany[header] = '';
                    return;
                }
                
                let value = '';
                const sourceField = fieldMapping[header];
                
                if (sourceField === 'CONTACT_PERSON') {
                    // 연락담당자는 이메일에서 추출
                    value = this.extractContactPerson(company['이메일']);
                } else if (sourceField === 'SOCIAL_MEDIA') {
                    // 소셜미디어는 통합
                    value = this.combineSocialMedia(company);
                } else if (sourceField && company[sourceField] !== undefined) {
                    // 직접 매핑
                    value = company[sourceField];
                    
                    // 웹사이트 정규화
                    if (sourceField === '웹사이트') {
                        value = this.normalizeWebsite(value);
                    }
                }
                
                // 값 후처리
                if (typeof value === 'boolean') {
                    value = value ? 'Y' : 'N';
                } else if (value === null || value === undefined) {
                    value = '';
                } else {
                    value = value.toString().trim();
                }
                
                mappedCompany[header] = value;
            });
            
            return mappedCompany;
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
                if (!header) return { wch: 10 };
                
                const maxLength = Math.max(
                    header.length,
                    ...convertedData.map(row => (row[header] || '').toString().length)
                );
                return { wch: Math.min(maxLength + 2, 50) };
            });
            worksheet['!cols'] = columnWidths;

            // 시트 추가 (실제 양식의 시트명 사용)
            const sheetName = this.actualTemplateHeaders ? 'Exhibitor List' : 'Exhibitor List';
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

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
        console.log(`총 변환 기업 수: ${this.sourceData.length}`);
        
        // 국가별 분포
        const countryStats = {};
        this.sourceData.forEach(company => {
            const country = company['국가'] || '미상';
            countryStats[country] = (countryStats[country] || 0) + 1;
        });
        
        console.log('\n국가별 기업 분포 (상위 10개):');
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

        console.log('\n업종별 기업 분포 (상위 10개):');
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

        console.log('\n데이터 품질 지표:');
        console.log(`  이메일 보유: ${qualityStats.withEmail}개 (${(qualityStats.withEmail/this.sourceData.length*100).toFixed(1)}%)`);
        console.log(`  웹사이트 보유: ${qualityStats.withWebsite}개 (${(qualityStats.withWebsite/this.sourceData.length*100).toFixed(1)}%)`);
        console.log(`  전화번호 보유: ${qualityStats.withPhone}개 (${(qualityStats.withPhone/this.sourceData.length*100).toFixed(1)}%)`);
        console.log(`  부스번호 보유: ${qualityStats.withStand}개 (${(qualityStats.withStand/this.sourceData.length*100).toFixed(1)}%)`);
    }

    // 필드 매핑 상태 확인
    checkFieldMapping() {
        const fieldMapping = this.createFieldMapping();
        
        console.log('\n=== 필드 매핑 확인 ===');
        console.log(`기준 양식 헤더 수: ${this.ambienteHeaders.length}`);
        console.log(`매핑된 필드 수: ${Object.keys(fieldMapping).length}`);
        
        // 매핑되지 않은 헤더 확인
        const unmappedHeaders = this.ambienteHeaders.filter(header => 
            header && !fieldMapping[header]
        );
        
        if (unmappedHeaders.length > 0) {
            console.log('\n⚠️  매핑되지 않은 헤더:');
            unmappedHeaders.forEach(header => {
                console.log(`  "${header}"`);
            });
        }
        
        return fieldMapping;
    }

    // 메인 실행 함수
    async run() {
        try {
            console.log('🚀 Conxemar → Ambiente 2025 양식 변환을 시작합니다...\n');
            
            // 1. 기준 양식 분석 (선택사항)
            await this.analyzeTemplateFile();
            
            // 2. 데이터 로드 (엑셀 우선, 실패시 JSON)
            try {
                await this.loadConxemarData();
            } catch {
                console.log('엑셀 파일 로드 실패, JSON 파일 시도...');
                await this.loadConxemarDataFromJSON();
            }
            
            // 3. 필드 매핑 확인
            this.checkFieldMapping();
            
            // 4. 변환 통계 출력
            this.printConversionStats();
            
            // 5. Ambiente 양식으로 변환 및 저장
            await this.generateAmbienteExcel();
            
            console.log('\n✅ 변환 완료!');
            console.log('📁 생성된 파일: conxemar_ambiente_format.xlsx');
            console.log('\n💡 변환된 파일을 확인하시고, 필요시 필드 매핑을 조정해주세요.');
            
        } catch (error) {
            console.error('❌ 변환 과정에서 오류 발생:', error.message);
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