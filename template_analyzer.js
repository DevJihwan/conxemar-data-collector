const XLSX = require('xlsx');
const fs = require('fs').promises;

/**
 * Ambiente 2025 기준 양식 파일을 분석하여 정확한 헤더 구조를 파악하는 도구
 */
class TemplateAnalyzer {
    constructor() {
        this.templatePath = 'Ambiente 2025 Exhibitor 리스트.xlsx';
        this.templateData = null;
        this.headers = null;
    }

    // 기준 양식 파일 분석
    async analyzeTemplate() {
        try {
            console.log('📊 Ambiente 2025 기준 양식 분석을 시작합니다...\n');
            
            const templateBuffer = await fs.readFile(this.templatePath);
            const workbook = XLSX.read(templateBuffer, {
                cellStyles: true,
                cellFormulas: true,
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });

            console.log('📁 워크북 정보:');
            console.log(`  파일명: ${this.templatePath}`);
            console.log(`  시트 수: ${workbook.SheetNames.length}`);
            console.log(`  시트 목록: ${workbook.SheetNames.join(', ')}`);

            // 각 시트 분석
            for (let i = 0; i < workbook.SheetNames.length; i++) {
                const sheetName = workbook.SheetNames[i];
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                console.log(`\n📋 시트 ${i + 1}: "${sheetName}"`);
                console.log(`  총 행 수: ${sheetData.length}`);
                
                if (sheetData.length > 0) {
                    console.log('  헤더 (첫 번째 행):');
                    sheetData[0].forEach((header, index) => {
                        console.log(`    ${index + 1}. "${header}"`);
                    });
                    
                    // 메인 시트 데이터 저장
                    if (i === 0) {
                        this.templateData = sheetData;
                        this.headers = sheetData[0];
                    }
                    
                    // 샘플 데이터 출력
                    if (sheetData.length > 1) {
                        console.log('  첫 번째 데이터 행 예시:');
                        sheetData[1].forEach((value, index) => {
                            if (sheetData[0][index] && value) {
                                console.log(`    ${sheetData[0][index]}: "${value}"`);
                            }
                        });
                    }
                }
            }

            return {
                sheetNames: workbook.SheetNames,
                mainSheet: workbook.SheetNames[0],
                headers: this.headers,
                sampleData: this.templateData ? this.templateData.slice(1, 3) : [],
                totalRows: this.templateData ? this.templateData.length : 0
            };

        } catch (error) {
            console.error('❌ 기준 양식 파일 분석 실패:', error.message);
            console.error('💡 다음을 확인해주세요:');
            console.error('   1. 파일이 프로젝트 폴더에 있는지 확인');
            console.error('   2. 파일명이 정확한지 확인 (Ambiente 2025 Exhibitor 리스트.xlsx)');
            console.error('   3. 파일이 손상되지 않았는지 확인');
            throw error;
        }
    }

    // 필드 매핑 제안 생성
    generateMappingSuggestions() {
        if (!this.headers) {
            console.log('헤더 정보가 없습니다. 먼저 analyzeTemplate()을 실행해주세요.');
            return;
        }

        console.log('\n🔧 필드 매핑 제안:');
        console.log('// ambiente_formatter.js에서 사용할 수 있는 매핑 코드\n');
        
        console.log('const fieldMapping = {');
        this.headers.forEach(header => {
            if (!header) return;
            
            const lowerHeader = header.toLowerCase();
            let suggestion = '';
            
            // 매핑 제안 로직
            if (lowerHeader.includes('company') && lowerHeader.includes('name')) {
                suggestion = '회사명';
            } else if (lowerHeader.includes('stand') || lowerHeader.includes('booth')) {
                suggestion = '부스번호';
            } else if (lowerHeader.includes('country')) {
                suggestion = '국가';
            } else if (lowerHeader.includes('city')) {
                suggestion = '도시';
            } else if (lowerHeader.includes('address')) {
                suggestion = '주소';
            } else if (lowerHeader.includes('postal') || lowerHeader.includes('zip')) {
                suggestion = '우편번호';
            } else if (lowerHeader.includes('phone') || lowerHeader.includes('tel')) {
                suggestion = '전화번호';
            } else if (lowerHeader.includes('fax')) {
                suggestion = '팩스';
            } else if (lowerHeader.includes('email') || lowerHeader.includes('mail')) {
                suggestion = '이메일';
            } else if (lowerHeader.includes('website') || lowerHeader.includes('web')) {
                suggestion = '웹사이트';
            } else if (lowerHeader.includes('contact')) {
                suggestion = 'CONTACT_PERSON';
            } else if (lowerHeader.includes('industry') || lowerHeader.includes('sector')) {
                suggestion = '업종';
            } else if (lowerHeader.includes('product')) {
                suggestion = '하위업종';
            } else if (lowerHeader.includes('description')) {
                suggestion = '설명';
            } else if (lowerHeader.includes('hall') || lowerHeader.includes('pavilion')) {
                suggestion = '전시관';
            } else if (lowerHeader.includes('region') || lowerHeader.includes('state')) {
                suggestion = '지역';
            } else if (lowerHeader.includes('social')) {
                suggestion = 'SOCIAL_MEDIA';
            } else {
                suggestion = '// 매핑 필요';
            }
            
            console.log(`    "${header}": "${suggestion}",`);
        });
        console.log('};');
    }

    // JavaScript 배열 형태로 헤더 출력
    generateHeadersArray() {
        if (!this.headers) {
            console.log('헤더 정보가 없습니다.');
            return;
        }

        console.log('\n📝 헤더 배열 (코드에서 사용):');
        console.log('const ambienteHeaders = [');
        this.headers.forEach(header => {
            console.log(`    "${header}",`);
        });
        console.log('];');
    }

    // 메인 실행 함수
    async run() {
        try {
            const result = await this.analyzeTemplate();
            this.generateMappingSuggestions();
            this.generateHeadersArray();
            
            console.log('\n✅ 분석 완료!');
            console.log('💡 위의 정보를 ambiente_formatter.js에 반영하여 정확한 양식 변환을 수행하세요.');
            
            return result;
        } catch (error) {
            console.error('분석 실패:', error.message);
        }
    }
}

// 실행
async function analyzeAmbienteTemplate() {
    const analyzer = new TemplateAnalyzer();
    return await analyzer.run();
}

// 직접 실행
if (require.main === module) {
    analyzeAmbienteTemplate();
}

module.exports = TemplateAnalyzer;