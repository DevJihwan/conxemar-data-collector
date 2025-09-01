const ConxemarDataCollector = require('./conxemar_scraper');
const TemplateAnalyzer = require('./template_analyzer');
const ConxemarToAmbienteFormatter = require('./ambiente_formatter');
const fs = require('fs').promises;

/**
 * 전체 워크플로우를 자동으로 실행하는 메인 스크립트
 */
class MainWorkflow {
    constructor() {
        this.collector = new ConxemarDataCollector();
        this.analyzer = new TemplateAnalyzer();
        this.formatter = new ConxemarToAmbienteFormatter();
    }

    // 필수 파일 존재 여부 확인
    async checkRequiredFiles() {
        console.log('📋 필수 파일 확인 중...');
        
        const templateFile = 'Ambiente 2025 Exhibitor 리스트.xlsx';
        
        try {
            await fs.access(templateFile);
            console.log('✅ 기준 양식 파일 확인됨:', templateFile);
            return true;
        } catch (error) {
            console.log('❌ 기준 양식 파일을 찾을 수 없습니다:', templateFile);
            console.log('💡 다음 중 하나를 수행해주세요:');
            console.log('   1. Ambiente 2025 기준 양식 파일을 프로젝트 폴더에 복사');
            console.log('   2. --skip-template 옵션으로 기본 양식 사용');
            return false;
        }
    }

    // 전체 워크플로우 실행
    async runFullWorkflow(options = {}) {
        try {
            console.log('🎬 Conxemar → Ambiente 2025 전체 워크플로우 시작\n');
            console.log('=' .repeat(60));
            
            // 1. 필수 파일 확인
            if (!options.skipTemplate) {
                const hasTemplate = await this.checkRequiredFiles();
                if (!hasTemplate && !options.force) {
                    console.log('\n❌ 워크플로우를 중단합니다.');
                    console.log('💡 --force 옵션으로 강제 실행하거나 기준 양식 파일을 추가해주세요.');
                    return;
                }
            }

            // 2. 데이터 수집
            console.log('\n🔍 STEP 1: 데이터 수집');
            console.log('-' .repeat(40));
            await this.collector.run();

            // 3. 기준 양식 분석 (선택적)
            if (!options.skipTemplate) {
                console.log('\n📊 STEP 2: 기준 양식 분석');
                console.log('-' .repeat(40));
                try {
                    await this.analyzer.run();
                } catch (error) {
                    console.log('⚠️  기준 양식 분석 실패, 기본 양식 사용');
                }
            }

            // 4. 양식 변환
            console.log('\n🔄 STEP 3: 양식 변환');
            console.log('-' .repeat(40));
            await this.formatter.run();

            // 5. 완료 요약
            console.log('\n' + '=' .repeat(60));
            console.log('🎉 전체 워크플로우 완료!');
            console.log('\n📁 생성된 파일:');
            console.log('   • conxemar_companies.json      (원본 JSON 데이터)');
            console.log('   • conxemar_companies.xlsx      (한글 헤더 엑셀)');
            console.log('   • conxemar_ambiente_format.xlsx (Ambiente 양식 엑셀)');
            
            console.log('\n💡 다음 단계:');
            console.log('   1. conxemar_ambiente_format.xlsx 파일을 확인하세요');
            console.log('   2. 필요시 필드 매핑을 조정하여 재실행하세요');
            console.log('   3. 최종 파일을 Ambiente 2025 시스템에 업로드하세요');

        } catch (error) {
            console.error('\n❌ 워크플로우 실행 중 오류 발생:', error.message);
            console.error('\n🔧 문제 해결 방법:');
            console.error('   1. 네트워크 연결 상태 확인');
            console.error('   2. 필수 파일 존재 여부 확인');
            console.error('   3. Node.js 버전 확인 (권장: v14 이상)');
        }
    }

    // 개별 단계 실행
    async runDataCollection() {
        console.log('🔍 데이터 수집만 실행합니다...\n');
        await this.collector.run();
    }

    async runTemplateAnalysis() {
        console.log('📊 기준 양식 분석만 실행합니다...\n');
        await this.analyzer.run();
    }

    async runFormatConversion() {
        console.log('🔄 양식 변환만 실행합니다...\n');
        await this.formatter.run();
    }
}

// 명령행 인수 처리
function parseArguments() {
    const args = process.argv.slice(2);
    const options = {
        skipTemplate: args.includes('--skip-template'),
        force: args.includes('--force'),
        collectOnly: args.includes('--collect-only'),
        analyzeOnly: args.includes('--analyze-only'),
        formatOnly: args.includes('--format-only')
    };
    
    return options;
}

// 도움말 출력
function printHelp() {
    console.log(`
🏢 Conxemar Data Collector & Ambiente Formatter

사용법:
  node main.js [옵션]

옵션:
  --collect-only     데이터 수집만 실행
  --analyze-only     양식 분석만 실행  
  --format-only      양식 변환만 실행
  --skip-template    기준 양식 분석 생략
  --force           오류 무시하고 강제 실행
  --help            이 도움말 출력

예시:
  node main.js                    # 전체 워크플로우 실행
  node main.js --collect-only     # 데이터 수집만 실행
  node main.js --skip-template    # 기본 양식으로 변환
  node main.js --force            # 오류 무시하고 강제 실행
`);
}

// 메인 실행
async function main() {
    const options = parseArguments();
    
    if (process.argv.includes('--help') || process.argv.includes('-h')) {
        printHelp();
        return;
    }

    const workflow = new MainWorkflow();

    try {
        if (options.collectOnly) {
            await workflow.runDataCollection();
        } else if (options.analyzeOnly) {
            await workflow.runTemplateAnalysis();
        } else if (options.formatOnly) {
            await workflow.runFormatConversion();
        } else {
            await workflow.runFullWorkflow(options);
        }
    } catch (error) {
        console.error('메인 워크플로우 실행 실패:', error.message);
        process.exit(1);
    }
}

// 스크립트 직접 실행 시
if (require.main === module) {
    main();
}

module.exports = MainWorkflow;