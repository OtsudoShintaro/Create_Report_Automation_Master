// グローバル変数
let rawData = [];
let processedData = [];
let departments = [];
let isProcessing = false;
let clientAnalysisData = null;

// タスク管理用グローバル変数
let taskRawData = [];
let taskProcessedData = [];
let taskClassificationData = {};

// インターン生管理用グローバル変数
let internMembersData = [];

// 初期化
document.addEventListener('DOMContentLoaded', function() {
    console.log('=== 統合システム初期化開始 ===');
    
    // ライブラリ読み込み確認
    checkLibraries();
    
    // タブ機能初期化
    initializeTabs();
    
    // 各タブのイベントリスナー初期化
    initializeProgressAnalysisTab();
    initializeTaskManagementTab();
    initializeInternManagementTab();
    
    console.log('=== 統合システム初期化完了 ===');
});

// ライブラリ読み込み確認
function checkLibraries() {
    const libraries = [
        { name: 'Papa Parse', check: () => typeof Papa !== 'undefined' },
        { name: 'PptxGenJS', check: () => typeof PptxGenJS !== 'undefined' },
        { name: 'XLSX', check: () => typeof XLSX !== 'undefined' }
    ];
    
    libraries.forEach(lib => {
        if (lib.check()) {
            console.log(`✅ ${lib.name} ライブラリが正常に読み込まれました`);
        } else {
            console.warn(`⚠️ ${lib.name} ライブラリの読み込みに失敗しました`);
        }
    });
}

// タブ機能初期化
function initializeTabs() {
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    
    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const targetTab = button.getAttribute('data-tab');
            
            // アクティブタブの切り替え
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));
            
            button.classList.add('active');
            document.getElementById(targetTab).classList.add('active');
            
            console.log(`タブ切り替え: ${targetTab}`);
        });
    });
}

// 部署別進捗分析タブの初期化
function initializeProgressAnalysisTab() {
    const fileInput = document.getElementById('fileInput1');
    const uploadArea = document.getElementById('uploadArea1');
    const uploadBtn = document.getElementById('uploadBtn1');
    
    if (!fileInput || !uploadArea || !uploadBtn) {
        console.error('部署別進捗分析タブの要素が見つかりません');
        return;
    }
    
    // ファイル選択イベント
    fileInput.addEventListener('change', (event) => handleProgressAnalysisFile(event));
    
    // アップロードボタンのクリックイベント
    uploadBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileInput.click();
    });
    
    // ドラッグ&ドロップイベント
    uploadArea.addEventListener('dragover', (e) => handleDragOver(e));
    uploadArea.addEventListener('dragleave', (e) => handleDragLeave(e));
    uploadArea.addEventListener('drop', (e) => handleProgressAnalysisDrop(e));
    
    // エクスポートボタンを無効化
    disableExportButtons('1');
}

// タスク管理タブの初期化
function initializeTaskManagementTab() {
    const fileInput = document.getElementById('fileInput2');
    const uploadArea = document.getElementById('uploadArea2');
    const uploadBtn = document.getElementById('uploadBtn2');
    
    if (!fileInput || !uploadArea || !uploadBtn) {
        console.error('タスク管理タブの要素が見つかりません');
        return;
    }
    
    // ファイル選択イベント
    fileInput.addEventListener('change', (event) => handleTaskManagementFile(event));
    
    // アップロードボタンのクリックイベント
    uploadBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileInput.click();
    });
    
    // ドラッグ&ドロップイベント
    uploadArea.addEventListener('dragover', (e) => handleDragOver(e));
    uploadArea.addEventListener('dragleave', (e) => handleDragLeave(e));
    uploadArea.addEventListener('drop', (e) => handleTaskManagementDrop(e));
    
    // エクスポートボタンを無効化
    disableExportButtons('2');
}

// インターン生管理タブの初期化
function initializeInternManagementTab() {
    const fileInput = document.getElementById('fileInput3');
    const uploadArea = document.getElementById('uploadArea3');
    const uploadBtn = document.getElementById('uploadBtn3');
    
    if (!fileInput || !uploadArea || !uploadBtn) {
        console.error('インターン生管理タブの要素が見つかりません');
        return;
    }
    
    // ファイル選択イベント
    fileInput.addEventListener('change', (event) => handleInternManagementFile(event));
    
    // アップロードボタンのクリックイベント
    uploadBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileInput.click();
    });
    
    // ドラッグ&ドロップイベント
    uploadArea.addEventListener('dragover', (e) => handleDragOver(e));
    uploadArea.addEventListener('dragleave', (e) => handleDragLeave(e));
    uploadArea.addEventListener('drop', (e) => handleInternManagementDrop(e));
    
    // エクスポートボタンを無効化
    disableExportButtons('3');
}

// エクスポートボタンを有効化する関数
function enableExportButtons(tabId) {
    console.log(`エクスポートボタンを有効化 - タブ ${tabId}`);
    
    if (tabId === '1') {
        const buttons = ['downloadProgressCSV1', 'downloadProcessedDataCSV1', 'downloadSortedByDeployDate1', 'downloadPowerPoint1'];
        buttons.forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (btn) {
                btn.disabled = false;
                btn.style.opacity = '1';
                btn.style.cursor = 'pointer';
            }
        });
    } else if (tabId === '2') {
        const buttons = ['downloadTaskCSV2', 'downloadTaskPowerPoint2'];
        buttons.forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (btn) {
                btn.disabled = false;
                btn.style.opacity = '1';
                btn.style.cursor = 'pointer';
            }
        });
    } else if (tabId === '3') {
        const buttons = ['downloadInternPowerPoint3'];
        buttons.forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (btn) {
                btn.disabled = false;
                btn.style.opacity = '1';
                btn.style.cursor = 'pointer';
            }
        });
    }
}

// エクスポートボタンを無効化する関数
function disableExportButtons(tabId) {
    console.log(`エクスポートボタンを無効化 - タブ ${tabId}`);
    
    if (tabId === '1') {
        const buttons = ['downloadProgressCSV1', 'downloadProcessedDataCSV1', 'downloadSortedByDeployDate1', 'downloadPowerPoint1'];
        buttons.forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (btn) {
                btn.disabled = true;
                btn.style.opacity = '0.5';
                btn.style.cursor = 'not-allowed';
            }
        });
    } else if (tabId === '2') {
        const buttons = ['downloadTaskCSV2', 'downloadTaskPowerPoint2'];
        buttons.forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (btn) {
                btn.disabled = true;
                btn.style.opacity = '0.5';
                btn.style.cursor = 'not-allowed';
            }
        });
    } else if (tabId === '3') {
        const buttons = ['downloadInternPowerPoint3'];
        buttons.forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (btn) {
                btn.disabled = true;
                btn.style.opacity = '0.5';
                btn.style.cursor = 'not-allowed';
            }
        });
    }
}

// 共通のドラッグ&ドロップ処理
function handleDragOver(event) {
    event.preventDefault();
    event.currentTarget.classList.add('dragover');
}

function handleDragLeave(event) {
    event.preventDefault();
    event.currentTarget.classList.remove('dragover');
}

// 部署別進捗分析のファイル処理
function handleProgressAnalysisFile(event) {
    const file = event.target.files[0];
    if (file) {
        console.log('部署別進捗分析ファイルが選択されました:', file.name);
        processProgressAnalysisFile(file);
        event.target.value = '';
    }
}

function handleProgressAnalysisDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        console.log('部署別進捗分析ファイルがドロップされました:', files[0].name);
        processProgressAnalysisFile(files[0]);
    }
}

// タスク管理のファイル処理
function handleTaskManagementFile(event) {
    const file = event.target.files[0];
    if (file) {
        console.log('タスク管理ファイルが選択されました:', file.name);
        processTaskManagementFile(file);
        event.target.value = '';
    }
}

function handleTaskManagementDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        console.log('タスク管理ファイルがドロップされました:', files[0].name);
        processTaskManagementFile(files[0]);
    }
}

// インターン生管理のファイル処理
function handleInternManagementFile(event) {
    const file = event.target.files[0];
    if (file) {
        console.log('インターン生管理ファイルが選択されました:', file.name);
        processInternManagementFile(file);
        event.target.value = '';
    }
}

function handleInternManagementDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        console.log('インターン生管理ファイルがドロップされました:', files[0].name);
        processInternManagementFile(files[0]);
    }
}

// 部署別進捗分析の処理
function processProgressAnalysisFile(file) {
    if (!file.name.toLowerCase().endsWith('.csv')) {
        showError('CSVファイルを選択してください。');
        return;
    }
    
    if (isProcessing) {
        console.log('既に処理中のため、新しいファイルの処理をスキップします');
        return;
    }
    
    isProcessing = true;
    showLoading();
    
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: function(results) {
            hideLoading();
            console.log('部署別進捗分析CSV解析完了:', results.data.length, '件のデータ');
            
            if (results.errors.length > 0) {
                console.error('CSV解析エラー:', results.errors);
                showError('CSVファイルの解析中にエラーが発生しました。');
                isProcessing = false;
                return;
            }
            
            if (results.data.length === 0) {
                showError('データが見つかりません。');
                isProcessing = false;
                return;
            }
            
            rawData = results.data;
            
            try {
                // データ処理を実行
                processProgressAnalysisData();
                clientAnalysisData = generateAnalysisData();
                displayProgressAnalysisResults();
                
                setTimeout(() => {
                    showSuccessMessage('部署別進捗分析が完了しました。エクスポート機能が利用可能です。');
                    enableExportButtons('1');
                    isProcessing = false;
                }, 500);
                
            } catch (error) {
                console.error('部署別進捗分析エラー:', error);
                showError('データ処理中にエラーが発生しました: ' + error.message);
                isProcessing = false;
            }
        },
        error: function(error) {
            hideLoading();
            console.error('ファイル読み込みエラー:', error);
            showError('ファイルの読み込みに失敗しました。');
            isProcessing = false;
        }
    });
}

// タスク管理の処理
function processTaskManagementFile(file) {
    if (!file.name.toLowerCase().endsWith('.csv')) {
        showError('CSVファイルを選択してください。');
        return;
    }
    
    if (isProcessing) {
        console.log('既に処理中のため、新しいファイルの処理をスキップします');
        return;
    }
    
    isProcessing = true;
    showLoading();
    
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: function(results) {
            hideLoading();
            console.log('タスク管理CSV解析完了:', results.data.length, '件のデータ');
            
            if (results.errors.length > 0) {
                console.error('CSV解析エラー:', results.errors);
                showError('CSVファイルの解析中にエラーが発生しました。');
                isProcessing = false;
                return;
            }
            
            if (results.data.length === 0) {
                showError('データが見つかりません。');
                isProcessing = false;
                return;
            }
            
            taskRawData = results.data;
            
            try {
                // タスクデータ処理を実行
                processTaskData();
                displayTaskManagementResults();
                
                setTimeout(() => {
                    showSuccessMessage('タスク管理処理が完了しました。エクスポート機能が利用可能です。');
                    enableExportButtons('2');
                    isProcessing = false;
                }, 500);
                
            } catch (error) {
                console.error('タスク管理処理エラー:', error);
                showError('データ処理中にエラーが発生しました: ' + error.message);
                isProcessing = false;
            }
        },
        error: function(error) {
            hideLoading();
            console.error('ファイル読み込みエラー:', error);
            showError('ファイルの読み込みに失敗しました。');
            isProcessing = false;
        }
    });
}

// インターン生管理の処理
function processInternManagementFile(file) {
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.csv')) {
        showError('Excel（.xlsx）またはCSVファイルを選択してください。');
        return;
    }
    
    if (isProcessing) {
        console.log('既に処理中のため、新しいファイルの処理をスキップします');
        return;
    }
    
    isProcessing = true;
    showLoading();
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            let data;
            
            if (fileName.endsWith('.csv')) {
                // CSVファイルの処理
                const csvText = e.target.result;
                const workbook = XLSX.read(csvText, { type: 'string' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            } else {
                // Excelファイルの処理
                const workbook = XLSX.read(e.target.result, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            }
            
            console.log('インターン生データ解析完了:', data.length, '行のデータ');
            
            // インターン生データ処理を実行
            processInternData(data);
            displayInternManagementResults();
            
            hideLoading();
            setTimeout(() => {
                showSuccessMessage('インターン生管理処理が完了しました。PowerPointスライドを生成できます。');
                enableExportButtons('3');
                isProcessing = false;
            }, 500);
            
        } catch (error) {
            console.error('インターン生管理処理エラー:', error);
            hideLoading();
            showError('データ処理中にエラーが発生しました: ' + error.message);
            isProcessing = false;
        }
    };
    
    if (fileName.endsWith('.csv')) {
        reader.readAsText(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

// 部署別進捗分析のデータ処理
function processProgressAnalysisData() {
    console.log('部署別進捗分析データ処理を開始');
    
    processedData = rawData.map((row, index) => {
        // 進捗状況の分類
        const progressCategory = categorizeProgressStatus(row['進捗状況']);
        
        return {
            ...row,
            '進捗カテゴリ': progressCategory
        };
    });
    
    // 部署の抽出
    departments = [...new Set(processedData.map(row => row['部署']).filter(Boolean))];
    
    console.log('部署別進捗分析データ処理完了:', processedData.length, '件処理済み,', departments.length, '部署');
}

// 進捗状況の分類
function categorizeProgressStatus(status) {
    if (!status) return '未着手';
    
    const statusStr = status.toString().toLowerCase();
    
    // リリース済み
    if (statusStr.includes('完了') || statusStr.includes('完成') || statusStr.includes('終了') || 
        statusStr.includes('リリース済み') || statusStr.includes('公開済み')) {
        return 'リリース済み';
    }
    
    // リリース準備中
    if (statusStr.includes('fb待ち') || statusStr.includes('公開可能') || 
        statusStr.includes('公開済み（アップデート予定）') || statusStr.includes('川島さんに確認待ち')) {
        return 'リリース準備中';
    }
    
    // 開発中
    if (statusStr.includes('実装') || statusStr.includes('設計') || 
        statusStr.includes('進行中') || statusStr.includes('作業中') || statusStr.includes('実施中')) {
        return '開発中';
    }
    
    // 検討
    if (statusStr.includes('検討中') || statusStr.includes('検討') || 
        statusStr.includes('計画中') || statusStr.includes('推奨（宮川案）')) {
        return '検討';
    }
    
    // 保留
    if (statusStr.includes('保留') || statusStr.includes('一時停止') || statusStr.includes('停止')) {
        return '保留';
    }
    
    // 中断
    if (statusStr.includes('中断') || statusStr.includes('中止') || statusStr.includes('キャンセル')) {
        return '中断';
    }
    
    // 未着手・未確認は除外対象
    if (statusStr.includes('未着手') || statusStr.includes('未確認')) {
        return '除外対象';
    }
    
    return '除外対象';
}

// クライアント側で分析データを生成
function generateAnalysisData() {
    console.log('クライアント側分析データ生成開始');
    
    const departmentMetrics = calculateDepartmentMetrics(processedData);
    const priorityTasks = identifyPriorityTasks(processedData);
    const efficiencyReport = generateEfficiencyReport(departmentMetrics);
    
    return {
        total_tasks: processedData.filter(row => row['進捗カテゴリ'] !== '除外対象').length,
        departments: departments,
        department_metrics: departmentMetrics,
        priority_tasks: priorityTasks,
        efficiency_report: efficiencyReport
    };
}

// 部署別メトリクスを計算
function calculateDepartmentMetrics(data) {
    const metrics = {};
    
    departments.forEach(dept => {
        const deptData = data.filter(row => row['部署'] === dept && row['進捗カテゴリ'] !== '除外対象');
        
        const released = deptData.filter(row => row['進捗カテゴリ'] === 'リリース済み').length;
        const releaseReady = deptData.filter(row => row['進捗カテゴリ'] === 'リリース準備中').length;
        const inDevelopment = deptData.filter(row => row['進捗カテゴリ'] === '開発中').length;
        const underConsideration = deptData.filter(row => row['進捗カテゴリ'] === '検討').length;
        const suspended = deptData.filter(row => row['進捗カテゴリ'] === '中断').length;
        const onHold = deptData.filter(row => row['進捗カテゴリ'] === '保留').length;
        
        const developmentTotal = released + releaseReady + inDevelopment;
        const considerationTotal = underConsideration + suspended + onHold;
        
        const totalHours = deptData.reduce((sum, row) => sum + (parseFloat(row['月工数（h）']) || 0), 0);
        const totalRevenue = deptData.reduce((sum, row) => sum + (parseFloat(row['収益インパクト(月の工数＊単価)']) || 0), 0);
        
        const completionRate = deptData.length > 0 ? (released / deptData.length * 100) : 0;
        
        metrics[dept] = {
            total_tasks: deptData.length,
            released: released,
            release_ready: releaseReady,
            in_development: inDevelopment,
            under_consideration: underConsideration,
            suspended: suspended,
            on_hold: onHold,
            development_total: developmentTotal,
            consideration_total: considerationTotal,
            completion_rate: completionRate,
            total_hours: totalHours,
            total_revenue: totalRevenue
        };
    });
    
    return metrics;
}

// 優先度の高いタスクを特定
function identifyPriorityTasks(data, topN = 10) {
    const validData = data.filter(row => row['進捗カテゴリ'] !== '除外対象');
    
    const sortedData = validData.sort((a, b) => {
        const aRevenue = parseFloat(a['収益インパクト(月の工数＊単価)']) || 0;
        const bRevenue = parseFloat(b['収益インパクト(月の工数＊単価)']) || 0;
        return bRevenue - aRevenue;
    });
    
    return sortedData.slice(0, topN).map(row => ({
        'タスク内容': row['タスク内容'] || '',
        '部署': row['部署'] || '',
        '進捗状況': row['進捗状況'] || '',
        '収益インパクト(月の工数＊単価)': parseFloat(row['収益インパクト(月の工数＊単価)']) || 0,
        '月工数（h）': parseFloat(row['月工数（h）']) || 0
    }));
}

// 効率性レポートを生成
function generateEfficiencyReport(departmentMetrics) {
    const report = {
        '生成日時': new Date().toLocaleString('ja-JP'),
        '分析対象部署数': Object.keys(departmentMetrics).length,
        '効率性ランキング': [],
        '改善推奨部署': [],
        '優秀部署': []
    };
    
    // 効率性ランキング
    const efficiencyRanking = [];
    Object.entries(departmentMetrics).forEach(([dept, metrics]) => {
        if (metrics.total_hours > 0) {
            const efficiency = metrics.total_revenue / metrics.total_hours;
            efficiencyRanking.push({
                '部署': dept,
                '効率性（時給）': efficiency,
                '総月工数': metrics.total_hours,
                '総収益インパクト': metrics.total_revenue
            });
        }
    });
    
    efficiencyRanking.sort((a, b) => b['効率性（時給）'] - a['効率性（時給）']);
    report['効率性ランキング'] = efficiencyRanking;
    
    return report;
}

// 部署別進捗分析の結果表示
function displayProgressAnalysisResults() {
    console.log('部署別進捗分析結果表示を開始');
    
    updateProgressAnalysisSummary();
    
    const resultSection = document.getElementById('resultSection1');
    if (resultSection) {
        resultSection.style.display = 'block';
    }
}

// サマリー情報の更新
function updateProgressAnalysisSummary() {
    const processedDepartmentsElement = document.getElementById('processedDepartments1');
    const totalTasksElement = document.getElementById('totalTasks1');
    const completionRateElement = document.getElementById('completionRate1');
    
    const validTasks = processedData.filter(row => row['進捗カテゴリ'] !== '除外対象');
    
    if (processedDepartmentsElement) {
        processedDepartmentsElement.textContent = departments.length;
    }
    if (totalTasksElement) {
        totalTasksElement.textContent = validTasks.length;
    }
    
    const completedTasks = validTasks.filter(row => row['進捗カテゴリ'] === 'リリース済み').length;
    const completionRate = validTasks.length > 0 ? (completedTasks / validTasks.length * 100).toFixed(1) : 0;
    
    if (completionRateElement) {
        completionRateElement.textContent = `${completionRate}%`;
    }
}

// タスク管理のデータ処理
function processTaskData() {
    console.log('タスク管理データ処理を開始');
    
    const progressMapping = {
        'リリース済み': ['公開済み（アップデート予定）', '完了'],
        'リリース準備': ['公開可能（川島さん確認待ち）', 'FB待ち'],
        '開発中': ['推奨（宮川案）', '設計', '実装']
    };
    
    const deleteStatus = ['未着手', '未確認', '検討', '保留', '中断'];
    
    const filteredResults = taskRawData.filter(row => 
        !deleteStatus.includes(row['進捗状況'])
    );
    
    console.log(`タスクフィルタリング後: ${filteredResults.length} 件（元データ: ${taskRawData.length} 件）`);
    
    function classifyProgress(status) {
        for (const [category, values] of Object.entries(progressMapping)) {
            for (const value of values) {
                if (String(status).includes(value)) {
                    return category;
                }
            }
        }
        return '開発中';
    }
    
    filteredResults.forEach(row => {
        row['進捗状況_分類'] = classifyProgress(row['進捗状況']);
    });
    
    taskClassificationData = {};
    filteredResults.forEach(row => {
        const status = row['進捗状況_分類'];
        taskClassificationData[status] = (taskClassificationData[status] || 0) + 1;
    });
    
    taskProcessedData = filteredResults;
    
    console.log('タスク管理データ処理完了:', taskProcessedData.length, '件処理済み');
}

// タスク管理の結果表示
function displayTaskManagementResults() {
    console.log('タスク管理結果表示を開始');
    
    updateTaskManagementSummary();
    displayTaskClassificationResults();
    
    const resultSection = document.getElementById('resultSection2');
    if (resultSection) {
        resultSection.style.display = 'block';
    }
}

// タスク管理のサマリー更新
function updateTaskManagementSummary() {
    const totalTasksElement = document.getElementById('totalTasks2');
    const processedTasksElement = document.getElementById('processedTasks2');
    const excludedTasksElement = document.getElementById('excludedTasks2');
    
    const totalTasks = taskRawData.length;
    const processedTasks = taskProcessedData.length;
    const excludedTasks = totalTasks - processedTasks;
    
    if (totalTasksElement) {
        totalTasksElement.textContent = totalTasks;
    }
    if (processedTasksElement) {
        processedTasksElement.textContent = processedTasks;
    }
    if (excludedTasksElement) {
        excludedTasksElement.textContent = excludedTasks;
    }
}

// タスク分類結果の表示
function displayTaskClassificationResults() {
    const classificationSummary = document.getElementById('classificationSummary2');
    if (!classificationSummary) return;
    
    const summaryHtml = Object.entries(taskClassificationData).map(([status, count]) => {
        return `
            <div class="classification-item">
                <span class="label">${status}</span>
                <span class="value">${count}件</span>
            </div>
        `;
    }).join('');
    
    classificationSummary.innerHTML = `<div class="classification-summary">${summaryHtml}</div>`;
}

// インターン生データの処理
function processInternData(data) {
    console.log('インターン生データ処理を開始');
    
    internMembersData = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        if (row && row.length >= 6) {
            const startMonthNum = parseInt(row[5]) || 6;
            const startMonth = `${startMonthNum}月入社`;
            
            const member = {
                id: i,
                name: row[1] || '',
                furigana: row[2] || '',
                age: parseInt(row[3]) || 0,
                affiliation: row[4] || '',
                startMonth: startMonth,
                description: row[6] || '',
                imagePath: row[0] ? `./sample/images/${row[0]}` : `./images/intern${i}.jpg`,
                highlight: null
            };
            internMembersData.push(member);
        }
    }
    
    console.log(`${internMembersData.length}人のメンバーデータを読み込みました。`);
}

// インターン生管理の結果表示
function displayInternManagementResults() {
    console.log('インターン生管理結果表示を開始');
    
    updateInternManagementSummary();
    displayMemberPreview();
    
    const resultSection = document.getElementById('resultSection3');
    if (resultSection) {
        resultSection.style.display = 'block';
    }
}

// インターン生管理のサマリー更新
function updateInternManagementSummary() {
    const totalMembersElement = document.getElementById('totalMembers3');
    const totalSlidesElement = document.getElementById('totalSlides3');
    
    const totalMembers = internMembersData.length;
    const totalSlides = Math.ceil(totalMembers / 6);
    
    if (totalMembersElement) {
        totalMembersElement.textContent = totalMembers;
    }
    if (totalSlidesElement) {
        totalSlidesElement.textContent = totalSlides;
    }
}

// メンバー一覧プレビューの表示
function displayMemberPreview() {
    const previewList = document.getElementById('memberPreviewList3');
    if (!previewList) return;
    
    const memberCards = internMembersData.map(member => {
        return `
            <div class="member-card">
                <div class="furigana">${member.furigana}</div>
                <div class="name">${member.name}</div>
                <div class="details">${member.age}歳 / ${member.affiliation}</div>
                <div class="description">${member.description}</div>
                <div class="start-month">${member.startMonth}</div>
            </div>
        `;
    }).join('');
    
    previewList.innerHTML = `<div class="member-list">${memberCards}</div>`;
}

// 共通ユーティリティ関数
function showLoading() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
        loadingOverlay.style.display = 'flex';
    }
}

function hideLoading() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
        loadingOverlay.style.display = 'none';
    }
}

function showError(message) {
    console.error('エラー:', message);
    const errorMessage = document.getElementById('errorMessage');
    const errorModal = document.getElementById('errorModal');
    
    if (errorMessage) {
        errorMessage.textContent = message;
    }
    if (errorModal) {
        errorModal.style.display = 'flex';
    }
}

function closeModal() {
    const errorModal = document.getElementById('errorModal');
    if (errorModal) {
        errorModal.style.display = 'none';
    }
}

function showSuccessMessage(message) {
    console.log('成功:', message);
    
    let successModal = document.getElementById('successModal');
    if (!successModal) {
        successModal = document.createElement('div');
        successModal.id = 'successModal';
        successModal.className = 'modal';
        successModal.innerHTML = `
            <div class="modal-content">
                <div class="modal-header">
                    <h3><i class="fas fa-check-circle"></i> 成功</h3>
                    <button class="close-btn" onclick="closeSuccessModal()">&times;</button>
                </div>
                <div class="modal-body">
                    <p id="successMessage"></p>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-primary" onclick="closeSuccessModal()">OK</button>
                </div>
            </div>
        `;
        document.body.appendChild(successModal);
    }
    
    const successMessage = document.getElementById('successMessage');
    if (successMessage) {
        successMessage.textContent = message;
    }
    
    successModal.style.display = 'flex';
    
    setTimeout(() => {
        closeSuccessModal();
    }, 3000);
}

function closeSuccessModal() {
    const successModal = document.getElementById('successModal');
    if (successModal) {
        successModal.style.display = 'none';
    }
}

// === エクスポート関数の実装 ===

// CSV変換関数
function convertToCSV(data) {
    console.log('CSV変換を開始:', data.length, '件のデータ');
    
    if (data.length === 0) {
        console.warn('変換するデータがありません');
        return '';
    }
    
    try {
        const headers = Object.keys(data[0]);
        console.log('CSVヘッダー:', headers);
        
        const csvRows = [headers.join(',')];
        
        data.forEach((row, index) => {
            const values = headers.map(header => {
                const value = row[header];
                if (value === null || value === undefined) return '';
                const stringValue = value.toString();
                // カンマ、改行、ダブルクォートを含む場合はダブルクォートで囲む
                if (stringValue.includes(',') || stringValue.includes('\n') || stringValue.includes('"')) {
                    return `"${stringValue.replace(/"/g, '""')}"`;
                }
                return stringValue;
            });
            csvRows.push(values.join(','));
        });
        
        const result = csvRows.join('\n');
        console.log('CSV変換完了:', result.length, '文字');
        return result;
    } catch (error) {
        console.error('CSV変換エラー:', error);
        throw error;
    }
}

// CSVダウンロード関数
function downloadCSV(content, filename) {
    console.log('CSVダウンロードを開始:', filename, 'サイズ:', content.length, '文字');
    
    try {
        // BOMを追加してUTF-8エンコーディングを明示
        const bom = '\uFEFF';
        const blob = new Blob([bom + content], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        
        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', filename);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            
            console.log('CSVダウンロード完了:', filename);
        } else {
            console.error('ダウンロード機能がサポートされていません');
            showError('お使いのブラウザではダウンロード機能がサポートされていません。');
        }
    } catch (error) {
        console.error('CSVダウンロードエラー:', error);
        showError('ファイルのダウンロードに失敗しました: ' + error.message);
    }
}

// 部署別進捗CSVダウンロード
function downloadProgressCSV() {
    console.log('=== 部署別進捗状況CSVのダウンロード開始 ===');
    
    if (!clientAnalysisData || !clientAnalysisData.department_metrics) {
        showError('分析データがありません。CSVファイルを再アップロードしてください。');
        return;
    }
    
    try {
        const departmentMetrics = clientAnalysisData.department_metrics;
        const progressData = Object.keys(departmentMetrics).map(dept => {
            const metrics = departmentMetrics[dept];
            
            return {
                '部署': dept,
                '総タスク数': metrics.total_tasks || 0,
                'リリース済み': metrics.released || 0,
                'リリース準備中': metrics.release_ready || 0,
                '開発中': metrics.in_development || 0,
                '検討': metrics.under_consideration || 0,
                '中断': metrics.suspended || 0,
                '保留': metrics.on_hold || 0,
                '開発対象_小計': metrics.development_total || 0,
                '開発検討_小計': metrics.consideration_total || 0,
                '完了率（%）': metrics.completion_rate.toFixed(1),
                '総月工数（h）': metrics.total_hours.toFixed(1),
                '総収益インパクト（円）': metrics.total_revenue,
                '処理日時': new Date().toLocaleString('ja-JP')
            };
        });
        
        const csvContent = convertToCSV(progressData);
        downloadCSV(csvContent, '部署別進捗状況.csv');
        
        console.log('=== 部署別進捗状況CSVのダウンロード完了 ===');
    } catch (error) {
        console.error('CSVダウンロードエラー:', error);
        showError('CSVファイルの作成中にエラーが発生しました: ' + error.message);
    }
}

// 処理済みデータCSVダウンロード
function downloadProcessedDataCSV() {
    console.log('=== 処理済みデータCSVのダウンロード開始 ===');
    
    if (!clientAnalysisData) {
        showError('分析データがありません。CSVファイルを再アップロードしてください。');
        return;
    }
    
    try {
        const filteredData = processedData.filter(row => row['進捗カテゴリ'] !== '除外対象');
        
        console.log('除外対象タスクを除外:', processedData.length - filteredData.length, '件');
        
        const csvContent = convertToCSV(filteredData);
        downloadCSV(csvContent, '処理済みデータ.csv');
        
        console.log('=== 処理済みデータCSVのダウンロード完了 ===');
    } catch (error) {
        console.error('CSVダウンロードエラー:', error);
        showError('CSVファイルの作成中にエラーが発生しました: ' + error.message);
    }
}

// デプロイ日順ソートCSVダウンロード
function downloadSortedByDeployDate() {
    console.log('=== デプロイ日順CSVエクスポート開始 ===');
    
    if (!clientAnalysisData || !rawData.length) {
        showError('分析データがありません。CSVファイルを再アップロードしてください。');
        return;
    }
    
    showLoading();
    
    try {
        const sortedData = sortByDeployDate(rawData);
        
        if (!sortedData || sortedData.length === 0) {
            throw new Error('ソート処理に失敗しました');
        }
        
        console.log('デプロイ日順ソート完了:', sortedData.length, '件');
        
        const csvContent = convertToCSV(sortedData);
        const filename = `デプロイ日順ソート_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.csv`;
        downloadCSV(csvContent, filename);
        
        hideLoading();
        console.log('デプロイ日順CSVエクスポート完了');
        
    } catch (error) {
        console.error('デプロイ日順CSVエクスポートエラー:', error);
        hideLoading();
        showError('デプロイ日順CSVの生成に失敗しました: ' + error.message);
    }
}

// デプロイ日順でデータをソート
function sortByDeployDate(data) {
    try {
        console.log('デプロイ日順ソート開始:', data.length, '件のデータ');
        
        const deployDateColumns = ['デプロイ日時', 'デプロイ日', 'リリース日', '公開日'];
        let deployColumn = null;
        
        for (const col of deployDateColumns) {
            if (data.length > 0 && col in data[0]) {
                deployColumn = col;
                console.log('デプロイ日カラム発見:', col);
                break;
            }
        }
        
        if (!deployColumn) {
            console.log('デプロイ日関連のカラムが見つかりません');
            return data;
        }
        
        function normalizeDate(dateStr) {
            if (!dateStr || dateStr === '') {
                return null;
            }
            
            const str = String(dateStr).trim();
            
            try {
                if (str.includes(',') && /\b(January|February|March|April|May|June|July|August|September|October|November|December)\b/.test(str)) {
                    return new Date(str);
                }
                
                if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) {
                    return new Date(str);
                }
                
                if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
                    const parts = str.split('/');
                    return new Date(parts[2], parts[1] - 1, parts[0]);
                }
                
                const date = new Date(str);
                return isNaN(date.getTime()) ? null : date;
                
            } catch (e) {
                return null;
            }
        }
        
        const sortedData = [...data].sort((a, b) => {
            const dateA = normalizeDate(a[deployColumn]);
            const dateB = normalizeDate(b[deployColumn]);
            
            if (!dateA && !dateB) return 0;
            if (!dateA) return 1;
            if (!dateB) return -1;
            
            return dateA.getTime() - dateB.getTime();
        });
        
        return sortedData;
        
    } catch (error) {
        console.error('デプロイ日順ソートエラー:', error);
        return data;
    }
}

// PowerPoint生成関数群
function downloadPowerPoint() {
    console.log('=== PowerPointダウンロード開始 ===');
    
    if (!clientAnalysisData || !clientAnalysisData.department_metrics) {
        showError('分析データがありません。CSVファイルを再アップロードしてください。');
        return;
    }
    
    if (typeof PptxGenJS === 'undefined') {
        showError('PowerPoint機能が利用できません。ページを再読み込みしてください。');
        return;
    }
    
    showLoading();
    
    try {
        const pptx = new PptxGenJS();
        
        pptx.author = '部署別進捗状況CSV変換システム';
        pptx.company = '開発チーム';
        pptx.title = '部署別進捗状況レポート';
        pptx.subject = '部署別タスクの進捗状況';
        
        createTitleSlide(pptx);
        createProgressTableSlide(pptx);
        createSummarySlide(pptx);
        
        const filename = `部署別進捗状況レポート_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.pptx`;
        pptx.writeFile({ fileName: filename }).then(() => {
            console.log('PowerPointダウンロード完了');
            hideLoading();
        }).catch(error => {
            console.error('PowerPoint保存エラー:', error);
            hideLoading();
            showError('PowerPointファイルの保存に失敗しました: ' + error.message);
        });
        
    } catch (error) {
        console.error('PowerPoint生成エラー:', error);
        hideLoading();
        showError('PowerPointファイルの生成に失敗しました: ' + error.message);
    }
}

// タイトルスライドを作成
function createTitleSlide(pptx) {
    const slide = pptx.addSlide();
    
    slide.addText('部署別進捗状況レポート', {
        x: 1, y: 2, w: 8, h: 1.5,
        fontSize: 36, bold: true, align: 'center',
        color: '2F4F4F'
    });
    
    slide.addText(`作成日時: ${new Date().toLocaleString('ja-JP')}`, {
        x: 1, y: 4, w: 8, h: 0.5,
        fontSize: 16, align: 'center',
        color: '666666'
    });
    
    const validTasks = processedData.filter(row => row['進捗カテゴリ'] !== '除外対象').length;
    const completedTasks = processedData.filter(row => row['進捗カテゴリ'] === 'リリース済み').length;
    const completionRate = validTasks > 0 ? (completedTasks / validTasks * 100).toFixed(1) : 0;
    
    slide.addText([
        { text: '概要\n', options: { fontSize: 20, bold: true, color: '2F4F4F' } },
        { text: `・対象部署数: ${departments.length}部署\n`, options: { fontSize: 16 } },
        { text: `・総タスク数: ${validTasks}件\n`, options: { fontSize: 16 } },
        { text: `・完了タスク数: ${completedTasks}件\n`, options: { fontSize: 16 } },
        { text: `・完了率: ${completionRate}%`, options: { fontSize: 16 } }
    ], {
        x: 1, y: 5, w: 8, h: 2,
        align: 'left'
    });
}

// 部署別進捗状況テーブルスライドを作成
function createProgressTableSlide(pptx) {
    const slide = pptx.addSlide();
    
    slide.addText('部署別進捗状況', {
        x: 0.5, y: 0.3, w: 9, h: 0.7,
        fontSize: 24, bold: true, align: 'center',
        color: '2F4F4F'
    });
    
    const tableData = [
        ['部署', 'リリース済み', 'リリース準備中', '開発中', '開発対象小計', '検討', '中断', '保留', '開発検討小計', '総計', '完了率(%)']
    ];
    
    Object.entries(clientAnalysisData.department_metrics).forEach(([dept, metrics]) => {
        tableData.push([
            dept,
            metrics.released.toString(),
            metrics.release_ready.toString(),
            metrics.in_development.toString(),
            metrics.development_total.toString(),
            metrics.under_consideration.toString(),
            metrics.suspended.toString(),
            metrics.on_hold.toString(),
            metrics.consideration_total.toString(),
            (metrics.development_total + metrics.consideration_total).toString(),
            metrics.completion_rate.toFixed(1)
        ]);
    });
    
    const totalMetrics = calculateTotalMetrics();
    tableData.push([
        '総計',
        totalMetrics.released.toString(),
        totalMetrics.release_ready.toString(),
        totalMetrics.in_development.toString(),
        totalMetrics.development_total.toString(),
        totalMetrics.under_consideration.toString(),
        totalMetrics.suspended.toString(),
        totalMetrics.on_hold.toString(),
        totalMetrics.consideration_total.toString(),
        (totalMetrics.development_total + totalMetrics.consideration_total).toString(),
        totalMetrics.completion_rate
    ]);
    
    slide.addTable(tableData, {
        x: 0.2, y: 1.2, w: 9.6, h: 5.5,
        colW: [1.2, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8],
        border: { pt: 1, color: 'CCCCCC' },
        fill: { color: 'F9F9F9' },
        rowH: 0.4,
        fontSize: 10
    });
}

// サマリスライドを作成
function createSummarySlide(pptx) {
    const slide = pptx.addSlide();
    
    slide.addText('分析サマリ', {
        x: 0.5, y: 0.3, w: 9, h: 0.7,
        fontSize: 24, bold: true, align: 'center',
        color: '2F4F4F'
    });
    
    const validTasks = processedData.filter(row => row['進捗カテゴリ'] !== '除外対象');
    const developmentTasks = validTasks.filter(row => 
        ['リリース済み', 'リリース準備中', '開発中'].includes(row['進捗カテゴリ'])
    );
    const considerationTasks = validTasks.filter(row => 
        ['検討', '中断', '保留'].includes(row['進捗カテゴリ'])
    );
    const completedTasks = validTasks.filter(row => row['進捗カテゴリ'] === 'リリース済み');
    
    const summaryText = [
        { text: '全体統計\n', options: { fontSize: 18, bold: true, color: '2F4F4F' } },
        { text: `依頼数: ${validTasks.length}件\n`, options: { fontSize: 14 } },
        { text: `開発対象: ${developmentTasks.length}件（うち完了${completedTasks.length}件）\n`, options: { fontSize: 14 } },
        { text: `検討: ${considerationTasks.length}件\n\n`, options: { fontSize: 14 } },
        
        { text: '進捗状況別内訳\n', options: { fontSize: 18, bold: true, color: '2F4F4F' } },
        { text: `・リリース済み: ${completedTasks.length}件\n`, options: { fontSize: 14 } },
        { text: `・リリース準備中: ${validTasks.filter(row => row['進捗カテゴリ'] === 'リリース準備中').length}件\n`, options: { fontSize: 14 } },
        { text: `・開発中: ${validTasks.filter(row => row['進捗カテゴリ'] === '開発中').length}件\n`, options: { fontSize: 14 } },
        { text: `・検討: ${validTasks.filter(row => row['進捗カテゴリ'] === '検討').length}件\n`, options: { fontSize: 14 } },
        { text: `・中断: ${validTasks.filter(row => row['進捗カテゴリ'] === '中断').length}件\n`, options: { fontSize: 14 } },
        { text: `・保留: ${validTasks.filter(row => row['進捗カテゴリ'] === '保留').length}件`, options: { fontSize: 14 } }
    ];
    
    slide.addText(summaryText, {
        x: 1, y: 1.5, w: 8, h: 5,
        align: 'left'
    });
}

// 総計メトリクスを計算
function calculateTotalMetrics() {
    const totalMetrics = {
        released: 0,
        release_ready: 0,
        in_development: 0,
        under_consideration: 0,
        suspended: 0,
        on_hold: 0,
        development_total: 0,
        consideration_total: 0
    };
    
    Object.values(clientAnalysisData.department_metrics).forEach(metrics => {
        totalMetrics.released += metrics.released;
        totalMetrics.release_ready += metrics.release_ready;
        totalMetrics.in_development += metrics.in_development;
        totalMetrics.under_consideration += metrics.under_consideration;
        totalMetrics.suspended += metrics.suspended;
        totalMetrics.on_hold += metrics.on_hold;
        totalMetrics.development_total += metrics.development_total;
        totalMetrics.consideration_total += metrics.consideration_total;
    });
    
    const totalTasks = totalMetrics.development_total + totalMetrics.consideration_total;
    totalMetrics.completion_rate = totalTasks > 0 ? 
        (totalMetrics.released / totalTasks * 100).toFixed(1) : '0.0';
    
    return totalMetrics;
}

// タスク管理のエクスポート関数
function downloadTaskCSV() {
    if (!taskProcessedData.length) {
        showError('タスクデータがありません。');
        return;
    }
    
    try {
        const csvContent = convertToCSV(taskProcessedData);
        const filename = `タスク管理データ_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.csv`;
        downloadCSV(csvContent, filename);
    } catch (error) {
        console.error('タスクCSVダウンロードエラー:', error);
        showError('CSVファイルの作成中にエラーが発生しました: ' + error.message);
    }
}

function downloadTaskPowerPoint() {
    if (!taskProcessedData.length) {
        showError('タスクデータがありません。');
        return;
    }
    
    if (typeof PptxGenJS === 'undefined') {
        showError('PowerPoint機能が利用できません。ページを再読み込みしてください。');
        return;
    }
    
    showLoading();
    
    try {
        const pptx = new PptxGenJS();
        
        pptx.author = 'タスク管理システム';
        pptx.company = '開発チーム';
        pptx.title = 'タスク管理レポート';
        pptx.subject = 'タスクの進捗状況分析';
        
        createTaskTitleSlide(pptx);
        createTaskSummarySlide(pptx);
        
        const filename = `タスク管理レポート_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.pptx`;
        pptx.writeFile({ fileName: filename }).then(() => {
            console.log('タスクPowerPointダウンロード完了');
            hideLoading();
        }).catch(error => {
            console.error('タスクPowerPoint保存エラー:', error);
            hideLoading();
            showError('PowerPointファイルの保存に失敗しました: ' + error.message);
        });
        
    } catch (error) {
        console.error('タスクPowerPoint生成エラー:', error);
        hideLoading();
        showError('PowerPointファイルの生成に失敗しました: ' + error.message);
    }
}

function createTaskTitleSlide(pptx) {
    const slide = pptx.addSlide();
    
    slide.addText('タスク管理レポート', {
        x: 1, y: 2, w: 8, h: 1.5,
        fontSize: 36, bold: true, align: 'center',
        color: '2F4F4F'
    });
    
    slide.addText(`作成日時: ${new Date().toLocaleString('ja-JP')}`, {
        x: 1, y: 4, w: 8, h: 0.5,
        fontSize: 16, align: 'center',
        color: '666666'
    });
    
    slide.addText([
        { text: '概要\n', options: { fontSize: 20, bold: true, color: '2F4F4F' } },
        { text: `・総タスク数: ${taskRawData.length}件\n`, options: { fontSize: 16 } },
        { text: `・処理済みタスク数: ${taskProcessedData.length}件\n`, options: { fontSize: 16 } },
        { text: `・除外タスク数: ${taskRawData.length - taskProcessedData.length}件`, options: { fontSize: 16 } }
    ], {
        x: 1, y: 5, w: 8, h: 2,
        align: 'left'
    });
}

function createTaskSummarySlide(pptx) {
    const slide = pptx.addSlide();
    
    slide.addText('進捗状況分類結果', {
        x: 0.5, y: 0.3, w: 9, h: 0.7,
        fontSize: 24, bold: true, align: 'center',
        color: '2F4F4F'
    });
    
    const summaryData = Object.entries(taskClassificationData).map(([status, count]) => [status, count.toString()]);
    
    slide.addTable([['進捗状況', '件数'], ...summaryData], {
        x: 2, y: 2, w: 6, h: 4,
        colW: [4, 2],
        border: { pt: 1, color: 'CCCCCC' },
        fill: { color: 'F9F9F9' },
        rowH: 0.5,
        fontSize: 14
    });
}

// インターン生管理のPowerPoint生成
function downloadInternPowerPoint() {
    if (!internMembersData.length) {
        showError('インターン生データがありません。');
        return;
    }
    
    if (typeof PptxGenJS === 'undefined') {
        showError('PowerPoint機能が利用できません。ページを再読み込みしてください。');
        return;
    }
    
    showLoading();
    
    try {
        const pptx = new PptxGenJS();
        
        pptx.author = 'インターン生管理システム';
        pptx.company = '開発チーム';
        pptx.title = 'インターン生メンバー一覧';
        pptx.subject = 'インターン生のプロフィール';
        
        generateInternSlides(pptx, internMembersData);
        
        const filename = `インターン生メンバー一覧_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.pptx`;
        pptx.writeFile({ fileName: filename }).then(() => {
            console.log('インターン生PowerPointダウンロード完了');
            hideLoading();
        }).catch(error => {
            console.error('インターン生PowerPoint保存エラー:', error);
            hideLoading();
            showError('PowerPointファイルの保存に失敗しました: ' + error.message);
        });
        
    } catch (error) {
        console.error('インターン生PowerPoint生成エラー:', error);
        hideLoading();
        showError('PowerPointファイルの生成に失敗しました: ' + error.message);
    }
}

function generateInternSlides(pptx, members) {
    const membersPerSlide = 6;
    const totalSlides = Math.ceil(members.length / membersPerSlide);
    
    console.log(`${members.length}人のメンバーを${totalSlides}スライドに分割します（1スライドにつき${membersPerSlide}人まで）`);
    
    for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
        const slide = pptx.addSlide();
        
        slide.background = { color: 'FFFFFF' };
        
        const startIndex = slideIndex * membersPerSlide;
        const endIndex = Math.min(startIndex + membersPerSlide, members.length);
        const slideMembers = members.slice(startIndex, endIndex);
        
        console.log(`スライド${slideIndex + 1}: メンバー${startIndex + 1}〜${endIndex}（${slideMembers.length}人）`);
        
        const groupedMembers = groupMembersByStartMonth(slideMembers);
        layoutInternMembers(slide, groupedMembers);
    }
}

function groupMembersByStartMonth(members) {
    const grouped = {};
    members.forEach(member => {
        if (!grouped[member.startMonth]) {
            grouped[member.startMonth] = [];
        }
        grouped[member.startMonth].push(member);
    });
    return grouped;
}

function layoutInternMembers(slide, groupedMembers) {
    const startY = 0.5;
    const leftColumnX = 0.7;
    const rightColumnX = 5.5;
    const columnWidth = 4.3;
    const memberHeight = 1.6;
    const imageSize = 1.0;
    const textMargin = 0.15;

    let leftY = startY;
    let rightY = startY;

    const monthOrder = [];
    for (let i = 1; i <= 12; i++) {
        monthOrder.push(`${i}月入社`);
    }

    monthOrder.forEach(month => {
        if (!groupedMembers[month]) return;

        const members = groupedMembers[month];

        members.forEach((member, index) => {
            const isLeftColumn = leftY <= rightY;
            const x = isLeftColumn ? leftColumnX : rightColumnX;
            let y = isLeftColumn ? leftY : rightY;

            addStartMonthBar(slide, month, x, y, memberHeight);
            addMemberProfile(slide, member, x, y, columnWidth, memberHeight, imageSize, textMargin);

            if (isLeftColumn) {
                leftY += memberHeight + 0.2;
            } else {
                rightY += memberHeight + 0.2;
            }
        });
    });
}

function addStartMonthBar(slide, month, x, y, height) {
    slide.addShape('rect', {
        x: x - 0.3,
        y: y,
        w: 0.25,
        h: height,
        fill: { color: '4472C4' }
    });

    const verticalText = month.split('').join('\n');
    slide.addText(verticalText, {
        x: x - 0.4,
        y: y + height / 2 - 0.2,
        w: 0.35,
        h: 0.4,
        fontSize: 15,
        color: 'FFFFFF',
        align: 'center',
        valign: 'middle',
        bold: true
    });
}

function addMemberProfile(slide, member, x, y, width, height, imageSize, textMargin) {
    slide.addShape('rect', {
        x: x,
        y: y,
        w: width,
        h: height,
        fill: { color: 'F8F9FA' },
        line: { color: 'E9ECEF', width: 1 }
    });

    const imageX = x + textMargin;
    const imageY = y + textMargin;
    
    // プレースホルダー画像（実際の画像は表示できないため）
    slide.addShape('rect', {
        x: imageX,
        y: imageY,
        w: imageSize,
        h: imageSize,
        fill: { color: 'E9ECEF' },
        line: { color: 'CCCCCC', width: 1 }
    });
    
    slide.addText('画像', {
        x: imageX,
        y: imageY + imageSize/2 - 0.1,
        w: imageSize,
        h: 0.2,
        fontSize: 10,
        color: '999999',
        align: 'center',
        valign: 'middle'
    });

    const textX = imageX + imageSize + textMargin;
    const nameWidth = width - imageSize - (textMargin * 3);

    // ふりがな
    slide.addText(member.furigana, {
        x: textX,
        y: imageY,
        w: nameWidth,
        h: 0.2,
        fontSize: 10,
        color: '666666',
        align: 'left',
        valign: 'top'
    });

    // 名前・年齢・所属
    slide.addText(`${member.name} (${member.age}歳) ${member.affiliation}`, {
        x: textX,
        y: imageY + 0.2,
        w: nameWidth,
        h: 0.3,
        fontSize: 12,
        color: '333333',
        align: 'left',
        valign: 'top',
        bold: true
    });

    // 紹介文
    slide.addText(member.description, {
        x: textX,
        y: imageY + 0.5,
        w: nameWidth,
        h: height - 0.7,
        fontSize: 10,
        color: '555555',
        align: 'left',
        valign: 'top'
    });
} 