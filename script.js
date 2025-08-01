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
    displayDepartmentProgressTable();
    displayProgressCategoryStats();
    displayProcessedDataTable();
    
    const resultSection = document.getElementById('resultSection1');
    if (resultSection) {
        resultSection.style.display = 'block';
    }
}

// 部署別進捗状況テーブルの表示
function displayDepartmentProgressTable() {
    const table = document.getElementById('departmentProgressTable1');
    if (!table || !clientAnalysisData) return;
    
    const tbody = table.querySelector('tbody');
    tbody.innerHTML = '';
    
    // 部署別データを追加
    Object.entries(clientAnalysisData.department_metrics).forEach(([dept, metrics]) => {
        const row = document.createElement('tr');
        const completionRate = metrics.completion_rate.toFixed(1);
        const total = metrics.development_total + metrics.consideration_total;
        
        row.innerHTML = `
            <td><strong>${dept}</strong></td>
            <td>${metrics.released}</td>
            <td>${metrics.release_ready}</td>
            <td>${metrics.in_development}</td>
            <td><strong>${metrics.development_total}</strong></td>
            <td>${metrics.under_consideration}</td>
            <td>${metrics.suspended}</td>
            <td>${metrics.on_hold}</td>
            <td><strong>${metrics.consideration_total}</strong></td>
            <td><strong>${total}</strong></td>
            <td><strong>${completionRate}%</strong></td>
        `;
        tbody.appendChild(row);
    });
    
    // 総計行を追加
    const totalMetrics = calculateTotalMetrics();
    const totalRow = document.createElement('tr');
    totalRow.className = 'total-row';
    const grandTotal = totalMetrics.development_total + totalMetrics.consideration_total;
    
    totalRow.innerHTML = `
        <td><strong>総計</strong></td>
        <td><strong>${totalMetrics.released}</strong></td>
        <td><strong>${totalMetrics.release_ready}</strong></td>
        <td><strong>${totalMetrics.in_development}</strong></td>
        <td><strong>${totalMetrics.development_total}</strong></td>
        <td><strong>${totalMetrics.under_consideration}</strong></td>
        <td><strong>${totalMetrics.suspended}</strong></td>
        <td><strong>${totalMetrics.on_hold}</strong></td>
        <td><strong>${totalMetrics.consideration_total}</strong></td>
        <td><strong>${grandTotal}</strong></td>
        <td><strong>${totalMetrics.completion_rate}%</strong></td>
    `;
    tbody.appendChild(totalRow);
}

// 進捗カテゴリ統計の表示
function displayProgressCategoryStats() {
    const container = document.getElementById('progressCategoryStats1');
    if (!container || !clientAnalysisData) return;
    
    const totalMetrics = calculateTotalMetrics();
    const statsData = [
        { label: 'リリース済み', value: totalMetrics.released, class: 'stat-released' },
        { label: 'リリース準備中', value: totalMetrics.release_ready, class: 'stat-ready' },
        { label: '開発中', value: totalMetrics.in_development, class: 'stat-development' },
        { label: '検討', value: totalMetrics.under_consideration, class: 'stat-consideration' },
        { label: '中断', value: totalMetrics.suspended, class: 'stat-suspended' },
        { label: '保留', value: totalMetrics.on_hold, class: 'stat-hold' }
    ];
    
    container.innerHTML = statsData.map(stat => `
        <div class="stat-card ${stat.class}">
            <div class="stat-value">${stat.value}</div>
            <div class="stat-label">${stat.label}</div>
        </div>
    `).join('');
}

// ページネーション関連変数
let currentPage1 = 1;
let currentPage2 = 1;
const itemsPerPage = 20;
let filteredData1 = [];
let filteredData2 = [];

// 処理済みデータテーブルの表示
function displayProcessedDataTable() {
    if (!processedData.length) return;
    
    // フィルター選択肢を設定
    setupFilters1();
    
    // 初期データ表示
    filteredData1 = processedData.filter(row => row['進捗カテゴリ'] !== '除外対象');
    currentPage1 = 1;
    renderDataTable1();
    
    // フィルターイベントリスナー
    document.getElementById('dataFilter1').addEventListener('input', applyFilters1);
    document.getElementById('categoryFilter1').addEventListener('change', applyFilters1);
}

// フィルター設定
function setupFilters1() {
    const categoryFilter = document.getElementById('categoryFilter1');
    const categories = [...new Set(processedData.map(row => row['進捗カテゴリ']).filter(Boolean))];
    
    categoryFilter.innerHTML = '<option value="">すべてのカテゴリ</option>';
    categories.forEach(category => {
        if (category !== '除外対象') {
            categoryFilter.innerHTML += `<option value="${category}">${category}</option>`;
        }
    });
}

// フィルター適用
function applyFilters1() {
    const searchText = document.getElementById('dataFilter1').value.toLowerCase();
    const selectedCategory = document.getElementById('categoryFilter1').value;
    
    filteredData1 = processedData.filter(row => {
        if (row['進捗カテゴリ'] === '除外対象') return false;
        
        // テキスト検索
        const searchMatch = !searchText || Object.values(row).some(value => 
            String(value).toLowerCase().includes(searchText)
        );
        
        // カテゴリフィルター
        const categoryMatch = !selectedCategory || row['進捗カテゴリ'] === selectedCategory;
        
        return searchMatch && categoryMatch;
    });
    
    currentPage1 = 1;
    renderDataTable1();
}

// データテーブル描画
function renderDataTable1() {
    const table = document.getElementById('processedDataTable1');
    if (!table) return;
    
    const startIndex = (currentPage1 - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const pageData = filteredData1.slice(startIndex, endIndex);
    
    // ヘッダー設定
    if (pageData.length > 0) {
        const headers = Object.keys(pageData[0]);
        table.querySelector('thead').innerHTML = `
            <tr>${headers.map(header => `<th>${header}</th>`).join('')}</tr>
        `;
        
        // データ行設定
        const tbody = table.querySelector('tbody');
        tbody.innerHTML = pageData.map(row => `
            <tr>${headers.map(header => {
                let cellValue = row[header] || '';
                
                // 金額項目の表示形式調整
                if (header === 'インターン費用' || header === '外注時費用' || header === '収益貢献（年）') {
                    const numValue = parseFloat(cellValue);
                    if (numValue > 0) {
                        cellValue = `${numValue.toLocaleString()}円`;
                    } else {
                        cellValue = '0円';
                    }
                }
                // 進捗状況にバッジを適用
                else if (header === '進捗状況') {
                    cellValue = `<span class="status-badge status-${cellValue.replace(/\s+/g, '-')}">${cellValue}</span>`;
                }
                return `<td>${cellValue}</td>`;
            }).join('')}</tr>
        `).join('');
    }
    
    // ページネーション更新
    updatePagination1();
}

// ページネーション更新
function updatePagination1() {
    const totalPages = Math.ceil(filteredData1.length / itemsPerPage);
    
    document.getElementById('dataCount1').textContent = `データ件数: ${filteredData1.length}`;
    document.getElementById('pageInfo1').textContent = `${currentPage1} / ${totalPages}`;
    
    document.getElementById('prevPage1').disabled = currentPage1 <= 1;
    document.getElementById('nextPage1').disabled = currentPage1 >= totalPages;
}

// ページ変更
function changePage(tabId, direction) {
    if (tabId === 1) {
        const totalPages = Math.ceil(filteredData1.length / itemsPerPage);
        currentPage1 = Math.max(1, Math.min(totalPages, currentPage1 + direction));
        renderDataTable1();
    } else if (tabId === 2) {
        const totalPages = Math.ceil(filteredData2.length / itemsPerPage);
        currentPage2 = Math.max(1, Math.min(totalPages, currentPage2 + direction));
        renderTaskDataTable2();
    }
}

// タスク管理の結果表示
function displayTaskManagementResults() {
    console.log('タスク管理結果表示を開始');
    
    updateTaskManagementSummary();
    displayTaskClassificationChart();
    displayTaskDataTable();
    displayTaskStats();
    
    const resultSection = document.getElementById('resultSection2');
    if (resultSection) {
        resultSection.style.display = 'block';
    }
}

// タスク分類結果をチャート表示
function displayTaskClassificationChart() {
    const summaryContainer = document.getElementById('classificationSummary2');
    const chartContainer = document.getElementById('classificationChart2');
    
    if (!summaryContainer || !chartContainer || !taskClassificationData) return;
    
    // サマリーカード表示
    const summaryHtml = Object.entries(taskClassificationData).map(([status, count]) => `
        <div class="classification-item">
            <span class="label">${status}</span>
            <span class="value">${count}</span>
        </div>
    `).join('');
    summaryContainer.innerHTML = summaryHtml;
    
    // バーチャート表示
    const totalTasks = Object.values(taskClassificationData).reduce((sum, count) => sum + count, 0);
    const chartHtml = Object.entries(taskClassificationData).map(([status, count]) => {
        const percentage = totalTasks > 0 ? (count / totalTasks * 100) : 0;
        return `
            <div class="chart-bar">
                <div class="chart-bar-label">${status}</div>
                <div class="chart-bar-visual">
                    <div class="chart-bar-fill" style="width: ${percentage}%"></div>
                </div>
                <div class="chart-bar-value">${count}件</div>
            </div>
        `;
    }).join('');
    chartContainer.innerHTML = chartHtml;
}

// タスクデータテーブルの表示
function displayTaskDataTable() {
    if (!taskProcessedData.length) return;
    
    // フィルター設定
    setupTaskFilters2();
    
    // 初期データ表示
    filteredData2 = [...taskProcessedData];
    currentPage2 = 1;
    renderTaskDataTable2();
    
    // フィルターイベントリスナー
    const taskFilter = document.getElementById('taskFilter2');
    const statusFilter = document.getElementById('taskStatusFilter2');
    const categoryFilter = document.getElementById('taskCategoryFilter2');
    
    if (taskFilter) taskFilter.addEventListener('input', applyTaskFilters2);
    if (statusFilter) statusFilter.addEventListener('change', applyTaskFilters2);
    if (categoryFilter) categoryFilter.addEventListener('change', applyTaskFilters2);
}

// タスクフィルター設定
function setupTaskFilters2() {
    const statusFilter = document.getElementById('taskStatusFilter2');
    const categoryFilter = document.getElementById('taskCategoryFilter2');
    
    if (statusFilter) {
        // 進捗状況フィルター（分類後の進捗状況を使用）
        const statuses = [...new Set(taskProcessedData.map(row => row['進捗状況']).filter(Boolean))];
        statusFilter.innerHTML = '<option value="">すべての進捗状況</option>';
        statuses.forEach(status => {
            statusFilter.innerHTML += `<option value="${status}">${status}</option>`;
        });
    }
    
    if (categoryFilter) {
        // 部署フィルター
        const departments = [...new Set(taskProcessedData.map(row => row['依頼部署']).filter(Boolean))];
        categoryFilter.innerHTML = '<option value="">すべての部署</option>';
        departments.forEach(dept => {
            categoryFilter.innerHTML += `<option value="${dept}">${dept}</option>`;
        });
        // カテゴリフィルターのラベルも変更
        const categoryLabel = document.querySelector('label[for="taskCategoryFilter2"]');
        if (categoryLabel) {
            categoryLabel.textContent = '部署フィルター:';
        }
    }
}

// タスクフィルター適用
function applyTaskFilters2() {
    const taskFilter = document.getElementById('taskFilter2');
    const statusFilter = document.getElementById('taskStatusFilter2');
    const categoryFilter = document.getElementById('taskCategoryFilter2');
    
    const searchText = taskFilter ? taskFilter.value.toLowerCase() : '';
    const selectedStatus = statusFilter ? statusFilter.value : '';
    const selectedDepartment = categoryFilter ? categoryFilter.value : '';
    
    filteredData2 = taskProcessedData.filter(row => {
        const searchMatch = !searchText || Object.values(row).some(value => 
            String(value).toLowerCase().includes(searchText)
        );
        const statusMatch = !selectedStatus || row['進捗状況'] === selectedStatus;
        const departmentMatch = !selectedDepartment || row['依頼部署'] === selectedDepartment;
        
        return searchMatch && statusMatch && departmentMatch;
    });
    
    currentPage2 = 1;
    renderTaskDataTable2();
}

// タスクテーブル描画
function renderTaskDataTable2() {
    const table = document.getElementById('taskDataTable2');
    if (!table) return;
    
    const startIndex = (currentPage2 - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const pageData = filteredData2.slice(startIndex, endIndex);
    
    if (pageData.length > 0) {
        const headers = Object.keys(pageData[0]);
        table.querySelector('thead').innerHTML = `
            <tr>${headers.map(header => `<th>${header}</th>`).join('')}</tr>
        `;
        
        const tbody = table.querySelector('tbody');
        tbody.innerHTML = pageData.map(row => `
            <tr>${headers.map(header => {
                let cellValue = row[header] || '';
                
                // 金額項目の表示形式調整
                if (header === 'インターン費用' || header === '外注時費用' || header === '収益貢献（年）') {
                    const numValue = parseFloat(cellValue);
                    if (numValue > 0) {
                        cellValue = `${numValue.toLocaleString()}円`;
                    } else {
                        cellValue = '0円';
                    }
                }
                // 進捗状況にバッジを適用
                else if (header === '進捗状況') {
                    cellValue = `<span class="status-badge status-${cellValue.replace(/\s+/g, '-')}">${cellValue}</span>`;
                }
                return `<td>${cellValue}</td>`;
            }).join('')}</tr>
        `).join('');
    }
    
    updateTaskPagination2();
}

// タスクページネーション更新
function updateTaskPagination2() {
    const totalPages = Math.ceil(filteredData2.length / itemsPerPage);
    
    const taskCount = document.getElementById('taskCount2');
    const pageInfo = document.getElementById('taskPageInfo2');
    const prevBtn = document.getElementById('prevTaskPage2');
    const nextBtn = document.getElementById('nextTaskPage2');
    
    if (taskCount) taskCount.textContent = `タスク件数: ${filteredData2.length}`;
    if (pageInfo) pageInfo.textContent = `${currentPage2} / ${totalPages}`;
    if (prevBtn) prevBtn.disabled = currentPage2 <= 1;
    if (nextBtn) nextBtn.disabled = currentPage2 >= totalPages;
}

// タスク統計情報の表示
function displayTaskStats() {
    const container = document.getElementById('taskStats2');
    if (!container || !taskClassificationData) return;
    
    const totalProcessed = taskProcessedData.length;
    const totalRaw = taskRawData.length;
    const excludedCount = totalRaw - totalProcessed;
    
    // 新しい統計情報を計算
    const statistics = calculateTaskStatistics(taskProcessedData);
    
    const statsData = [
        { label: '総タスク数', value: totalRaw, class: 'stat-total' },
        { label: '処理済み', value: totalProcessed, class: 'stat-processed' },
        { label: '除外タスク', value: excludedCount, class: 'stat-excluded' },
        { label: '処理率', value: `${((totalProcessed / totalRaw) * 100).toFixed(1)}%`, class: 'stat-rate' },
        { label: 'インターン費用', value: `${(statistics.totalInternCost / 10000).toFixed(0)}万円`, class: 'stat-intern-cost' },
        { label: '外注費用', value: `${(statistics.totalOutsourcingCost / 10000).toFixed(0)}万円`, class: 'stat-outsourcing-cost' },
        { label: '年間収益貢献', value: `${(statistics.totalAnnualRevenue / 10000).toFixed(0)}万円`, class: 'stat-revenue' },
        { label: '費用対効果', value: `${statistics.costEffectiveness.toFixed(1)}倍`, class: 'stat-efficiency' }
    ];
    
    container.innerHTML = statsData.map(stat => `
        <div class="stat-card ${stat.class}">
            <div class="stat-value">${stat.value}</div>
            <div class="stat-label">${stat.label}</div>
        </div>
    `).join('');
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
    console.log('=== タスク管理データ処理開始 ===');
    console.log(`元データ件数: ${taskRawData.length} 件`);
    
    if (taskRawData.length === 0) {
        console.error('データが空です');
        return;
    }
    
    const firstRow = taskRawData[0];
    console.log(`列数: ${Object.keys(firstRow).length} 列`);
    console.log('利用可能な列名:', Object.keys(firstRow));
    
    // 進捗状況の分類定義（dev_list.jsと同じ）
    const progressMapping = {
        'リリース済み': ['公開済み（アップデート予定）', '完了'],
        'リリース準備': ['公開可能（川島さん確認待ち）', 'FB待ち'],
        '開発中': ['推奨（宮川案）', '設計', '実装']
    };
    
    // 削除対象の進捗状況
    const deleteStatus = ['未着手', '未確認', '検討', '保留', '中断'];
    
    // 削除対象の行を除外
    const filteredResults = taskRawData.filter(row => 
        !deleteStatus.includes(row['進捗状況'])
    );
    
    console.log(`\n=== フィルタリング後 ===`);
    console.log(`処理対象件数: ${filteredResults.length} 件（削除: ${taskRawData.length - filteredResults.length} 件）`);
    
    // 進捗状況を分類する関数
    function classifyProgress(status) {
        for (const [category, values] of Object.entries(progressMapping)) {
            // 部分一致で判定（カンマ区切りの場合も対応）
            for (const value of values) {
                if (String(status).includes(value)) {
                    return category;
                }
            }
        }
        return '開発中'; // デフォルト
    }
    
    // 進捗状況を分類して新しい列を作成
    filteredResults.forEach(row => {
        row['進捗状況_分類'] = classifyProgress(row['進捗状況']);
    });
    
    // 分類結果の確認
    console.log("\n=== 進捗状況分類結果 ===");
    taskClassificationData = {};
    filteredResults.forEach(row => {
        const status = row['進捗状況_分類'];
        taskClassificationData[status] = (taskClassificationData[status] || 0) + 1;
    });
    console.log(taskClassificationData);
    
    // 項番の列名を探す
    const availableColumns = Object.keys(firstRow);
    const noColumn = availableColumns.find(col => 
        col.includes('No') || col.includes('項番') || col.includes('番号')
    );
    console.log("項番の列名:", noColumn || '見つかりません');
    
    // 列名マッピング（dev_list.jsと同じ）
    const columnMapping = {
        '項番': noColumn || 'No',
        '依頼部署': '部署',
        '開発分類': '開発分類',
        'タスク内容': 'タスク内容',
        '進捗状況': '進捗状況_分類', // 分類後の進捗状況を使用
        '外注時費用': 'アウトソーシング費用'
    };
    
    // 新しいデータ配列を作成
    const newData = [];
    
    // 対応する列を抽出
    filteredResults.forEach(row => {
        const newRow = {};
        for (const [newCol, originalCol] of Object.entries(columnMapping)) {
            if (originalCol && row.hasOwnProperty(originalCol)) {
                newRow[newCol] = row[originalCol];
            } else {
                // 対応する列がない場合は空の列を作成
                newRow[newCol] = '';
            }
        }
        newData.push(newRow);
    });
    
    // インターン費用を開発工数(h)×2000で算出
    let internCosts = [];
    if (firstRow.hasOwnProperty('開発工数(h)')) {
        internCosts = filteredResults.map(row => {
            const devHours = parseFloat(row['開発工数(h)']) || 0;
            return devHours * 2000;
        });
        console.log('インターン費用を算出しました（開発工数×2000）');
    } else {
        internCosts = new Array(filteredResults.length).fill(0);
        console.log('開発工数(h)列が見つからないため、インターン費用は0に設定');
    }
    
    // 外注時費用を数値に変換
    let outsourcingCosts = [];
    if (newData[0] && newData[0].hasOwnProperty('外注時費用')) {
        outsourcingCosts = newData.map(row => {
            return parseFloat(row['外注時費用']) || 0;
        });
        console.log('外注時費用を数値変換しました');
    } else {
        outsourcingCosts = new Array(filteredResults.length).fill(0);
        console.log('外注時費用列が見つからないため、0に設定');
    }
    
    // 収益貢献（年）を(月工数×500000/150)×12で算出
    let annualRevenues = [];
    if (firstRow.hasOwnProperty('月工数（h）')) {
        annualRevenues = filteredResults.map(row => {
            const monthlyHours = parseFloat(row['月工数（h）']) || 0;
            // 月工数×500000/150×12ヶ月で年間収益貢献を算出
            return Math.round(monthlyHours * 500000 / 150) * 12;
        });
        console.log('収益貢献（年）を算出しました（月工数×500000/150×12）');
    } else {
        annualRevenues = new Array(filteredResults.length).fill(0);
        console.log('月工数（h）列が見つからないため、収益貢献（年）は0に設定');
    }
    
    // インターン費用と収益貢献（年）を追加
    newData.forEach((row, index) => {
        row['インターン費用'] = internCosts[index];
        row['収益貢献（年）'] = annualRevenues[index];
    });
    
    // 項番で昇順ソート
    if (newData[0] && newData[0].hasOwnProperty('項番') && newData[0]['項番']) {
        newData.sort((a, b) => {
            const aNum = parseInt(String(a['項番']).match(/\d+/)?.[0] || '0');
            const bNum = parseInt(String(b['項番']).match(/\d+/)?.[0] || '0');
            return aNum - bNum;
        });
        console.log('項番でソートしました');
    }
    
    // 列の順序を指定
    const columnOrder = ['項番', '依頼部署', '開発分類', 'タスク内容', '進捗状況', 'インターン費用', '外注時費用', '収益貢献（年）'];
    
    // 順序に従ってデータを整理
    const orderedData = newData.map(row => {
        const orderedRow = {};
        columnOrder.forEach(col => {
            orderedRow[col] = row[col] || '';
        });
        return orderedRow;
    });
    
    taskProcessedData = orderedData;
    
    console.log(`\n=== タスク管理データ処理完了 ===`);
    console.log(`最終データ件数: ${taskProcessedData.length} 件`);
    console.log(`出力列数: ${columnOrder.length} 列`);
    console.log('出力列順:', columnOrder);
    
    // 統計情報の計算
    const statistics = calculateTaskStatistics(taskProcessedData);
    console.log('\n=== 統計情報 ===');
    console.log(`総インターン費用: ${statistics.totalInternCost.toLocaleString()} 円`);
    console.log(`総外注費用: ${statistics.totalOutsourcingCost.toLocaleString()} 円`);
    console.log(`総収益貢献: ${statistics.totalAnnualRevenue.toLocaleString()} 円`);
    console.log(`費用対効果: ${statistics.costEffectiveness.toFixed(2)}`);
}

// タスク統計情報を計算
function calculateTaskStatistics(data) {
    const totalInternCost = data.reduce((sum, row) => sum + (parseFloat(row['インターン費用']) || 0), 0);
    const totalOutsourcingCost = data.reduce((sum, row) => sum + (parseFloat(row['外注時費用']) || 0), 0);
    const totalAnnualRevenue = data.reduce((sum, row) => sum + (parseFloat(row['収益貢献（年）']) || 0), 0);
    const totalCost = totalInternCost + totalOutsourcingCost;
    const costEffectiveness = totalCost > 0 ? totalAnnualRevenue / totalCost : 0;
    
    return {
        totalInternCost,
        totalOutsourcingCost,
        totalAnnualRevenue,
        totalCost,
        costEffectiveness
    };
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
        
        console.log('タスクCSVダウンロード完了:', filename);
        console.log('出力データ件数:', taskProcessedData.length);
        console.log('出力列:', Object.keys(taskProcessedData[0] || {}));
        
    } catch (error) {
        console.error('タスクCSVダウンロードエラー:', error);
        showError('CSVファイルの作成中にエラーが発生しました: ' + error.message);
    }
}

// タスク管理のPowerPoint出力（ページネーション対応）
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
        
        // フィルタリング済みデータを使用（現在ブラウザに表示されているデータ）
        // filteredData2が空の場合はtaskProcessedDataを使用
        const dataToExport = (filteredData2 && filteredData2.length > 0) ? filteredData2 : taskProcessedData;
        const totalPages = Math.ceil(dataToExport.length / itemsPerPage);
        
        console.log(`PowerPoint生成開始: ${dataToExport.length}件のデータを${totalPages}ページに分割`);
        
        // サマリースライドを作成
        createTaskSummarySlideForPPT(pptx, dataToExport);
        
        // データを20件ずつに分けてスライド作成
        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            const startIndex = (pageNum - 1) * itemsPerPage;
            const endIndex = Math.min(startIndex + itemsPerPage, dataToExport.length);
            const pageData = dataToExport.slice(startIndex, endIndex);
            
            console.log(`スライド${pageNum + 1}を作成中: ${startIndex + 1}〜${endIndex}件目（${pageData.length}件）`);
            
            createTaskDataSlide(pptx, pageData, pageNum, totalPages, startIndex + 1, endIndex);
        }
        
        const filename = `タスク管理レポート_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.pptx`;
        pptx.writeFile({ fileName: filename }).then(() => {
            console.log(`PowerPoint出力完了: ${filename}`);
            console.log(`総スライド数: ${totalPages + 1}枚（サマリー1枚 + データ${totalPages}枚）`);
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

// タスク管理サマリースライド作成
function createTaskSummarySlideForPPT(pptx, data) {
    const slide = pptx.addSlide();
    
    // スライドタイトル
    slide.addText('タスク管理レポート - サマリー', {
        x: 0.5, y: 0.3, w: 9, h: 0.7,
        fontSize: 24, bold: true, align: 'center',
        color: '2F4F4F'
    });
    
    // 作成日時
    slide.addText(`作成日時: ${new Date().toLocaleString('ja-JP')}`, {
        x: 0.5, y: 1.0, w: 9, h: 0.4,
        fontSize: 12, align: 'center',
        color: '666666'
    });
    
    // 統計情報
    const statistics = calculateTaskStatistics(data);
    const totalRaw = taskRawData.length;
    const totalProcessed = data.length;
    const excludedCount = totalRaw - totalProcessed;
    
    const summaryText = [
        { text: '処理統計\n', options: { fontSize: 18, bold: true, color: '2F4F4F' } },
        { text: `・総タスク数: ${totalRaw}件\n`, options: { fontSize: 14 } },
        { text: `・処理済みタスク数: ${totalProcessed}件\n`, options: { fontSize: 14 } },
        { text: `・除外タスク数: ${excludedCount}件\n`, options: { fontSize: 14 } },
        { text: `・処理率: ${((totalProcessed / totalRaw) * 100).toFixed(1)}%\n\n`, options: { fontSize: 14 } },
        
        { text: '費用・収益統計\n', options: { fontSize: 18, bold: true, color: '2F4F4F' } },
        { text: `・総インターン費用: ${(statistics.totalInternCost / 10000).toFixed(0)}万円\n`, options: { fontSize: 14 } },
        { text: `・総外注費用: ${(statistics.totalOutsourcingCost / 10000).toFixed(0)}万円\n`, options: { fontSize: 14 } },
        { text: `・年間収益貢献: ${(statistics.totalAnnualRevenue / 10000).toFixed(0)}万円\n`, options: { fontSize: 14 } },
        { text: `・費用対効果: ${statistics.costEffectiveness.toFixed(1)}倍`, options: { fontSize: 14 } }
    ];
    
    slide.addText(summaryText, {
        x: 1, y: 1.8, w: 8, h: 4,
        align: 'left'
    });
    
    // 進捗分類統計
    const classificationText = [
        { text: '進捗状況分類\n', options: { fontSize: 18, bold: true, color: '2F4F4F' } }
    ];
    
    Object.entries(taskClassificationData).forEach(([status, count]) => {
        const percentage = totalProcessed > 0 ? ((count / totalProcessed) * 100).toFixed(1) : 0;
        classificationText.push({
            text: `・${status}: ${count}件 (${percentage}%)\n`,
            options: { fontSize: 14 }
        });
    });
    
    slide.addText(classificationText, {
        x: 1, y: 5.5, w: 8, h: 2,
        align: 'left'
    });
}

// タスクデータスライド作成
function createTaskDataSlide(pptx, pageData, pageNum, totalPages, startNum, endNum) {
    const slide = pptx.addSlide();
    
    // スライドタイトル
    slide.addText(`タスク一覧 - ${pageNum}/${totalPages}ページ (${startNum}〜${endNum}件目)`, {
        x: 0.3, y: 0.2, w: 9.4, h: 0.5,
        fontSize: 16, bold: true, align: 'center',
        color: '2F4F4F'
    });
    
    // テーブルデータの準備
    const headers = Object.keys(pageData[0] || {});
    const tableData = [headers]; // ヘッダー行
    
    // データ行を追加
    pageData.forEach(row => {
        const rowData = headers.map(header => {
            let value = row[header] || '';
            
            // 金額項目の表示形式調整
            if (header === 'インターン費用' || header === '外注時費用' || header === '収益貢献（年）') {
                const numValue = parseFloat(value);
                if (numValue > 0) {
                    value = numValue.toLocaleString(); // カンマ区切りで表示
                } else {
                    value = '0';
                }
            }
            
            // 長いテキストの切り詰め
            if (typeof value === 'string' && value.length > 30) {
                value = value.substring(0, 27) + '...';
            }
            
            return value.toString();
        });
        tableData.push(rowData);
    });
    
    // 列幅の動的計算
    const totalWidth = 9.4;
    const colCount = headers.length;
    const baseWidth = totalWidth / colCount;
    
    // 列ごとの幅調整
    const colWidths = headers.map(header => {
        switch (header) {
            case '項番': return baseWidth * 0.6;
            case '依頼部署': return baseWidth * 1.0;
            case '開発分類': return baseWidth * 1.2;
            case 'タスク内容': return baseWidth * 2.0;
            case '進捗状況': return baseWidth * 1.0;
            case 'インターン費用': return baseWidth * 1.0;
            case '外注時費用': return baseWidth * 1.0;
            case '収益貢献（年）': return baseWidth * 1.2;
            default: return baseWidth;
        }
    });
    
    // テーブルをスライドに追加
    slide.addTable(tableData, {
        x: 0.3, y: 0.8, w: totalWidth, h: 6.5,
        fontSize: 10,
        colW: colWidths,
        border: { type: 'solid', color: 'CCCCCC', pt: 1 },
        fill: { color: 'FFFFFF' },
        headerRow: { 
            fill: { color: 'F2F2F2' }, 
            bold: true,
            fontSize: 11,
            color: '2F4F4F'
        },
        align: 'center',
        valign: 'middle'
    });
    
    // ページ情報をフッターに追加
    const totalDisplayed = (filteredData2 && filteredData2.length > 0) ? filteredData2.length : taskProcessedData.length;
    slide.addText(`${pageData.length}件表示 | 全${totalDisplayed}件中`, {
        x: 0.3, y: 7.4, w: 9.4, h: 0.3,
        fontSize: 10, align: 'right',
        color: '999999'
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