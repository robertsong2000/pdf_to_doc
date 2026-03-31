/**
 * PDF to CSV Extraction Logic
 */
let csvSelectedFile = null;
let csvTaskId = null;
let isExtracting = false;

const csvUploadArea = document.getElementById('csvUploadArea');
const csvFileInput = document.getElementById('csvFileInput');
const csvExtractBtn = document.getElementById('csvExtractBtn');
const csvFileInfo = document.getElementById('csvFileInfo');
const csvFileName = document.getElementById('csvFileName');
const csvFileSize = document.getElementById('csvFileSize');
const csvRemoveFileBtn = document.getElementById('csvRemoveFile');
const csvPageRangeContainer = document.getElementById('csvPageRangeContainer');
const csvProgressContainer = document.getElementById('csvProgressContainer');
const csvProgressBar = document.getElementById('csvProgressBar');
const csvProgressPercentage = document.getElementById('csvProgressPercentage');
const csvStatusMessage = document.getElementById('csvStatusMessage');
const csvStatusText = document.getElementById('csvStatusText');
const csvDownloadContainer = document.getElementById('csvDownloadContainer');
const csvDownloadBtn = document.getElementById('csvDownloadBtn');
const csvDownloadTitle = document.getElementById('csvDownloadTitle');
const csvTableCount = document.getElementById('csvTableCount');

// Page mode toggle for CSV tab
document.querySelectorAll('input[name="csvPageMode"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        document.getElementById('csvPageRangeInputs').style.display =
            e.target.value === 'range' ? 'flex' : 'none';
    });
});

// Upload area interactions
csvUploadArea.addEventListener('click', () => csvFileInput.click());
csvFileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleCsvFileSelect(e.target.files[0]);
});
csvUploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    csvUploadArea.classList.add('dragover');
});
csvUploadArea.addEventListener('dragleave', () => {
    csvUploadArea.classList.remove('dragover');
});
csvUploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    csvUploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') handleCsvFileSelect(file);
    else alert('请选择PDF文件');
});
csvRemoveFileBtn.addEventListener('click', resetCsvUpload);
csvExtractBtn.addEventListener('click', startCsvExtraction);
csvDownloadBtn.addEventListener('click', () => {
    if (csvTaskId) {
        window.location.href = `/api/csv/download/${csvTaskId}`;
        setTimeout(() => cleanupCsvFiles(), 1000);
    }
});

function handleCsvFileSelect(file) {
    if (file.type !== 'application/pdf') { alert('请选择PDF文件'); return; }
    if (file.size > 80 * 1024 * 1024) { alert('文件大小不能超过80MB'); return; }
    csvSelectedFile = file;
    csvFileName.textContent = file.name;
    csvFileSize.textContent = `(${formatFileSize(file.size)})`;
    csvUploadArea.classList.add('hidden');
    csvFileInfo.classList.remove('hidden');
    csvPageRangeContainer.classList.remove('hidden');
    csvExtractBtn.disabled = false;
    resetCsvProgress();
}

function resetCsvUpload() {
    csvSelectedFile = null;
    csvFileInput.value = '';
    csvUploadArea.classList.remove('hidden');
    csvFileInfo.classList.add('hidden');
    csvPageRangeContainer.classList.add('hidden');
    csvExtractBtn.disabled = true;
    isExtracting = false;
    resetCsvProgress();
}

function resetCsvProgress() {
    csvProgressContainer.classList.add('hidden');
    csvDownloadContainer.classList.add('hidden');
    csvStatusMessage.className = 'status-message info';
}

async function startCsvExtraction() {
    if (!csvSelectedFile || isExtracting) return;
    isExtracting = true;

    const formData = new FormData();
    formData.append('file', csvSelectedFile);

    // Page range
    const csvPageMode = document.querySelector('input[name="csvPageMode"]:checked').value;
    if (csvPageMode === 'range') {
        const s = parseInt(document.getElementById('csvStartPage').value);
        const e = parseInt(document.getElementById('csvEndPage').value);
        if (s > 0 && e > 0 && e >= s) {
            formData.append('start_page', s);
            formData.append('end_page', e);
        } else {
            alert('请输入有效的页码范围');
            isExtracting = false;
            return;
        }
    }

    // Output mode
    const outputMode = document.querySelector('input[name="csvOutputMode"]:checked').value;
    formData.append('output_mode', outputMode);

    csvExtractBtn.disabled = true;
    csvExtractBtn.innerHTML = '<span class="loading-spinner"></span>提取中...';
    csvProgressContainer.classList.remove('hidden');
    csvStatusText.textContent = '📤 正在上传文件...';
    csvProgressBar.style.width = '5%';

    try {
        const response = await fetch('/api/csv/extract', { method: 'POST', body: formData });
        const result = await response.json();
        if (response.ok) {
            csvTaskId = result.task_id;
            csvStatusText.textContent = '✅ 文件上传成功，开始提取...';
            csvProgressBar.style.width = '10%';
            pollCsvStatus();
        } else {
            throw new Error(result.error || '上传失败');
        }
    } catch (error) {
        csvStatusText.textContent = `❌ ${error.message}`;
        csvStatusMessage.className = 'status-message error';
        csvExtractBtn.disabled = false;
        csvExtractBtn.innerHTML = '📊 开始提取';
        isExtracting = false;
    }
}

async function pollCsvStatus() {
    if (!csvTaskId) return;
    try {
        const response = await fetch(`/api/csv/status/${csvTaskId}`);
        const data = await response.json();
        if (response.ok) {
            csvProgressBar.style.width = `${data.progress}%`;
            csvProgressPercentage.textContent = `${data.progress}%`;
            csvStatusText.textContent = data.message;

            if (data.status === 'processing') {
                setTimeout(pollCsvStatus, 1500);
            } else if (data.status === 'completed') {
                csvStatusMessage.className = 'status-message success';
                if (data.no_tables) {
                    csvDownloadTitle.textContent = '⚠️ 未发现表格';
                    csvTableCount.textContent = '该PDF文件中未检测到表格数据';
                    csvDownloadContainer.classList.remove('hidden');
                    csvDownloadBtn.classList.add('hidden');
                } else {
                    csvDownloadTitle.textContent = '✅ 表格提取完成！';
                    csvTableCount.textContent = `共提取 ${data.tables_found} 个表格`;
                    csvDownloadContainer.classList.remove('hidden');
                    csvDownloadBtn.classList.remove('hidden');
                }
                csvExtractBtn.innerHTML = '📊 重新提取';
                csvExtractBtn.disabled = false;
                isExtracting = false;
            } else if (data.status === 'error') {
                csvStatusMessage.className = 'status-message error';
                csvExtractBtn.disabled = false;
                csvExtractBtn.innerHTML = '📊 重新提取';
                isExtracting = false;
            }
        }
    } catch (error) {
        csvStatusText.textContent = '获取状态失败';
        csvStatusMessage.className = 'status-message error';
        csvExtractBtn.disabled = false;
        csvExtractBtn.innerHTML = '📊 重新提取';
        isExtracting = false;
    }
}

async function cleanupCsvFiles() {
    if (csvTaskId) {
        try { await fetch(`/api/csv/cleanup/${csvTaskId}`, { method: 'DELETE' }); }
        catch (e) { console.error('CSV cleanup failed:', e); }
    }
}
