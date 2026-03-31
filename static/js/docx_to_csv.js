/**
 * DOCX to CSV Extraction Logic
 */
let docxCsvSelectedFile = null;
let docxCsvTaskId = null;
let isDocxCsvExtracting = false;

const docxCsvUploadArea = document.getElementById('docxCsvUploadArea');
const docxCsvFileInput = document.getElementById('docxCsvFileInput');
const docxCsvExtractBtn = document.getElementById('docxCsvExtractBtn');
const docxCsvFileInfo = document.getElementById('docxCsvFileInfo');
const docxCsvFileNameEl = document.getElementById('docxCsvFileNameDisplay');
const docxCsvFileSizeEl = document.getElementById('docxCsvFileSizeDisplay');
const docxCsvRemoveFileBtn = document.getElementById('docxCsvRemoveFile');
const docxCsvOptionsContainer = document.getElementById('docxCsvOptionsContainer');
const docxCsvProgressContainer = document.getElementById('docxCsvProgressContainer');
const docxCsvProgressBarEl = document.getElementById('docxCsvProgressBar');
const docxCsvProgressPercentage = document.getElementById('docxCsvProgressPercentage');
const docxCsvStatusMessage = document.getElementById('docxCsvStatusMessage');
const docxCsvStatusText = document.getElementById('docxCsvStatusText');
const docxCsvDownloadContainer = document.getElementById('docxCsvDownloadContainer');
const docxCsvDownloadBtn = document.getElementById('docxCsvDownloadBtn');
const docxCsvDownloadTitle = document.getElementById('docxCsvDownloadTitle');
const docxCsvTableCount = document.getElementById('docxCsvTableCount');

// Upload area click
docxCsvUploadArea.addEventListener('click', (e) => {
    if (e.target.tagName !== 'INPUT') {
        docxCsvFileInput.click();
    }
});

// File input change
docxCsvFileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleDocxCsvFileSelect(e.target.files[0]);
});

// Drag and drop
docxCsvUploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    docxCsvUploadArea.classList.add('dragover');
});
docxCsvUploadArea.addEventListener('dragleave', () => {
    docxCsvUploadArea.classList.remove('dragover');
});
docxCsvUploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    docxCsvUploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.name.toLowerCase().endsWith('.docx')) handleDocxCsvFileSelect(file);
    else alert('请选择DOCX文件');
});

// Remove file
docxCsvRemoveFileBtn.addEventListener('click', resetDocxCsvUpload);

// Extract button
docxCsvExtractBtn.addEventListener('click', startDocxCsvExtraction);

// Download button
docxCsvDownloadBtn.addEventListener('click', () => {
    if (docxCsvTaskId) {
        window.location.href = `/api/docx-csv/download/${docxCsvTaskId}`;
        setTimeout(() => cleanupDocxCsvFiles(), 1000);
    }
});

function handleDocxCsvFileSelect(file) {
    if (!file.name.toLowerCase().endsWith('.docx')) { alert('请选择DOCX文件'); return; }
    if (file.size > 80 * 1024 * 1024) { alert('文件大小不能超过80MB'); return; }
    docxCsvSelectedFile = file;
    docxCsvFileNameEl.textContent = file.name;
    docxCsvFileSizeEl.textContent = `(${formatFileSize(file.size)})`;
    docxCsvUploadArea.classList.add('hidden');
    docxCsvFileInfo.classList.remove('hidden');
    docxCsvOptionsContainer.classList.remove('hidden');
    docxCsvExtractBtn.disabled = false;
    resetDocxCsvProgress();
}

function resetDocxCsvUpload() {
    docxCsvSelectedFile = null;
    docxCsvFileInput.value = '';
    docxCsvUploadArea.classList.remove('hidden');
    docxCsvFileInfo.classList.add('hidden');
    docxCsvOptionsContainer.classList.add('hidden');
    docxCsvExtractBtn.disabled = true;
    isDocxCsvExtracting = false;
    resetDocxCsvProgress();
}

function resetDocxCsvProgress() {
    docxCsvProgressContainer.classList.add('hidden');
    docxCsvDownloadContainer.classList.add('hidden');
    docxCsvStatusMessage.className = 'status-message info';
}

async function startDocxCsvExtraction() {
    if (!docxCsvSelectedFile || isDocxCsvExtracting) return;
    isDocxCsvExtracting = true;

    const formData = new FormData();
    formData.append('file', docxCsvSelectedFile);

    const outputMode = document.querySelector('input[name="docxCsvOutputMode"]:checked').value;
    formData.append('output_mode', outputMode);

    // Hide download container and reset task ID from previous extraction
    docxCsvDownloadContainer.classList.add('hidden');
    docxCsvTaskId = null;

    docxCsvExtractBtn.disabled = true;
    docxCsvExtractBtn.innerHTML = '<span class="loading-spinner"></span>提取中...';
    docxCsvProgressContainer.classList.remove('hidden');
    docxCsvStatusText.textContent = '📤 正在上传文件...';
    docxCsvProgressBarEl.style.width = '5%';

    try {
        const response = await fetch('/api/docx-csv/extract', { method: 'POST', body: formData });
        const result = await response.json();
        if (response.ok) {
            docxCsvTaskId = result.task_id;
            docxCsvStatusText.textContent = '✅ 文件上传成功，开始提取...';
            docxCsvProgressBarEl.style.width = '10%';
            pollDocxCsvStatus();
        } else {
            throw new Error(result.error || '上传失败');
        }
    } catch (error) {
        docxCsvStatusText.textContent = `❌ ${error.message}`;
        docxCsvStatusMessage.className = 'status-message error';
        docxCsvExtractBtn.disabled = false;
        docxCsvExtractBtn.innerHTML = '📋 开始提取';
        isDocxCsvExtracting = false;
    }
}

async function pollDocxCsvStatus() {
    if (!docxCsvTaskId) return;
    try {
        const response = await fetch(`/api/docx-csv/status/${docxCsvTaskId}`);
        const data = await response.json();
        if (response.ok) {
            docxCsvProgressBarEl.style.width = `${data.progress}%`;
            docxCsvProgressPercentage.textContent = `${data.progress}%`;
            docxCsvStatusText.textContent = data.message;

            if (data.status === 'processing') {
                setTimeout(pollDocxCsvStatus, 1000);
            } else if (data.status === 'completed') {
                docxCsvStatusMessage.className = 'status-message success';
                if (data.no_tables) {
                    docxCsvDownloadTitle.textContent = '⚠️ 未发现表格';
                    docxCsvTableCount.textContent = '该DOCX文件中未检测到表格数据';
                    docxCsvDownloadContainer.classList.remove('hidden');
                    docxCsvDownloadBtn.classList.add('hidden');
                } else {
                    docxCsvDownloadTitle.textContent = '✅ 表格提取完成！';
                    docxCsvTableCount.textContent = `共提取 ${data.tables_found} 个表格`;
                    docxCsvDownloadContainer.classList.remove('hidden');
                    docxCsvDownloadBtn.classList.remove('hidden');
                }
                docxCsvExtractBtn.innerHTML = '📋 重新提取';
                docxCsvExtractBtn.disabled = false;
                isDocxCsvExtracting = false;
            } else if (data.status === 'error') {
                docxCsvStatusMessage.className = 'status-message error';
                docxCsvExtractBtn.disabled = false;
                docxCsvExtractBtn.innerHTML = '📋 重新提取';
                isDocxCsvExtracting = false;
            }
        }
    } catch (error) {
        docxCsvStatusText.textContent = '获取状态失败';
        docxCsvStatusMessage.className = 'status-message error';
        docxCsvExtractBtn.disabled = false;
        docxCsvExtractBtn.innerHTML = '📋 重新提取';
        isDocxCsvExtracting = false;
    }
}

async function cleanupDocxCsvFiles() {
    if (docxCsvTaskId) {
        try { await fetch(`/api/docx-csv/cleanup/${docxCsvTaskId}`, { method: 'DELETE' }); }
        catch (e) { console.error('DOCX-CSV cleanup failed:', e); }
    }
}
