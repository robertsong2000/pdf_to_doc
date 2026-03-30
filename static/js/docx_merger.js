/**
 * DOCX Merger Logic
 */
let mergeSelectedFiles = [];

const mergeUploadArea = document.getElementById('mergeUploadArea');
const mergeFileInput = document.getElementById('mergeFileInput');
const mergeFileList = document.getElementById('mergeFileList');
const mergeBtn = document.getElementById('mergeBtn');
const mergeResult = document.getElementById('mergeResult');

// Click upload area
mergeUploadArea.addEventListener('click', () => {
    mergeFileInput.click();
});

// File input change
mergeFileInput.addEventListener('change', (e) => {
    handleMergeFiles(e.target.files);
});

// Drag and drop
mergeUploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    mergeUploadArea.classList.add('dragover');
});

mergeUploadArea.addEventListener('dragleave', () => {
    mergeUploadArea.classList.remove('dragover');
});

mergeUploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    mergeUploadArea.classList.remove('dragover');
    handleMergeFiles(e.dataTransfer.files);
});

function handleMergeFiles(files) {
    const MAX_TOTAL_SIZE = 200 * 1024 * 1024; // 200MB
    let currentTotalSize = mergeSelectedFiles.reduce((sum, f) => sum + f.size, 0);

    for (const file of files) {
        if (file.name.toLowerCase().endsWith('.docx')) {
            if (!mergeSelectedFiles.find(f => f.name === file.name)) {
                // Check if adding this file would exceed the limit
                if (currentTotalSize + file.size > MAX_TOTAL_SIZE) {
                    alert(`文件总大小不能超过200MB。当前已选择${formatFileSize(currentTotalSize)}，添加${file.name}(${formatFileSize(file.size)})会超出限制`);
                    continue;
                }
                mergeSelectedFiles.push(file);
                currentTotalSize += file.size;
            }
        }
    }
    updateMergeFileList();
}

function updateMergeFileList() {
    const MAX_TOTAL_SIZE = 200 * 1024 * 1024; // 200MB

    if (mergeSelectedFiles.length === 0) {
        mergeFileList.innerHTML = '<div class="empty-state">尚未选择文件</div>';
        mergeBtn.disabled = true;
        return;
    }

    const totalSize = mergeSelectedFiles.reduce((sum, f) => sum + f.size, 0);
    const sizePercentage = Math.round((totalSize / MAX_TOTAL_SIZE) * 100);
    const sizeWarning = totalSize > MAX_TOTAL_SIZE * 0.8 ? 'text-warning' : '';

    mergeBtn.disabled = mergeSelectedFiles.length < 2 || totalSize > MAX_TOTAL_SIZE;
    mergeFileList.innerHTML = mergeSelectedFiles.map((file, index) => `
        <div class="file-item">
            <div class="file-info">
                <span class="file-icon">📄</span>
                <span class="file-name">${file.name}</span>
                <span class="file-size">(${formatFileSize(file.size)})</span>
            </div>
            <button class="remove-btn" onclick="removeMergeFile(${index})">&times;</button>
        </div>
    `).join('');

    // Add total size info
    mergeFileList.innerHTML += `
        <div class="total-size-info ${sizeWarning}">
            <span>总大小: ${formatFileSize(totalSize)} / 200MB (${sizePercentage}%)</span>
        </div>
    `;
}

window.removeMergeFile = function(index) {
    mergeSelectedFiles.splice(index, 1);
    updateMergeFileList();
};

// Merge button click
mergeBtn.addEventListener('click', async () => {
    if (mergeSelectedFiles.length < 2) {
        showMergeResult('请至少选择2个文件进行合并', 'error');
        return;
    }

    const formData = new FormData();
    mergeSelectedFiles.forEach(file => formData.append('files', file));
    formData.append('page_break', document.getElementById('pageBreak').checked);

    const outputName = document.getElementById('outputName').value.trim();
    if (outputName) {
        formData.append('output_name', outputName);
    }

    // Show loading state
    mergeBtn.disabled = true;
    mergeBtn.innerHTML = '<span class="loading-spinner"></span>合并中...';

    try {
        const response = await fetch('/api/merge', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (response.status === 413) {
            showMergeResult(data.error || '文件总大小超过限制', 'error');
        } else if (data.success) {
            showMergeResult(`
                <h4>✅ ${data.message}</h4>
                <button class="download-btn" onclick="window.location.href='${data.download_url}'">
                    ⬇️ 下载合并后的文件
                </button>
            `, 'success');
        } else {
            showMergeResult(data.error || '合并失败', 'error');
        }
    } catch (error) {
        showMergeResult('网络错误: ' + error.message, 'error');
    } finally {
        mergeBtn.disabled = false;
        mergeBtn.innerHTML = '🔄 开始合并';
    }
});

function showMergeResult(message, type) {
    mergeResult.style.display = 'block';
    mergeResult.className = type === 'success' ? 'result-success' : 'result-error';
    mergeResult.innerHTML = message;
}
