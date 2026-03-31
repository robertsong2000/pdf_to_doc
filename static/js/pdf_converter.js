/**
 * PDF to DOCX Converter Logic
 */
let selectedFile = null;
let currentTaskId = null;
let isConverting = false;

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const convertBtn = document.getElementById('convertBtn');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeFileBtn = document.getElementById('removeFile');
const progressContainer = document.getElementById('progressContainer');
const progressBar = document.getElementById('progressBar');
const statusMessage = document.getElementById('statusMessage');
const statusText = document.getElementById('statusText');
const downloadContainer = document.getElementById('downloadContainer');
const downloadBtn = document.getElementById('downloadBtn');
const pageRangeContainer = document.getElementById('pageRangeContainer');
const pageRangeInputs = document.getElementById('pageRangeInputs');
const pageModeAll = document.getElementById('pageModeAll');
const pageModeRange = document.getElementById('pageModeRange');
const pageModeSection = document.getElementById('pageModeSection');
const startPageInput = document.getElementById('startPage');
const endPageInput = document.getElementById('endPage');
const sectionRangeInputs = document.getElementById('sectionRangeInputs');
const startSectionInput = document.getElementById('startSection');
const endSectionInput = document.getElementById('endSection');
const resolveSectionBtn = document.getElementById('resolveSectionBtn');
const resolvedRangeDisplay = document.getElementById('resolvedRangeDisplay');
const resolvedRangeText = document.getElementById('resolvedRangeText');
const resolvedRangeError = document.getElementById('resolvedRangeError');

let resolvedStartPage = null;
let resolvedEndPage = null;

// Page mode toggle
document.querySelectorAll('input[name="pageMode"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        if (e.target.value === 'range') {
            pageRangeInputs.style.display = 'flex';
            sectionRangeInputs.style.display = 'none';
        } else if (e.target.value === 'section') {
            pageRangeInputs.style.display = 'none';
            sectionRangeInputs.style.display = 'block';
        } else {
            pageRangeInputs.style.display = 'none';
            sectionRangeInputs.style.display = 'none';
        }
    });
});

// Resolve section button (on-demand: sends file + section numbers directly)
resolveSectionBtn.addEventListener('click', async () => {
    if (!selectedFile || !startSectionInput.value || !endSectionInput.value) {
        resolvedRangeError.style.display = 'block';
        resolvedRangeError.textContent = '请输入起始和结束章节号';
        resolvedRangeDisplay.style.display = 'none';
        return;
    }

    resolveSectionBtn.disabled = true;
    resolveSectionBtn.textContent = '解析中...';

    try {
        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('start_section', startSectionInput.value.trim());
        formData.append('end_section', endSectionInput.value.trim());

        const response = await fetch('/api/sections/resolve', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (!response.ok) {
            resolvedRangeError.style.display = 'block';
            resolvedRangeError.textContent = result.error || '章节解析失败';
            resolvedRangeDisplay.style.display = 'none';
            return;
        }

        resolvedStartPage = result.start_page;
        resolvedEndPage = result.end_page;

        resolvedRangeDisplay.style.display = 'block';
        resolvedRangeText.textContent = result.resolved_range;
        resolvedRangeError.style.display = 'none';

    } catch (err) {
        resolvedRangeError.style.display = 'block';
        resolvedRangeError.textContent = '解析请求失败: ' + err.message;
        resolvedRangeDisplay.style.display = 'none';
    } finally {
        resolveSectionBtn.disabled = false;
        resolveSectionBtn.textContent = '解析';
    }
});

// Upload area click
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

// File input change
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        handleFileSelect(file);
    }
});

// Drag and drop
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');

    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') {
        handleFileSelect(file);
    } else {
        alert('请选择PDF文件');
    }
});

// Remove file
removeFileBtn.addEventListener('click', () => {
    resetUpload();
});

// Convert button
convertBtn.addEventListener('click', () => {
    if (selectedFile) {
        convertFile();
    }
});

// Download button
downloadBtn.addEventListener('click', () => {
    if (currentTaskId) {
        window.location.href = `/api/download/${currentTaskId}`;
        setTimeout(() => {
            cleanupFiles();
        }, 1000);
    }
});

function handleFileSelect(file) {
    if (file.type !== 'application/pdf') {
        alert('请选择PDF文件');
        return;
    }

    if (file.size > 80 * 1024 * 1024) {
        alert('文件大小不能超过80MB');
        return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = `(${formatFileSize(file.size)})`;

    uploadArea.classList.add('hidden');
    fileInfo.classList.remove('hidden');
    pageRangeContainer.classList.remove('hidden');
    convertBtn.disabled = false;

    resetProgress();
}

function resetUpload() {
    selectedFile = null;
    fileInput.value = '';
    uploadArea.classList.remove('hidden');
    fileInfo.classList.add('hidden');
    pageRangeContainer.classList.add('hidden');
    convertBtn.disabled = true;
    isConverting = false;

    // Reset section state
    resolvedStartPage = null;
    resolvedEndPage = null;
    if (startSectionInput) startSectionInput.value = '';
    if (endSectionInput) endSectionInput.value = '';
    if (resolvedRangeDisplay) resolvedRangeDisplay.style.display = 'none';
    if (resolvedRangeError) resolvedRangeError.style.display = 'none';

    resetProgress();
}

function resetProgress() {
    progressContainer.classList.add('hidden');
    downloadContainer.classList.add('hidden');
    statusMessage.className = 'status-message info';

    // Clear detailed status elements
    const stepIndicator = document.getElementById('stepIndicator');
    const progressPercent = document.getElementById('progressPercent');
    if (stepIndicator) stepIndicator.remove();
    if (progressPercent) progressPercent.remove();
}

async function convertFile() {
    if (!selectedFile || isConverting) return;

    isConverting = true;

    const formData = new FormData();
    formData.append('file', selectedFile);

    // Append page range if specified
    if (pageModeRange.checked) {
        const startVal = parseInt(startPageInput.value);
        const endVal = parseInt(endPageInput.value);
        if (startVal > 0 && endVal > 0 && endVal >= startVal) {
            formData.append('start_page', startVal);
            formData.append('end_page', endVal);
        } else {
            alert('请输入有效的页码范围（起始页不能大于结束页）');
            isConverting = false;
            convertBtn.disabled = false;
            convertBtn.innerHTML = '✨ 开始转换';
            return;
        }
    } else if (pageModeSection.checked) {
        if (!resolvedStartPage || !resolvedEndPage) {
            alert('请先点击"解析"按钮确认章节对应的页码范围');
            isConverting = false;
            convertBtn.disabled = false;
            convertBtn.innerHTML = '✨ 开始转换';
            return;
        }
        formData.append('start_page', resolvedStartPage);
        formData.append('end_page', resolvedEndPage);
    }

    convertBtn.disabled = true;
    convertBtn.innerHTML = '<span class="loading-spinner"></span>处理中...';
    progressContainer.classList.remove('hidden');

    // 显示详细的初始状态
    updateDetailedStatus({
        message: '📤 正在上传文件到服务器...',
        progress: 5,
        step: 'uploading',
        status: 'converting'
    });

    try {
        // 显示上传进行中的状态
        updateDetailedStatus({
            message: '📡 正在传输文件数据...',
            progress: 8,
            step: 'transferring',
            status: 'converting'
        });

        const response = await fetch('/api/convert', {
            method: 'POST',
            body: formData
        });

        // 上传完成，等待服务器处理
        updateDetailedStatus({
            message: '⏳ 服务器正在准备转换...',
            progress: 10,
            step: 'preparing',
            status: 'converting'
        });

        const result = await response.json();

        if (response.ok) {
            currentTaskId = result.task_id;
            updateDetailedStatus({
                message: '✅ 文件上传成功，启动转换引擎...',
                progress: 15,
                step: 'upload_complete',
                status: 'converting'
            });
            pollConversionStatus();
        } else {
            throw new Error(result.error || '上传失败');
        }
    } catch (error) {
        updateDetailedStatus({
            message: `❌ 上传失败: ${error.message}`,
            progress: 0,
            step: 'error',
            status: 'error'
        });
        convertBtn.disabled = false;
        convertBtn.innerHTML = '✨ 开始转换';
        isConverting = false;
    }
}

async function pollConversionStatus() {
    if (!currentTaskId) return;

    try {
        // Send heartbeat to indicate frontend is still active
        try {
            await fetch(`/api/heartbeat/${currentTaskId}`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                }
            });
        } catch (heartbeatError) {
            console.warn('Failed to send heartbeat:', heartbeatError);
        }

        const response = await fetch(`/api/status/${currentTaskId}`);
        const status = await response.json();

        if (response.ok) {
            if (status.status === 'converting') {
                updateDetailedStatus(status);

                // Adjust polling frequency based on current step
                let pollInterval = 3000; // Default 3 seconds
                if (status.step === 'processing_pages' || status.step === 'processing_pages_fallback') {
                    pollInterval = 2000; // More frequent updates during page processing
                }

                setTimeout(pollConversionStatus, pollInterval);
            } else if (status.status === 'completed') {
                updateDetailedStatus(status);
                statusMessage.className = 'status-message success';
                convertBtn.innerHTML = '✨ 重新转换';
                convertBtn.disabled = false;
                downloadContainer.classList.remove('hidden');
                isConverting = false;

                // Clean up page progress display
                const pageProgress = document.getElementById('pageProgress');
                if (pageProgress) {
                    pageProgress.remove();
                }
            } else if (status.status === 'error') {
                updateDetailedStatus(status);
                statusMessage.className = 'status-message error';
                convertBtn.disabled = false;
                convertBtn.innerHTML = '✨ 重新转换';
                isConverting = false;
            } else if (status.status === 'cancelled') {
                updateDetailedStatus(status);
                statusMessage.className = 'status-message warning';
                convertBtn.disabled = false;
                convertBtn.innerHTML = '✨ 重新转换';
                isConverting = false;
            }
        }
    } catch (error) {
        updateStatus('获取状态失败', 'error', 0);
        convertBtn.disabled = false;
        convertBtn.innerHTML = '✨ 重新转换';
        isConverting = false;
    }
}

function updateStatus(message, type, progress) {
    statusText.textContent = message;
    statusMessage.className = `status-message ${type}`;
    progressBar.style.width = `${progress}%`;
    progressBar.setAttribute('aria-valuenow', progress);
}

function updateDetailedStatus(data) {
    const { message, progress, step, status } = data;

    statusText.textContent = message;
    let statusClass = 'info';
    if (status === 'error') statusClass = 'error';
    else if (status === 'cancelled') statusClass = 'warning';

    statusMessage.className = `status-message ${statusClass}`;
    progressBar.style.width = `${progress}%`;
    progressBar.setAttribute('aria-valuenow', progress);

    // Update main progress percentage display
    const progressPercentage = document.getElementById('progressPercentage');
    if (progressPercentage) {
        progressPercentage.textContent = `${progress}%`;
    }

    // Add step indicator
    const stepIndicator = document.getElementById('stepIndicator');
    if (!stepIndicator) {
        const stepDiv = document.createElement('div');
        stepDiv.id = 'stepIndicator';
        stepDiv.style.fontSize = '0.9em';
        stepDiv.style.color = '#666';
        stepDiv.style.marginTop = '5px';
        statusMessage.appendChild(stepDiv);
    }

    const stepElement = document.getElementById('stepIndicator');
    if (stepElement && step) {
        // Show step description based on step name
        const stepDescriptions = {
            'uploading': '📤 正在上传文件',
            'transferring': '📡 正在传输数据',
            'preparing': '⏳ 服务器准备中',
            'upload_complete': '✅ 文件上传成功',
            'initialization': '🚀 开始初始化',
            'initializing': '🔧 设置转换器',
            'opening_document': '📂 打开PDF文档',
            'analyzing_document': '🔍 分析文档结构',
            'extracting_content': '📝 提取内容元素',
            'converting_elements': '🔄 转换PDF元素',
            'preserving_formatting': '🎨 保留格式布局',
            'generating_docx': '📄 生成Word文档',
            'processing_pages': '📖 处理PDF页面',
            'processing_pages_fallback': '📖 处理PDF页面 (备用模式)',
            'finalizing': '✨ 完成处理',
            'completed': '✅ 转换完成',
            'error': '❌ 发生错误'
        };

        stepElement.textContent = stepDescriptions[step] || `Step: ${step}`;

        // Show page progress if available
        if (data.current_page && data.total_pages) {
            const pageProgress = document.getElementById('pageProgress');
            if (!pageProgress) {
                const pageDiv = document.createElement('div');
                pageDiv.id = 'pageProgress';
                pageDiv.style.fontSize = '0.85em';
                pageDiv.style.color = '#555';
                pageDiv.style.marginTop = '3px';
                pageDiv.style.fontWeight = 'bold';
                statusMessage.appendChild(pageDiv);
            }

            const pageElement = document.getElementById('pageProgress');
            if (pageElement) {
                // Create a mini progress bar for pages
                const pagePercent = Math.round((data.current_page / data.total_pages) * 100);
                let etaHtml = '';
                if (data.eta) {
                    etaHtml = `<div style="margin-top: 3px; color: #6c757d;">${data.eta}</div>`;
                }
                pageElement.innerHTML = `
                    <div>页面进度: ${data.current_page} / ${data.total_pages} (${pagePercent}%)</div>
                    <div style="background-color: #e9ecef; border-radius: 4px; height: 6px; margin-top: 3px; width: 100%;">
                        <div style="background-color: #28a745; height: 100%; border-radius: 4px; width: ${pagePercent}%; transition: width 0.3s;"></div>
                    </div>
                    ${etaHtml}
                `;
            }
        }
    }

    // Add progress percentage in status message (smaller one)
    const progressElement = document.getElementById('progressPercent');
    if (progressElement) {
        progressElement.textContent = `${progress}%`;
    } else {
        const progressDiv = document.createElement('div');
        progressDiv.id = 'progressPercent';
        progressDiv.style.fontSize = '0.8em';
        progressDiv.style.fontWeight = 'bold';
        progressDiv.style.marginTop = '5px';
        progressDiv.textContent = `${progress}%`;
        statusMessage.appendChild(progressDiv);
    }
}

async function cleanupFiles() {
    if (currentTaskId) {
        try {
            await fetch(`/api/cleanup/${currentTaskId}`, {
                method: 'DELETE'
            });
        } catch (error) {
            console.error('Cleanup failed:', error);
        }
    }
}
