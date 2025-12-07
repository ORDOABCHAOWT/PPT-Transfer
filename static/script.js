// macOS Sequoia Style - PPT Transfer JavaScript

document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const uploadContent = dropZone.querySelector('.upload-content');
    const extractBtn = document.getElementById('extractBtn');
    const resultOverlay = document.getElementById('resultOverlay');
    const removeFileBtn = document.getElementById('removeFile');
    const extractAnotherBtn = document.getElementById('extractAnother');

    let selectedFile = null;

    // Click to select file
    dropZone.addEventListener('click', () => {
        if (!selectedFile) {
            fileInput.click();
        }
    });

    // File input change
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    // Drag and drop events
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');

        if (e.dataTransfer.files.length > 0) {
            const file = e.dataTransfer.files[0];
            if (file.name.endsWith('.pptx')) {
                handleFile(file);
            } else {
                showError('请选择 PPTX 文件');
            }
        }
    });

    // Handle file selection
    function handleFile(file) {
        selectedFile = file;

        // Update UI
        uploadContent.style.display = 'none';
        fileInfo.style.display = 'flex';

        document.getElementById('filename').textContent = file.name;
        document.getElementById('filesize').textContent = formatSize(file.size);

        extractBtn.disabled = false;

        // Add fade-in animation
        fileInfo.classList.add('fade-in');
    }

    // Remove file
    removeFileBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        resetUpload();
    });

    // Reset upload
    function resetUpload() {
        selectedFile = null;
        fileInput.value = '';
        uploadContent.style.display = 'block';
        fileInfo.style.display = 'none';
        extractBtn.disabled = true;
    }

    // Extract button
    extractBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        const columnSort = document.querySelector('input[name="column_sort"]').checked;
        const keepFormat = document.querySelector('input[name="keep_format"]').checked;

        // Update button state
        extractBtn.disabled = true;

        // 显示进度条
        showProgress();

        try {
            // Create form data
            const formData = new FormData();
            formData.append('file', selectedFile);
            formData.append('column_sort', columnSort);
            formData.append('keep_format', keepFormat);

            // Send request to start extraction
            const response = await fetch('/extract', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (response.ok && result.success && result.task_id) {
                // 连接到进度流
                connectToProgress(result.task_id);
            } else {
                hideProgress();
                showError(result.error || '提取失败，请重试');
                extractBtn.disabled = false;
            }
        } catch (error) {
            console.error('Error:', error);
            hideProgress();
            showError('网络错误，请重试');
            extractBtn.disabled = false;
        }
    });

    // 连接到进度流
    function connectToProgress(taskId) {
        const eventSource = new EventSource(`/progress/${taskId}`);

        eventSource.onmessage = (event) => {
            const data = JSON.parse(event.data);

            if (data.status === 'progress') {
                updateProgress(data.percent, data.message);
            } else if (data.status === 'completed') {
                eventSource.close();
                hideProgress();
                showResult({
                    totalSlides: data.total_slides,
                    textBlocks: data.text_blocks,
                    fileSize: data.file_size,
                    downloadUrl: data.download_url
                });
            } else if (data.status === 'error') {
                eventSource.close();
                hideProgress();
                showError(data.message || '提取失败');
                extractBtn.disabled = false;
            }
        };

        eventSource.onerror = (error) => {
            console.error('SSE Error:', error);
            eventSource.close();
            hideProgress();
            showError('连接中断，请重试');
            extractBtn.disabled = false;
        };
    }

    // 进度条相关函数
    function showProgress() {
        // 创建进度条容器（如果不存在）
        let progressContainer = document.getElementById('progressContainer');
        if (!progressContainer) {
            progressContainer = document.createElement('div');
            progressContainer.id = 'progressContainer';
            progressContainer.className = 'progress-container';
            progressContainer.innerHTML = `
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
                <div class="progress-text" id="progressText">正在准备...</div>
            `;
            // 插入到上传区域之后
            const uploadZone = document.getElementById('dropZone');
            const settingsPanel = document.querySelector('.settings-panel');
            uploadZone.parentNode.insertBefore(progressContainer, settingsPanel);
        }
        progressContainer.classList.add('show');
        updateProgress(0, '正在准备...');
    }

    function hideProgress() {
        const progressContainer = document.getElementById('progressContainer');
        if (progressContainer) {
            progressContainer.classList.remove('show');
        }
        extractBtn.disabled = false;
    }

    function updateProgress(percent, message) {
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');

        if (progressFill) {
            progressFill.style.width = percent + '%';
        }
        if (progressText) {
            progressText.textContent = message || `${percent}%`;
        }
    }

    // Show result overlay
    function showResult(data) {
        document.getElementById('totalSlides').textContent = data.totalSlides;
        document.getElementById('textBlocks').textContent = data.textBlocks || '-';
        document.getElementById('fileSize').textContent = data.fileSize;

        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = data.downloadUrl;

        resultOverlay.classList.add('show');
    }

    // Extract another file
    extractAnotherBtn.addEventListener('click', () => {
        resultOverlay.classList.remove('show');
        resetUpload();
    });

    // Show error message
    function showError(message) {
        alert(message);
    }

    // Format file size
    function formatSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
    }

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        // ESC to close result overlay
        if (e.key === 'Escape' && resultOverlay.classList.contains('show')) {
            resultOverlay.classList.remove('show');
            resetUpload();
        }

        // CMD/CTRL + O to open file
        if ((e.metaKey || e.ctrlKey) && e.key === 'o') {
            e.preventDefault();
            fileInput.click();
        }
    });

    // Close result overlay when clicking outside
    resultOverlay.addEventListener('click', (e) => {
        if (e.target === resultOverlay) {
            resultOverlay.classList.remove('show');
            resetUpload();
        }
    });
});
