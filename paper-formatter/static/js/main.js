// 论文格式自动转换工具 - 前端交互脚本

document.addEventListener('DOMContentLoaded', function() {
    // DOM元素
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const uploadBtn = document.getElementById('uploadBtn');
    const cancelBtn = document.getElementById('cancelBtn');
    const progressSection = document.getElementById('progressSection');
    const resultSection = document.getElementById('resultSection');
    const errorSection = document.getElementById('errorSection');
    const resultMessage = document.getElementById('resultMessage');
    const errorMessage = document.getElementById('errorMessage');
    const downloadBtn = document.getElementById('downloadBtn');
    const resetBtn = document.getElementById('retryBtn');
    const retryBtn = document.getElementById('retryBtn');

    let selectedFile = null;
    let downloadUrl = null;

    // 点击上传区域触发文件选择
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 文件选择
    fileInput.addEventListener('change', (e) => {
        handleFileSelect(e.target.files[0]);
    });

    // 拖拽上传
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        handleFileSelect(file);
    });

    // 处理文件选择
    function handleFileSelect(file) {
        if (!file) return;

        // 验证文件类型
        if (!file.name.endsWith('.docx')) {
            showError('只支持 .docx 格式的文件！');
            return;
        }

        // 验证文件大小 (10MB)
        if (file.size > 10 * 1024 * 1024) {
            showError('文件大小不能超过 10MB！');
            return;
        }

        selectedFile = file;
        fileName.textContent = file.name;
        uploadArea.style.display = 'none';
        fileInfo.style.display = 'block';
    }

    // 上传按钮
    uploadBtn.addEventListener('click', () => {
        if (!selectedFile) return;
        uploadFile();
    });

    // 取消按钮
    cancelBtn.addEventListener('click', () => {
        resetUpload();
    });

    // 重置上传
    function resetUpload() {
        selectedFile = null;
        fileInput.value = '';
        uploadArea.style.display = 'block';
        fileInfo.style.display = 'none';
        progressSection.style.display = 'none';
        resultSection.style.display = 'none';
        errorSection.style.display = 'none';
    }

    // 上传文件
    async function uploadFile() {
        const formData = new FormData();
        formData.append('file', selectedFile);

        // 显示进度
        fileInfo.style.display = 'none';
        progressSection.style.display = 'block';
        errorSection.style.display = 'none';

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            progressSection.style.display = 'none';

            if (response.ok && data.success) {
                // 成功
                downloadUrl = data.download_url;
                resultMessage.textContent = data.message;
                resultSection.style.display = 'block';
            } else {
                // 失败
                showError(data.error || '转换失败，请重试');
            }

        } catch (error) {
            progressSection.style.display = 'none';
            showError('网络错误，请检查连接后重试');
            console.error('Upload error:', error);
        }
    }

    // 显示错误
    function showError(message) {
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
        progressSection.style.display = 'none';
        resultSection.style.display = 'none';
    }

    // 下载按钮
    downloadBtn.addEventListener('click', () => {
        if (downloadUrl) {
            window.location.href = downloadUrl;
        }
    });

    // 重新开始按钮
    document.getElementById('resetBtn').addEventListener('click', () => {
        resetUpload();
    });

    // 重试按钮
    retryBtn.addEventListener('click', () => {
        resetUpload();
    });

    // 加载格式信息
    loadFormatInfo();

    async function loadFormatInfo() {
        try {
            const response = await fetch('/format-info');
            const data = await response.json();
            console.log('格式信息已加载:', data);
        } catch (error) {
            console.error('加载格式信息失败:', error);
        }
    }
});
