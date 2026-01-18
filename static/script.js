const uploadBox = document.getElementById('uploadBox');
const fileInput = document.getElementById('fileInput');
const selectedFile = document.getElementById('selectedFile');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeFile = document.getElementById('removeFile');
const convertBtn = document.getElementById('convertBtn');
const loading = document.getElementById('loading');
const successMessage = document.getElementById('successMessage');
const errorMessage = document.getElementById('errorMessage');
const errorText = document.getElementById('errorText');
const convertAnother = document.getElementById('convertAnother');
const tryAgain = document.getElementById('tryAgain');

let selectedFileObj = null;

// Handle file selection
fileInput.addEventListener('change', (e) => {
    handleFileSelect(e.target.files[0]);
});

// Drag and drop handlers
uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadBox.classList.add('drag-over');
});

uploadBox.addEventListener('dragleave', () => {
    uploadBox.classList.remove('drag-over');
});

uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadBox.classList.remove('drag-over');
    
    const file = e.dataTransfer.files[0];
    if (file) {
        handleFileSelect(file);
    }
});

uploadBox.addEventListener('click', () => {
    fileInput.click();
});

function handleFileSelect(file) {
    // Validate file type
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                       'application/vnd.ms-excel',
                       'application/vnd.ms-excel.sheet.macroEnabled.12',
                       'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                       'application/vnd.ms-powerpoint',
                       'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                       'application/msword',
                       'text/plain'];
    const validExtensions = ['.xlsx', '.xls', '.xlsm', '.pptx', '.ppt', '.docx', '.doc', '.txt'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    
    if (!validTypes.includes(file.type) && !validExtensions.includes(fileExtension)) {
        showError('Please select a valid Excel (.xlsx, .xls, .xlsm), PowerPoint (.pptx, .ppt), Word (.docx, .doc), or Text (.txt) file');
        return;
    }
    
    // Validate file size (16MB)
    if (file.size > 16 * 1024 * 1024) {
        showError('File size must be less than 16MB');
        return;
    }
    
    selectedFileObj = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    
    uploadBox.style.display = 'none';
    selectedFile.style.display = 'block';
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

removeFile.addEventListener('click', () => {
    resetUpload();
});

convertAnother.addEventListener('click', () => {
    resetUpload();
});

tryAgain.addEventListener('click', () => {
    resetUpload();
});

function resetUpload() {
    selectedFileObj = null;
    fileInput.value = '';
    uploadBox.style.display = 'block';
    selectedFile.style.display = 'none';
    loading.style.display = 'none';
    successMessage.style.display = 'none';
    errorMessage.style.display = 'none';
}

convertBtn.addEventListener('click', async () => {
    if (!selectedFileObj) return;
    
    // Show loading
    selectedFile.style.display = 'none';
    loading.style.display = 'block';
    
    const formData = new FormData();
    formData.append('file', selectedFileObj);
    
    try {
        const response = await fetch('/convert', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Conversion failed');
        }
        
        // Download the PDF
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = selectedFileObj.name.replace(/\.[^/.]+$/, '') + '.pdf';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        // Show success message
        loading.style.display = 'none';
        successMessage.style.display = 'block';
        
    } catch (error) {
        console.error('Error:', error);
        showError(error.message);
    }
});

function showError(message) {
    loading.style.display = 'none';
    selectedFile.style.display = 'none';
    errorMessage.style.display = 'block';
    errorText.textContent = message;
}
