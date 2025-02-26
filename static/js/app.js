// Drag and drop file upload
function setupFileUpload() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('excel_file');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const removeFileBtn = document.getElementById('remove-file');

    if (!dropZone) return;

    // Highlight drop area when item is dragged over
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('border-primary');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('border-primary');
    });

    // Handle dropped files
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('border-primary');
        
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            updateFileInfo();
        }
    });

    // Handle file selection via the browse button
    fileInput.addEventListener('change', updateFileInfo);

    // Remove selected file
    if (removeFileBtn) {
        removeFileBtn.addEventListener('click', () => {
            fileInput.value = '';
            fileInfo.classList.add('d-none');
        });
    }

    function updateFileInfo() {
        if (fileInput.files.length > 0) {
            fileName.textContent = fileInput.files[0].name;
            fileInfo.classList.remove('d-none');
        } else {
            fileInfo.classList.add('d-none');
        }
    }
}

// Loading animations
function setupLoadingIndicators() {
    const forms = document.querySelectorAll('form[data-loading]');
    
    forms.forEach(form => {
        form.addEventListener('submit', () => {
            const submitBtn = form.querySelector('[type="submit"]');
            const loadingText = form.getAttribute('data-loading') || 'Processing...';
            
            if (submitBtn && form.checkValidity()) {
                const originalText = submitBtn.innerHTML;
                submitBtn.disabled = true;
                submitBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-2"></span>${loadingText}`;
                
                // Store original text for restoration if needed
                submitBtn.setAttribute('data-original-text', originalText);
            }
        });
    });
}

// Toast notifications
function showToast(message, type = 'info') {
    const toastContainer = document.getElementById('toast-container');
    if (!toastContainer) {
        const container = document.createElement('div');
        container.id = 'toast-container';
        container.className = 'position-fixed bottom-0 end-0 p-3';
        document.body.appendChild(container);
    }
    
    const toastId = 'toast-' + Date.now();
    const toast = document.createElement('div');
    toast.className = `toast align-items-center text-white bg-${type}`;
    toast.id = toastId;
    toast.setAttribute('role', 'alert');
    toast.setAttribute('aria-live', 'assertive');
    toast.setAttribute('aria-atomic', 'true');
    
    toast.innerHTML = `
        <div class="d-flex">
            <div class="toast-body">
                ${message}
            </div>
            <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
        </div>
    `;
    
    document.getElementById('toast-container').appendChild(toast);
    const toastInstance = new bootstrap.Toast(toast);
    toastInstance.show();
    
    // Auto-remove from DOM after hiding
    toast.addEventListener('hidden.bs.toast', () => {
        toast.remove();
    });
}

// Initialize all components
document.addEventListener('DOMContentLoaded', () => {
    setupFileUpload();
    setupLoadingIndicators();
    
    // Add data-loading attribute to forms that need loading indicators
    const forms = document.querySelectorAll('#uploadForm, #reconcileForm, #formulaForm, #chatForm');
    forms.forEach(form => {
        form.setAttribute('data-loading', 'Processing...');
    });
});