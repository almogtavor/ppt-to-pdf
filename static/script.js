document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const browseBtn = document.querySelector('.browse-btn');
    const convertBtn = document.getElementById('convertBtn');
    const filesContainer = document.getElementById('files');
    const progressModal = document.getElementById('progressModal');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');

    let files = [];

    // Drag and drop handlers
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
        handleFiles(e.dataTransfer.files);
    });

    // Browse button handler
    browseBtn.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });

    // Convert button handler
    convertBtn.addEventListener('click', convertFiles);

    function handleFiles(newFiles) {
        for (let file of newFiles) {
            if (isValidFile(file)) {
                files.push(file);
            }
        }
        updateFileList();
    }

    function isValidFile(file) {
        const validTypes = ['application/pdf', 'application/vnd.ms-powerpoint', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'];
        return validTypes.includes(file.type);
    }

    function updateFileList() {
        filesContainer.innerHTML = '';
        files.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <span class="file-name">${file.name}</span>
                <button class="remove-btn" data-index="${index}">×</button>
            `;
            filesContainer.appendChild(fileItem);
        });

        // Add remove button handlers
        document.querySelectorAll('.remove-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const index = parseInt(e.target.dataset.index);
                files.splice(index, 1);
                updateFileList();
            });
        });

        convertBtn.disabled = files.length === 0;
    }

    async function convertFiles() {
        if (files.length === 0) return;

        const formData = new FormData();
        files.forEach(file => {
            formData.append('files[]', file);
        });

        // Add settings to form data
        formData.append('slides_per_row', document.getElementById('slidesPerRow').value);
        formData.append('gap', document.getElementById('gap').value);
        formData.append('margin', document.getElementById('margin').value);
        formData.append('top_margin', document.getElementById('topMargin').value);
        formData.append('single_file', document.getElementById('singleFile').checked);

        // Show progress modal
        progressModal.style.display = 'flex';
        progressBar.style.width = '0%';
        progressText.textContent = 'Processing...';

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error('Conversion failed');
            }

            // Get the blob from the response
            const blob = await response.blob();
            
            // Create a download link
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = document.getElementById('singleFile').checked ? 'combined.pdf' : 'converted.zip';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            // Update progress
            progressBar.style.width = '100%';
            progressText.textContent = 'Conversion completed!';
            
            // Clear files after successful conversion
            files = [];
            updateFileList();

            // Hide modal after a delay
            setTimeout(() => {
                progressModal.style.display = 'none';
            }, 2000);

        } catch (error) {
            progressText.textContent = `Error: ${error.message}`;
            progressBar.style.backgroundColor = 'var(--error-color)';
        }
    }
}); 