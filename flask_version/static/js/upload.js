// Upload page functionality

document.addEventListener('DOMContentLoaded', function() {
    const uploadZone = document.getElementById('upload-zone');
    const fileInput = document.getElementById('file-input');
    const fileListSection = document.getElementById('file-list-section');
    const fileStats = document.getElementById('file-stats');
    const fileList = document.getElementById('file-list');
    const proceedBtn = document.getElementById('proceed-btn');
    const loadingOverlay = document.getElementById('loading-overlay');
    
    // Drag and drop handlers
    uploadZone.addEventListener('dragover', function(e) {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });
    
    uploadZone.addEventListener('dragleave', function(e) {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
    });
    
    uploadZone.addEventListener('drop', function(e) {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFiles(files);
        }
    });
    
    // Click to upload
    uploadZone.addEventListener('click', function() {
        fileInput.click();
    });
    
    fileInput.addEventListener('change', function() {
        if (fileInput.files.length > 0) {
            handleFiles(fileInput.files);
        }
    });
    
    // Handle file upload
    function handleFiles(files) {
        // Show loading
        loadingOverlay.style.display = 'flex';
        
        // Create FormData
        const formData = new FormData();
        for (let i = 0; i < files.length; i++) {
            formData.append('files', files[i]);
        }
        
        // Upload files
        fetch('/api/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            loadingOverlay.style.display = 'none';
            
            if (data.error) {
                alert('Error: ' + data.error);
                return;
            }
            
            // Show file list
            displayFiles(data);
        })
        .catch(error => {
            loadingOverlay.style.display = 'none';
            alert('Upload failed: ' + error.message);
        });
    }
    
    // Display uploaded files
    function displayFiles(data) {
        fileListSection.style.display = 'block';
        
        // Stats
        fileStats.innerHTML = `
            <div class="stat">
                <span class="stat-label">Total Files:</span>
                <span class="stat-value">${data.files_count}</span>
            </div>
            <div class="stat">
                <span class="stat-label">CSV (Output):</span>
                <span class="stat-value">${data.csv_count}</span>
            </div>
            <div class="stat">
                <span class="stat-label">TXT (Input):</span>
                <span class="stat-value">${data.txt_count}</span>
            </div>
            <div class="stat">
                <span class="stat-label">Detected Unit:</span>
                <span class="stat-value">${data.unit}</span>
            </div>
        `;
        
        // File list
        let html = '';
        data.test_configs.forEach(config => {
            // Note: we're showing unique test configs, not individual files
        });
        
        // Show filenames
        const filenames = data.test_configs.length > 0 ? 
            `<p style="padding: 1rem; color: #666;">Detected ${data.test_configs.length} unique test configurations from ${data.files_count} files.</p>` :
            '';
        fileList.innerHTML = filenames;
        
        // Enable proceed button
        proceedBtn.disabled = false;
    }
});

// Clear files
function clearFiles() {
    fetch('/api/reset')
        .then(response => response.json())
        .then(data => {
            document.getElementById('file-list-section').style.display = 'none';
            document.getElementById('file-input').value = '';
            document.getElementById('proceed-btn').disabled = true;
        });
}

// Proceed to configure
function proceedToConfigure() {
    window.location.href = '/configure';
}
