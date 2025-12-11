// Configuration page functionality

document.addEventListener('DOMContentLoaded', function() {
    const configFileInput = document.getElementById('config-file-input');
    const equipmentModelInput = document.getElementById('equipment-model');
    const equipmentNumberInput = document.getElementById('equipment-number');
    const reportNamePreview = document.getElementById('report-name-preview');
    
    // Handle config file load
    configFileInput.addEventListener('change', function() {
        if (configFileInput.files.length > 0) {
            loadConfigFromFile(configFileInput.files[0]);
        }
    });
    
    // Update report name preview when inputs change
    function updateReportNamePreview() {
        const model = equipmentModelInput.value.trim();
        const number = equipmentNumberInput.value.trim();
        
        let reportName = '';
        if (model && number) {
            reportName = `${model}_${number}`;
        } else if (model) {
            reportName = model;
        } else if (number) {
            reportName = number;
        } else {
            reportName = 'Equipment';
        }
        
        reportNamePreview.textContent = reportName;
    }
    
    equipmentModelInput.addEventListener('input', updateReportNamePreview);
    equipmentNumberInput.addEventListener('input', updateReportNamePreview);
    
    // Initialize preview
    updateReportNamePreview();
});

// Collect measurement type selections
function getMeasurementTypeSelections() {
    const selections = {};
    const selects = document.querySelectorAll('.measurement-type-select');
    selects.forEach(select => {
        const filename = select.dataset.filename;
        selections[filename] = select.value;
    });
    return selections;
}

// Collect configuration from table
function getConfigurations() {
    const configs = [];
    const rows = document.querySelectorAll('#config-table tbody tr');
    
    rows.forEach(row => {
        const testValue = parseFloat(row.dataset.testValue);
        const rangeSetting = row.dataset.rangeSetting;
        const ioType = row.dataset.ioType;
        
        const rangeInput = row.querySelector('.range-input').value.trim();
        const reference = parseFloat(row.querySelector('.reference-input').value);
        const tolerance = parseFloat(row.querySelector('.tolerance-input').value);
        
        if (isNaN(reference) || isNaN(tolerance)) {
            throw new Error(`Invalid values for test point ${testValue} ${filesInfo.unit}`);
        }
        
        if (tolerance < 0) {
            throw new Error(`Tolerance must be positive for test point ${testValue} ${filesInfo.unit}`);
        }
        
        configs.push({
            test_value: testValue,
            range_setting: rangeSetting,
            io_type: ioType,
            range_input: rangeInput || 'N/A',
            reference: reference,
            tolerance: tolerance
        });
    });
    
    return configs;
}

// Process files
function processFiles() {
    const loadingOverlay = document.getElementById('loading-overlay');
    const loadingMessage = document.getElementById('loading-message');
    
    try {
        const measurementTypes = getMeasurementTypeSelections();
        const configs = getConfigurations();
        const equipmentModel = document.getElementById('equipment-model').value.trim();
        const equipmentNumber = document.getElementById('equipment-number').value.trim();
        
        loadingOverlay.style.display = 'flex';
        loadingMessage.textContent = 'Processing files...';
        
        fetch('/api/process', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                measurement_types: measurementTypes,
                configs: configs,
                equipment_model: equipmentModel,
                equipment_number: equipmentNumber
            })
        })
        .then(response => response.json())
        .then(data => {
            loadingOverlay.style.display = 'none';
            
            if (data.error) {
                alert('Error: ' + data.error + '\n\n' + (data.traceback || ''));
                return;
            }
            
            // Redirect to results page
            window.location.href = '/results';
        })
        .catch(error => {
            loadingOverlay.style.display = 'none';
            alert('Processing failed: ' + error.message);
        });
        
    } catch (error) {
        alert('Configuration error: ' + error.message);
    }
}

// Save configuration
function saveConfig() {
    try {
        const configs = getConfigurations();
        const configData = {
            unit: filesInfo.unit,
            configurations: configs.map(c => ({
                test_value: c.test_value,
                range_setting: c.range_setting === 'N/A' ? null : c.range_setting,
                io_type: c.io_type,
                range_input: c.range_input,
                reference: c.reference.toString(),
                tolerance: c.tolerance.toString()
            }))
        };
        
        fetch('/api/save-config', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(configData)
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // Download the config file
                const blob = new Blob([JSON.stringify(configData, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'test_config.json';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                showNotification('Configuration saved successfully!');
            }
        })
        .catch(error => {
            alert('Save failed: ' + error.message);
        });
        
    } catch (error) {
        alert('Configuration error: ' + error.message);
    }
}

// Load configuration
function loadConfig() {
    document.getElementById('config-file-input').click();
}

function loadConfigFromFile(file) {
    const formData = new FormData();
    formData.append('file', file);
    
    fetch('/api/load-config', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.error) {
            alert('Error loading config: ' + data.error);
            return;
        }
        
        applyConfig(data.config);
        showNotification('Configuration loaded successfully!');
    })
    .catch(error => {
        alert('Load failed: ' + error.message);
    });
}

function applyConfig(configData) {
    // Check unit compatibility
    if (configData.unit && configData.unit !== filesInfo.unit) {
        if (!confirm(`Config file unit (${configData.unit}) differs from current unit (${filesInfo.unit}). Load anyway?`)) {
            return;
        }
    }
    
    const configurations = configData.configurations || [];
    const rows = document.querySelectorAll('#config-table tbody tr');
    
    let loadedCount = 0;
    
    configurations.forEach(config => {
        // Normalize the config values for comparison
        const configTestValue = parseFloat(config.test_value);
        const configRangeSetting = normalizeRangeSetting(config.range_setting);
        const configIoType = config.io_type;
        
        // Find matching row by iterating through all rows
        rows.forEach(row => {
            const rowTestValue = parseFloat(row.dataset.testValue);
            const rowRangeSetting = normalizeRangeSetting(row.dataset.rangeSetting);
            const rowIoType = row.dataset.ioType;
            
            // Compare with tolerance for floating-point values
            const testValueMatch = Math.abs(rowTestValue - configTestValue) < 0.0001;
            const rangeMatch = rowRangeSetting === configRangeSetting;
            const ioTypeMatch = rowIoType === configIoType;
            
            if (testValueMatch && rangeMatch && ioTypeMatch) {
                const rangeInput = row.querySelector('.range-input');
                const referenceInput = row.querySelector('.reference-input');
                const toleranceInput = row.querySelector('.tolerance-input');
                
                if (config.range_input !== undefined && rangeInput) {
                    rangeInput.value = config.range_input;
                }
                if (config.reference !== undefined && referenceInput) {
                    referenceInput.value = config.reference;
                }
                if (config.tolerance !== undefined && toleranceInput) {
                    toleranceInput.value = config.tolerance;
                }
                loadedCount++;
            }
        });
    });
    
    console.log(`Applied ${loadedCount} configurations`);
}

// Helper function to normalize range setting values for comparison
function normalizeRangeSetting(value) {
    if (value === null || value === undefined || value === 'None' || value === 'null') {
        return 'N/A';
    }
    return String(value).trim();
}

// Show notification
function showNotification(message) {
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #2E7D32;
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 1001;
        animation: slideIn 0.3s ease;
    `;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
            document.body.removeChild(notification);
        }, 300);
    }, 3000);
}

// Add CSS animation for notifications
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    @keyframes slideOut {
        from { transform: translateX(0); opacity: 1; }
        to { transform: translateX(100%); opacity: 0; }
    }
`;
document.head.appendChild(style);
