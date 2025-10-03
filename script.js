
let studentData = null;

// Load the Excel file when the page loads
window.onload = function() {
    loadExcelFile();
};

function loadExcelFile() {
    const loadingIndicator = document.querySelector('.loading');
    if (loadingIndicator) {
        loadingIndicator.style.display = 'inline-block';
    }

    fetch('data.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Failed to load Excel file');
            }
            return response.arrayBuffer();
        })
        .then(buffer => {
            const workbook = XLSX.read(buffer, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            studentData = XLSX.utils.sheet_to_json(firstSheet);
            console.log('Data loaded successfully:', studentData.length, 'records found');
        })
        .catch(error => {
            console.error('Error loading Excel file:', error);
            showError('Error loading student data. Please try again later.');
        })
        .finally(() => {
            if (loadingIndicator) {
                loadingIndicator.style.display = 'none';
            }
        });
}

function searchStudent() {
    const enrollmentNumber = document.getElementById('enrollmentNumber').value.trim();
    const resultContainer = document.getElementById('resultContainer');
    const loadingIndicator = document.querySelector('.loading');
    const resetBtn = document.querySelector('.reset-btn');
    
    if (!enrollmentNumber) {
        showError('Please enter an enrollment number');
        return;
    }

    if (!studentData) {
        showError('Student data is still loading. Please try again in a moment.');
        return;
    }

    loadingIndicator.style.display = 'inline-block';

    setTimeout(() => {
        try {
            const student = studentData.find(s => 
                String(s.enrollmentNumber).toLowerCase() === enrollmentNumber.toLowerCase()
            );

            if (student) {
                resultContainer.style.display = 'block';
                resultContainer.className = 'result-container success-state';
                resultContainer.innerHTML = `
                    <div class="field">
                        <span class="label">Full Name</span>
                        <span class="value">${student['Full Name'] || 'N/A'}</span>
                    </div>
                    <div class="field">
                        <span class="label">Program</span>
                        <span class="value">${student['PROGRAM'] || 'N/A'}</span>
                    </div>
                    <div class="field">
                        <span class="label">Course</span>
                        <span class="value">${student['COURSE'] || 'N/A'}</span>
                    </div>
                    <div class="field">
                        <span class="label">Campus</span>
                        <span class="value">${student['Campus'] || 'N/A'}</span>
                    </div>
                    <div class="field">
                        <span class="label">Email</span>
                        <span class="value">${student['Official Email'] || 'N/A'}</span>
                    </div>
                `;
                resetBtn.style.display = 'inline-block';
            } else {
                resultContainer.style.display = 'block';
                resultContainer.className = 'result-container error-state';
                resultContainer.innerHTML = `
                    <div style="text-align: center; padding: 20px; color: #dc3545;">
                        <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="12" y1="8" x2="12" y2="12"></line>
                            <line x1="12" y1="16" x2="12.01" y2="16"></line>
                        </svg>
                        <h2 style="margin: 15px 0;">No Record Found</h2>
                        <p style="margin: 0;">No student found with enrollment number: ${enrollmentNumber}</p>
                    </div>
                `;
                resetBtn.style.display = 'inline-block';
            }
        } catch (error) {
            console.error('Error:', error);
            showError('An error occurred while searching');
        } finally {
            loadingIndicator.style.display = 'none';
        }
    }, 500);
}

function resetForm() {
    // Clear input
    document.getElementById('enrollmentNumber').value = '';
    
    // Hide result container
    document.getElementById('resultContainer').style.display = 'none';
    
    // Hide reset button
    document.querySelector('.reset-btn').style.display = 'none';
    
    // Reset label position
    document.getElementById('enrollmentNumber').focus();
    document.getElementById('enrollmentNumber').blur();
}

function showError(message) {
    const resultContainer = document.getElementById('resultContainer');
    const resetBtn = document.querySelector('.reset-btn');
    
    resultContainer.style.display = 'block';
    resultContainer.className = 'result-container error-state';
    resultContainer.innerHTML = `
        <div style="text-align: center; padding: 20px; color: #dc3545;">
            <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="8" x2="12" y2="12"></line>
                <line x1="12" y1="16" x2="12.01" y2="16"></line>
            </svg>
            <h2 style="margin: 15px 0;">Error</h2>
            <p style="margin: 0;">${message}</p>
        </div>
    `;
    resetBtn.style.display = 'inline-block';
}