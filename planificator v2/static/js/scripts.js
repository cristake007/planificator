let holidays = [];
let generatedSchedule = null;


document.addEventListener('DOMContentLoaded', function() {});

function selectAllMonths() {
    document.querySelectorAll('.month-checkbox').forEach(checkbox => {
        checkbox.checked = true;
    });
}

function deselectAllMonths() {
    document.querySelectorAll('.month-checkbox').forEach(checkbox => {
        checkbox.checked = false;
    });
}

function addHolidays() {
    const input = document.getElementById('holidayInput');
    const dates = input.value
        .split(/[\n,]/)
        .map(d => d.trim())
        .filter(d => d.length > 0);

    for (const date of dates) {
        if (isValidDate(date) && !holidays.includes(date)) {
            holidays.push(date);
        }
    }
    updateHolidayDisplay();
    input.value = '';
}

function removeHoliday(index) {
    holidays.splice(index, 1);
    updateHolidayDisplay();
}

function updateHolidayDisplay() {
    const container = document.getElementById('holidayList');
    container.innerHTML = holidays.map((date, index) => `
        <span class="holiday-tag">
            ${date}
            <button type="button" class="btn-close ms-1 btn-close-white" style="font-size: 0.5rem;" onclick="removeHoliday(${index})"></button>
        </span>
    `).join(' ');
}

function isValidDate(dateStr) {
    const regex = /^\d{2}\.\d{2}\.\d{4}$/;
    if (!regex.test(dateStr)) return false;

    const parts = dateStr.split('.');
    const date = new Date(parts[2], parts[1] - 1, parts[0]);
    return date instanceof Date && !isNaN(date);
}

document.getElementById('randomness').addEventListener('input', function() {
    document.getElementById('randomnessValue').textContent = this.value;
});

async function exportSchedule() {
    try {
        const response = await fetch('/export_schedule', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                schedule: generatedSchedule,
                year: document.getElementById('year').value,
                holidays: holidays
            })
        });

        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `course_schedule_${document.getElementById('year').value}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } else {
            throw new Error('Export failed');
        }
    } catch (error) {
        console.error('Error exporting schedule:', error);
        document.getElementById('errorAlert').textContent = 'Error exporting schedule: ' + error.message;
        document.getElementById('errorAlert').style.display = 'block';
    }
}

function displaySchedule(schedule) {
    const tableBody = document.getElementById('scheduleTableBody');
    tableBody.innerHTML = '';

    // Sort the schedule by original_order to maintain the original CSV order
    const sortedSchedule = [...schedule].sort((a, b) => a.original_order - b.original_order);

    const courseSchedules = new Map();

    // Collect courses using the original order
    sortedSchedule.forEach(item => {
        if (!courseSchedules.has(item.Title)) {
            courseSchedules.set(item.Title, {
                name: item.Title,
                duration: item['Durata Curs'],
                months: {},
                originalOrder: item.original_order
            });
        }
    });

    // Fill in the dates for each course
    sortedSchedule.forEach(item => {
        courseSchedules.get(item.Title).months[item.month] = item.date_range;
    });

    // Update course summary
    const courseCounter = document.getElementById('courseCounter');
    courseCounter.innerHTML = `
        <strong>Course Summary:</strong><br>
        Total unique courses loaded: ${courseSchedules.size}<br>
        Total scheduled sessions: ${schedule.length}
    `;
    courseCounter.style.display = 'block';

    // Create rows and maintain the original CSV order
    Array.from(courseSchedules.values())
        .forEach(course => {
            const row = document.createElement('tr');

            // Course name cell
            const nameCell = document.createElement('td');
            nameCell.className = 'course-name-cell';
            nameCell.textContent = course.name;
            row.appendChild(nameCell);

            // Duration cell
            const daysCell = document.createElement('td');
            daysCell.className = 'days-cell';
            daysCell.textContent = course.duration;
            row.appendChild(daysCell);

            // Add cells for each month
            for (let month = 1; month <= 12; month++) {
                const cell = document.createElement('td');
                cell.className = 'text-center';
                cell.textContent = course.months[month] || '';
                row.appendChild(cell);
            }

            tableBody.appendChild(row);
        });

    document.getElementById('scheduleResult').style.display = 'block';
}


document.getElementById('scheduleForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const formData = new FormData();
    const fileInput = document.getElementById('inputFile');
    const yearInput = document.getElementById('year');

    // Get selected months
    const selectedMonths = Array.from(document.querySelectorAll('.month-checkbox:checked'))
        .map(checkbox => checkbox.value);

    if (selectedMonths.length === 0) {
        alert('Please select at least one month');
        return;
    }

    if (!fileInput.files[0]) {
        alert('Please select a file');
        return;
    }

    formData.append('input_file', fileInput.files[0]);
    formData.append('year', yearInput.value);
    formData.append('holidays', holidays.join(','));
    formData.append('months', selectedMonths.join(','));
    formData.append('randomness', document.getElementById('randomness').value);

    // Show loading spinner
    document.getElementById('loadingSpinner').style.display = 'block';
    document.getElementById('errorAlert').style.display = 'none';

    try {
        const response = await fetch('/generate_schedule', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (data.success) {
            generatedSchedule = data.schedule;
            displaySchedule(data.schedule);
            document.getElementById('exportBtn').style.display = 'inline-block';
        } else {
            document.getElementById('errorAlert').textContent = data.error;
            document.getElementById('errorAlert').style.display = 'block';
        }
    } catch (error) {
        document.getElementById('errorAlert').textContent = 'Error generating schedule: ' + error.message;
        document.getElementById('errorAlert').style.display = 'block';
    } finally {
        document.getElementById('loadingSpinner').style.display = 'none';
    }
});


document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('xmlForm');
    const fileInput = document.getElementById('inputFile');
    const fileName = document.getElementById('fileName');
    const loadingSpinner = document.getElementById('loadingSpinner');
    const errorAlert = document.getElementById('errorAlert');
    const successMessage = document.getElementById('successMessage');

    // Debug logging
    console.log('Script loaded', {
        form: !!form,
        fileInput: !!fileInput,
        fileName: !!fileName,
        loadingSpinner: !!loadingSpinner,
        errorAlert: !!errorAlert,
        successMessage: !!successMessage
    });

    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            fileName.textContent = `Selected file: ${file.name}`;
            console.log('File selected:', file.name);
        } else {
            fileName.textContent = '';
        }
    });

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        console.log('Form submitted');

        const file = fileInput.files[0];
        if (!file) {
            showError('Please select a file');
            return;
        }

        const formData = new FormData();
        formData.append('input_file', file);

        loadingSpinner.style.display = 'flex';
        errorAlert.style.display = 'none';
        successMessage.style.display = 'none';

        try {
            console.log('Sending request to /format-xml');
            const response = await fetch('/format-xml', {
                method: 'POST',
                body: formData
            });

            console.log('Response received:', response.status);

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'formatted_courses.xml';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);

                successMessage.style.display = 'block';
                successMessage.style.opacity = '1';
                setTimeout(() => {
                    successMessage.style.display = 'none';
                }, 3000);

                form.reset();
                fileName.textContent = '';
                console.log('File downloaded successfully');
            } else {
                const error = await response.json();
                throw new Error(error.error || 'Failed to generate XML');
            }
        } catch (error) {
            console.error('Error:', error);
            showError(error.message);
        } finally {
            loadingSpinner.style.display = 'none';
        }
    });

    function showError(message) {
        console.error('Showing error:', message);
        errorAlert.textContent = message;
        errorAlert.style.display = 'block';
    }
});


document.addEventListener('DOMContentLoaded', function () {
    const wordForm = document.getElementById('wordForm');
    const inputFile = document.getElementById('inputFile');
    const loadingSpinner = document.getElementById('loadingSpinner');
    const successMessage = document.getElementById('successMessage');
    const errorAlert = document.getElementById('errorAlert');

    if (wordForm) {
        wordForm.addEventListener('submit', async (e) => {
            e.preventDefault(); // Prevent page refresh
            const formData = new FormData();
            const file = inputFile.files[0];
            formData.append('input_file', file);

            // Show the loading spinner
            loadingSpinner.style.display = 'block';

            try {
                const response = await fetch('/convert_word', {
                    method: 'POST',
                    body: formData,
                });

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(errorData.error || 'An unknown error occurred');
                }

                // Convert response to blob and trigger file download
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'courses.xlsx';
                a.click();
                window.URL.revokeObjectURL(url); // Clean up URL

                // Show success message
                successMessage.style.display = 'block';
                setTimeout(() => {
                    successMessage.style.display = 'none';
                }, 5000);
            } catch (error) {
                // Show error message
                errorAlert.textContent = error.message;
                errorAlert.style.display = 'block';
                setTimeout(() => {
                    errorAlert.style.display = 'none';
                }, 5000);
            } finally {
                // Hide the loading spinner
                loadingSpinner.style.display = 'none';
            }
        });
    }
});