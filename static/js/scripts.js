let holidays = [];
let generatedSchedule = null;

function showElement(el, displayClass = 'block') {
    if (!el) return;
    el.classList.remove('app-hidden');
    if (displayClass === 'flex') {
        el.classList.add('flex');
    }
}

function hideElement(el) {
    if (!el) return;
    el.classList.add('app-hidden');
    el.classList.remove('flex');
}

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
        <span class="inline-flex items-center gap-1 rounded-md bg-secondary text-secondary-foreground px-2 py-1 text-xs font-medium border border-secondary-line">
            ${date}
            <button type="button" class="inline-flex items-center justify-center size-4 rounded bg-secondary-hover text-secondary-foreground hover:bg-secondary-focus text-[0.6rem]" onclick="removeHoliday(${index})">×</button>
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

const randomnessInput = document.getElementById('randomness');
if (randomnessInput) {
    randomnessInput.addEventListener('input', function() {
        const randomnessValue = document.getElementById('randomnessValue');
        if (randomnessValue) {
            randomnessValue.textContent = this.value;
        }
    });
}

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
        showElement(document.getElementById('errorAlert'));
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
                investitie: item.investitie || item.Investitie || '',
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
    showElement(courseCounter);

    // Create rows and maintain the original CSV order
    Array.from(courseSchedules.values())
        .forEach(course => {
            const row = document.createElement('tr');

            // Course name cell
            const nameCell = document.createElement('td');
            nameCell.className = 'cell-wide';
            nameCell.textContent = course.name;
            row.appendChild(nameCell);

            // Duration cell
            const daysCell = document.createElement('td');
            daysCell.className = 'cell-tight';
            daysCell.textContent = course.duration;
            row.appendChild(daysCell);

            // Investitie cell
            const investitieCell = document.createElement('td');
            investitieCell.className = 'cell-tight text-center';
            investitieCell.textContent = course.investitie || '';
            row.appendChild(investitieCell);

            // Add cells for each month
            for (let month = 1; month <= 12; month++) {
                const cell = document.createElement('td');
                cell.className = 'text-center';
                cell.textContent = course.months[month] || '';
                row.appendChild(cell);
            }

            tableBody.appendChild(row);
        });

    showElement(document.getElementById('scheduleResult'));
}


const scheduleForm = document.getElementById('scheduleForm');
if (scheduleForm) {
scheduleForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const errorAlert = document.getElementById('errorAlert');

    const formData = new FormData();
    const fileInput = document.getElementById('inputFile');
    const yearInput = document.getElementById('year');

    // Get selected months
    const selectedMonths = Array.from(document.querySelectorAll('.month-checkbox:checked'))
        .map(checkbox => checkbox.value);

    if (selectedMonths.length === 0) {
        errorAlert.textContent = 'Please select at least one month before generating the schedule.';
        showElement(errorAlert);
        return;
    }

    if (!fileInput.files[0]) {
        errorAlert.textContent = 'Please upload an input file (.csv, .xlsx, or .xls) before generating the schedule.';
        showElement(errorAlert);
        return;
    }

    formData.append('input_file', fileInput.files[0]);
    formData.append('year', yearInput.value);
    formData.append('holidays', holidays.join(','));
    formData.append('months', selectedMonths.join(','));
    formData.append('randomness', document.getElementById('randomness').value);

    // Show loading spinner
    showElement(document.getElementById('loadingSpinner'));
    hideElement(document.getElementById('errorAlert'));

    try {
        const response = await fetch('/generate_schedule', {
            method: 'POST',
            body: formData
        });

        let data = null;
        try {
            data = await response.json();
        } catch (parseError) {
            console.error('Failed to parse response JSON:', parseError);
        }

        const exportBtn = document.getElementById('exportBtn');

        if (response.ok && data && data.success) {
            generatedSchedule = data.schedule;
            displaySchedule(data.schedule);
            showElement(exportBtn);
            hideElement(errorAlert);
        } else {
            generatedSchedule = null;
            hideElement(exportBtn);

            let message = (data && (data.error || data.message)) || 'Failed to generate schedule. Please verify the file format and selected months, then try again.';

            if (data && data.unscheduled_courses) {
                const monthNames = [
                    '', 'January', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'
                ];

                const details = Object.entries(data.unscheduled_courses)
                    .map(([month, courses]) => {
                        const monthIndex = parseInt(month, 10);
                        const monthLabel = Number.isNaN(monthIndex) ? month : monthNames[monthIndex] || month;
                        return `${monthLabel}: ${courses.join(', ')}`;
                    })
                    .join('\n');

                if (details && !message.includes(details)) {
                    message += `\n${details}`;
                }
            }

            errorAlert.innerHTML = message.replace(/\n/g, '<br>');
            showElement(errorAlert);
        }
    } catch (error) {
        generatedSchedule = null;
        hideElement(document.getElementById('exportBtn'));
        const errorAlert = document.getElementById('errorAlert');
        errorAlert.textContent = 'Error generating schedule: ' + error.message + '. Please check your network and input file, then retry.';
        showElement(errorAlert);
    } finally {
        hideElement(document.getElementById('loadingSpinner'));
    }
});
}


document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('xmlForm');
    const fileInput = document.getElementById('inputFile');
    const fileName = document.getElementById('fileName');
    const loadingSpinner = document.getElementById('loadingSpinner');
    const errorAlert = document.getElementById('errorAlert');
    const successMessage = document.getElementById('successMessage');

    if (!form || !fileInput || !fileName || !loadingSpinner || !errorAlert || !successMessage) {
        return;
    }

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

        showElement(loadingSpinner, 'flex');
        hideElement(errorAlert);
        hideElement(successMessage);

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

                showElement(successMessage);
                successMessage.style.opacity = '1';
                setTimeout(() => {
                    hideElement(successMessage);
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
            hideElement(loadingSpinner);
        }
    });

    function showError(message) {
        console.error('Showing error:', message);
        errorAlert.textContent = message;
        showElement(errorAlert);
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
            showElement(loadingSpinner);

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
                showElement(successMessage);
                setTimeout(() => {
                    hideElement(successMessage);
                }, 5000);
            } catch (error) {
                // Show error message
                errorAlert.textContent = error.message;
                showElement(errorAlert);
                setTimeout(() => {
                    hideElement(errorAlert);
                }, 5000);
            } finally {
                // Hide the loading spinner
                hideElement(loadingSpinner);
            }
        });
    }
});
