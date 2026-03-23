let holidays = [];
let generatedSchedule = null;

function selectAllMonths() { document.querySelectorAll('.month-checkbox').forEach(c => c.checked = true); }
function deselectAllMonths() { document.querySelectorAll('.month-checkbox').forEach(c => c.checked = false); }
function removeHoliday(index) { holidays.splice(index, 1); updateHolidayDisplay(); }

function updateHolidayDisplay() {
    const container = document.getElementById('holidayList');
    if (!container) return;
    container.innerHTML = holidays.map((date, index) => `<span class="badge badge-primary gap-1">${date}<button class="btn btn-xs" type="button" onclick="removeHoliday(${index})">×</button></span>`).join('');
}

function isValidDate(dateStr) {
    const regex = /^\d{2}\.\d{2}\.\d{4}$/;
    if (!regex.test(dateStr)) return false;
    const [d, m, y] = dateStr.split('.').map(Number);
    const dt = new Date(y, m - 1, d);
    return dt.getDate() === d && dt.getMonth() === m - 1 && dt.getFullYear() === y;
}

function addHolidays() {
    const input = document.getElementById('holidayInput');
    if (!input) return;
    input.value.split(/[\n,]/).map(d => d.trim()).filter(Boolean).forEach(date => {
        if (isValidDate(date) && !holidays.includes(date)) holidays.push(date);
    });
    input.value = '';
    updateHolidayDisplay();
}

function displaySchedule(schedule) {
    const tableBody = document.getElementById('scheduleTableBody');
    const courseCounter = document.getElementById('courseCounter');
    if (!tableBody || !courseCounter) return;
    tableBody.innerHTML = '';

    const sorted = [...schedule].sort((a, b) => a.original_order - b.original_order);
    const map = new Map();
    sorted.forEach(item => {
        if (!map.has(item.Title)) map.set(item.Title, {name: item.Title, duration: item['Durata Curs'], investitie: item.investitie || item.Investitie || '', months: {}});
        map.get(item.Title).months[item.month] = item.date_range;
    });

    const formatCellValue = (value) => String(value || '').replace(/\s*\n\s*/g, '').trim();

    map.forEach(course => {
        const row = document.createElement('tr');
        const monthsCells = Array.from({length: 12}, (_, i) => `<td class="text-sm">${formatCellValue(course.months[i + 1])}</td>`).join('');
        row.innerHTML = `<td class="course-title">${course.name}</td><td>${formatCellValue(course.duration)}</td><td>${formatCellValue(course.investitie)}</td>${monthsCells}`;
        tableBody.appendChild(row);
    });

    courseCounter.innerHTML = `<strong>Course Summary:</strong><br>Total unique courses loaded: ${map.size}<br>Total scheduled sessions: ${schedule.length}`;
    courseCounter.classList.remove('hidden');
    document.getElementById('scheduleResult')?.classList.remove('hidden');
}

async function exportSchedule() {
    const response = await fetch('/export_schedule', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({schedule: generatedSchedule, year: document.getElementById('year')?.value, holidays})
    });
    if (!response.ok) throw new Error('Export failed');
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `course_schedule_${document.getElementById('year')?.value || new Date().getFullYear()}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
}

document.addEventListener('DOMContentLoaded', () => {
    const randomness = document.getElementById('randomness');
    randomness?.addEventListener('input', () => { document.getElementById('randomnessValue').textContent = randomness.value; });

    const scheduleForm = document.getElementById('scheduleForm');
    scheduleForm?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const months = Array.from(document.querySelectorAll('.month-checkbox:checked')).map(c => c.value);
        if (months.length === 0) { alert('Please select at least one month'); return; }
        const file = document.getElementById('inputFile').files[0];
        if (!file) { alert('Please select a file'); return; }

        const fd = new FormData();
        fd.append('input_file', file);
        fd.append('year', document.getElementById('year').value);
        fd.append('holidays', holidays.join(','));
        fd.append('months', months.join(','));
        fd.append('randomness', randomness?.value || '5');

        const spinner = document.getElementById('loadingSpinner');
        const errorAlert = document.getElementById('errorAlert');
        spinner?.classList.remove('hidden');
        errorAlert?.classList.add('hidden');

        try {
            const response = await fetch('/generate_schedule', {method: 'POST', body: fd});
            const data = await response.json();
            if (response.ok && data.success) {
                generatedSchedule = data.schedule;
                displaySchedule(data.schedule);
                document.getElementById('exportBtn')?.classList.remove('hidden');
            } else {
                throw new Error(data.error || 'Failed to generate schedule');
            }
        } catch (error) {
            errorAlert.textContent = error.message;
            errorAlert.classList.remove('hidden');
        } finally {
            spinner?.classList.add('hidden');
        }
    });

    const xmlForm = document.getElementById('xmlForm');
    xmlForm?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const file = document.getElementById('xmlInputFile').files[0];
        if (!file) return;
        const fd = new FormData(); fd.append('input_file', file);
        const loading = document.getElementById('xmlLoading');
        const error = document.getElementById('xmlError');
        const success = document.getElementById('xmlSuccess');
        loading.classList.remove('hidden'); error.classList.add('hidden'); success.classList.add('hidden');
        try {
            const response = await fetch('/format-xml', {method: 'POST', body: fd});
            if (!response.ok) {
                const payload = await response.json();
                throw new Error(payload.error || 'Failed to generate XML');
            }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href = url; a.download = 'formatted_courses.xml'; a.click();
            URL.revokeObjectURL(url);
            success.classList.remove('hidden');
        } catch (err) {
            error.textContent = err.message;
            error.classList.remove('hidden');
        } finally {
            loading.classList.add('hidden');
        }
    });

    const wordForm = document.getElementById('convertWordForm');
    wordForm?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const wordFile = document.getElementById('wordFile').files[0];
        const permalinksFile = document.getElementById('permalinksFile').files[0];
        const loading = document.getElementById('wordLoading');
        const error = document.getElementById('wordError');
        const success = document.getElementById('wordSuccess');
        if (!wordFile || !permalinksFile) return;
        const fd = new FormData();
        fd.append('word_file', wordFile);
        fd.append('permalinks_file', permalinksFile);
        loading.classList.remove('hidden'); error.classList.add('hidden'); success.classList.add('hidden');
        try {
            const response = await fetch('/convert_word', {method: 'POST', body: fd});
            if (!response.ok) {
                const payload = await response.json();
                throw new Error(payload.error || 'Failed to convert file');
            }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href = url; a.download = 'matched_courses.docx'; a.click();
            URL.revokeObjectURL(url);
            success.classList.remove('hidden');
        } catch (err) {
            error.textContent = err.message;
            error.classList.remove('hidden');
        } finally {
            loading.classList.add('hidden');
        }
    });
});

window.selectAllMonths = selectAllMonths;
window.deselectAllMonths = deselectAllMonths;
window.addHolidays = addHolidays;
window.exportSchedule = exportSchedule;
window.removeHoliday = removeHoliday;
