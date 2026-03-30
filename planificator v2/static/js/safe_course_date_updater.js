(function () {
    const form = document.getElementById('safeUpdaterForm');
    if (!form) {
        return;
    }

    const previewBtn = document.getElementById('previewBtn');
    const errorBox = document.getElementById('previewError');
    const previewContainer = document.getElementById('previewContainer');
    const tableBody = document.getElementById('previewTableBody');

    function escapeHtml(value) {
        return String(value || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    function formatList(values) {
        if (!values || values.length === 0) {
            return '<span class="text-muted">-</span>';
        }
        return values.map(v => `<div>${escapeHtml(v)}</div>`).join('');
    }

    function renderRows(rows) {
        tableBody.innerHTML = '';

        rows.forEach((row) => {
            const tr = document.createElement('tr');
            tr.dataset.rowIndex = String(row.row_index);

            const statusText = row.error ? `error: ${row.error}` : row.status;
            const disabled = row.can_update ? '' : 'disabled';

            tr.innerHTML = `
                <td>${escapeHtml(row.title)}</td>
                <td><small>${escapeHtml(row.permalink)}</small></td>
                <td><code>${escapeHtml(row.slug)}</code></td>
                <td>${row.post_id ?? '-'}</td>
                <td>${formatList(row.existing_valid_dates)}</td>
                <td>${formatList(row.excel_dates)}</td>
                <td>
                    ${formatList(row.final_dates)}
                    <details class="mt-2">
                        <summary>Show payload</summary>
                        <pre class="small mb-0">${escapeHtml(JSON.stringify(row.payload, null, 2))}</pre>
                    </details>
                </td>
                <td class="status-cell">${escapeHtml(statusText)}</td>
                <td>
                    <button class="btn btn-sm btn-success update-row-btn" ${disabled}>Update</button>
                </td>
            `;

            const updateButton = tr.querySelector('.update-row-btn');
            updateButton.addEventListener('click', async () => {
                updateButton.disabled = true;
                const statusCell = tr.querySelector('.status-cell');
                statusCell.textContent = 'updating...';

                try {
                    const payload = {
                        wp_base_url: document.getElementById('wpBaseUrl').value.trim(),
                        wp_username: document.getElementById('wpUsername').value.trim(),
                        wp_app_password: document.getElementById('wpAppPassword').value.trim(),
                        post_id: row.post_id,
                        permalink: row.permalink,
                        slug: row.slug,
                        final_dates: row.final_dates || []
                    };

                    const response = await fetch('/api/safe-course-date-updater/update-row', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    });

                    const data = await response.json();
                    if (!response.ok || !data.success) {
                        throw new Error(data.error || 'Update failed');
                    }

                    statusCell.textContent = data.status || 'success';
                    if (data.status === 'success' || data.status === 'no changes') {
                        updateButton.disabled = true;
                    }
                } catch (error) {
                    statusCell.textContent = `error: ${error.message}`;
                    updateButton.disabled = false;
                }
            });

            tableBody.appendChild(tr);
        });

        previewContainer.classList.remove('d-none');
    }

    form.addEventListener('submit', async (event) => {
        event.preventDefault();
        errorBox.classList.add('d-none');
        previewBtn.disabled = true;
        tableBody.innerHTML = '';

        try {
            const fileInput = document.getElementById('excelFile');
            if (!fileInput.files || !fileInput.files[0]) {
                throw new Error('Please select an Excel file.');
            }

            const formData = new FormData();
            formData.append('input_file', fileInput.files[0]);
            formData.append('wp_base_url', document.getElementById('wpBaseUrl').value.trim());
            formData.append('wp_username', document.getElementById('wpUsername').value.trim());
            formData.append('wp_app_password', document.getElementById('wpAppPassword').value.trim());

            const response = await fetch('/api/safe-course-date-updater/preview', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();
            if (!response.ok || !data.success) {
                throw new Error(data.error || 'Unable to build preview.');
            }

            renderRows(data.rows || []);
        } catch (error) {
            errorBox.textContent = error.message;
            errorBox.classList.remove('d-none');
        } finally {
            previewBtn.disabled = false;
        }
    });
})();
