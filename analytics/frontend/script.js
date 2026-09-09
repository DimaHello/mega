document.getElementById('report-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const form = e.target;
    const btn = document.getElementById('generate-btn');
    const errorDiv = document.getElementById('error-message');
    
    const pType = form.periodType.value;
    const reportType = form.reportType.value;
    let dateVal;
    
    if (pType === 'custom') {
        const start = document.getElementById('dateStart').value;
        const end = document.getElementById('dateEnd').value;
        if (!start || !end) {
            alert("Выберите обе даты!");
            return;
        }
        dateVal = `${start}_to_${end}`;
    } else {
        dateVal = form.dateValue.value;
    }
    
    const parkCheckboxes = document.querySelectorAll('input[name="park"]');
    const selectedParks = Array.from(parkCheckboxes)
        .filter(cb => cb.checked)
        .map(cb => cb.value);
        
    if (selectedParks.length === 0) {
        alert("Выберите хотя бы один парк!");
        return;
    }
    
    // UI Loading state
    btn.classList.add('loading');
    btn.disabled = true;
    errorDiv.classList.add('hidden');
    errorDiv.textContent = '';
    
    try {
        const apiUrl = window.location.pathname.replace(/\/$/, '') + '/api/generate';
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ 
                dateValue: dateVal, 
                periodType: pType,
                parks: selectedParks,
                type: reportType 
            })
        });
        
        if (!response.ok) {
            let errorText = 'Ошибка при генерации отчета';
            try {
                const errJson = await response.json();
                errorText = errJson.error || errorText;
            } catch (e) {}
            throw new Error(errorText);
        }
        
        // Получаем файл
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        
        // Получаем имя файла из заголовков (если есть) или генерируем свое
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = `${reportType}_${dateVal}.xlsx`;
        if (contentDisposition && contentDisposition.includes('filename=')) {
            filename = contentDisposition.split('filename=')[1].replace(/"/g, '');
        }
        
        // Триггер скачивания
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        
        window.URL.revokeObjectURL(url);
        a.remove();
        
    } catch (error) {
        errorDiv.textContent = `❌ ${error.message}`;
        errorDiv.classList.remove('hidden');
    } finally {
        btn.classList.remove('loading');
        btn.disabled = false;
    }
});

const periodType = document.getElementById('periodType');
const dateInputContainer = document.getElementById('dateInputContainer');

if (periodType && dateInputContainer) {
    periodType.addEventListener('change', (e) => {
        const val = e.target.value;
        if (val === 'day') {
            dateInputContainer.innerHTML = `
                <label for="dateValue">Выбор даты</label>
                <input type="date" id="dateValue" name="dateValue" class="form-control" required value="2026-03-01">
            `;
        } else if (val === 'month') {
            dateInputContainer.innerHTML = `
                <label for="dateValue">Выбор даты</label>
                <input type="month" id="dateValue" name="dateValue" class="form-control" required value="2026-03">
            `;
        } else if (val === 'year') {
            dateInputContainer.innerHTML = `
                <label for="dateValue">Выбор даты</label>
                <input type="number" id="dateValue" name="dateValue" class="form-control" required placeholder="YYYY" value="2026" min="2020" max="2030">
            `;
        } else if (val === 'custom') {
            dateInputContainer.innerHTML = `
                <label>Интервал дат</label>
                <div style="display: flex; gap: 10px;">
                    <input type="date" id="dateStart" class="form-control" required value="2026-03-01" style="flex: 1;">
                    <span style="align-self: center; color: var(--text-muted);">по</span>
                    <input type="date" id="dateEnd" class="form-control" required value="2026-03-15" style="flex: 1;">
                </div>
            `;
        }
    });
}

const selectAll = document.getElementById('selectAllParks');
const deselectAll = document.getElementById('deselectAllParks');

if (selectAll && deselectAll) {
    selectAll.addEventListener('click', () => {
        document.querySelectorAll('input[name="park"]').forEach(cb => cb.checked = true);
    });
    deselectAll.addEventListener('click', () => {
        document.querySelectorAll('input[name="park"]').forEach(cb => cb.checked = false);
    });
}

