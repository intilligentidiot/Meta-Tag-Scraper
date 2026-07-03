let currentResults = [];

function getSeoPill(length, lo, hi) {
    if (length === 0) return `<span class="pill pill-bad">Missing</span>`;
    if (length < lo) return `<span class="pill pill-warn">${length} · Too Short</span>`;
    if (length > hi) return `<span class="pill pill-warn">${length} · Too Long</span>`;
    return `<span class="pill pill-ok">${length} · OK</span>`;
}

document.getElementById('scan-btn').addEventListener('click', async () => {
    const rawInput = document.getElementById('url-input').value;
    const urls = rawInput.split('\n').map(u => u.trim().replace(',', '')).filter(u => u);
    
    if (urls.length === 0) return;

    const btn = document.getElementById('scan-btn');
    const statusText = document.getElementById('status-text');
    
    btn.disabled = true;
    btn.textContent = "Scanning...";
    statusText.classList.remove('hidden');
    statusText.textContent = `Scanning ${urls.length} URL(s)...`;
    
    document.getElementById('results-container').classList.add('hidden');

    try {
        const response = await fetch('/api/index', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'scrape', urls })
        });
        
        const data = await response.json();
        currentResults = data.results || [];
        renderResults();
    } catch (error) {
        console.error("Error scraping:", error);
        statusText.textContent = "Error occurred while scraping.";
    } finally {
        btn.disabled = false;
        btn.textContent = "⚡ Scan All";
    }
});

document.getElementById('clear-btn').addEventListener('click', () => {
    document.getElementById('url-input').value = "";
    document.getElementById('results-container').classList.add('hidden');
    document.getElementById('status-text').classList.add('hidden');
    currentResults = [];
});

document.getElementById('export-btn').addEventListener('click', async () => {
    if (currentResults.length === 0) return;
    
    try {
        const response = await fetch('/api/index', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'export', results: currentResults })
        });
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'meta_data.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
    } catch (error) {
        console.error("Error exporting:", error);
    }
});

function renderResults() {
    const container = document.getElementById('results-container');
    const list = document.getElementById('results-list');
    const title = document.getElementById('results-title');
    
    list.innerHTML = "";
    title.textContent = `Results — ${currentResults.length} page${currentResults.length !== 1 ? 's' : ''}`;
    
    currentResults.forEach(r => {
        const tl = r.meta_title ? r.meta_title.length : 0;
        const dl = r.meta_description ? r.meta_description.length : 0;
        
        const cardClass = r.error ? "result-card has-error" : "result-card";
        const dotClass = r.error ? "dot dot-err" : "dot dot-ok";
        
        let html = `
            <div class="${cardClass}">
                <div class="result-url"><span class="${dotClass}"></span>${r.url}</div>
        `;
        
        if (r.error) {
            html += `
                <div class="meta-label">Error</div>
                <div class="meta-value err">${r.error}</div>
            `;
        } else {
            const titleDisplay = r.meta_title || "No title tag found";
            const descDisplay = r.meta_description || "No meta description found";
            const titleEmpty = r.meta_title ? "" : "empty";
            const descEmpty = r.meta_description ? "" : "empty";
            
            html += `
                <div class="meta-label">Meta Title ${getSeoPill(tl, 30, 60)}</div>
                <div class="meta-value ${titleEmpty}">${titleDisplay}</div>
                
                <div class="meta-label">Meta Description ${getSeoPill(dl, 70, 160)}</div>
                <div class="meta-value ${descEmpty}">${descDisplay}</div>
            `;
        }
        
        html += `</div>`;
        list.insertAdjacentHTML('beforeend', html);
    });
    
    document.getElementById('status-text').textContent = `${currentResults.length} URL(s) processed`;
    container.classList.remove('hidden');
}
