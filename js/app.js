/**
 * 主应用逻辑 — 支持 Electron 模式和 Web 模式
 */
(function () {
    'use strict';

    const isElectron = !!window.electronAPI;

    // ===== DOM 元素 =====
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const progressSection = document.getElementById('progressSection');
    const progressText = document.getElementById('progressText');
    const progressPercent = document.getElementById('progressPercent');
    const progressFill = document.getElementById('progressFill');
    const previewSection = document.getElementById('previewSection');
    const fileName = document.getElementById('fileName');
    const slideCount = document.getElementById('slideCount');
    const slidePreview = document.getElementById('slidePreview');
    const exportBtn = document.getElementById('exportBtn');
    const resetBtn = document.getElementById('resetBtn');
    const renderArea = document.getElementById('renderArea');
    const setupSection = document.getElementById('setupSection');
    const resultsSection = document.getElementById('resultsSection');
    const fileList = document.getElementById('fileList');
    const newConvertBtn = document.getElementById('newConvertBtn');

    // Web 模式状态
    let currentSlideElements = [];
    let currentSlideSize = null;
    let currentFileName = '';

    // Electron 模式状态
    let conversionQueue = [];
    let isConverting = false;
    let stopRequested = false;

    // ===== 初始化 =====

    if (isElectron) {
        initElectronMode();
    } else {
        initWebMode();
    }

    // 通用事件绑定
    dropZone.addEventListener('click', handleDropZoneClick);
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', (e) => { e.preventDefault(); e.stopPropagation(); dropZone.classList.remove('drag-over'); });
    dropZone.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileInputChange);

    // ===== Electron 模式 =====

    async function initElectronMode() {
        // 检查可用转换引擎
        const status = await window.electronAPI.checkLibreOffice();

        if (status.available) {
            const engineNames = { office: 'Microsoft Office', wps: 'WPS Office', libreoffice: 'LibreOffice' };
            const engineName = engineNames[status.engine] || status.engine;
            showToast(`已检测到 ${engineName} 转换引擎`, 'success');

            // 在副标题显示引擎信息（下拉框选择）
            const subtitle = document.querySelector('.subtitle');
            if (subtitle && status.allEngines && status.allEngines.length > 1) {
                const options = status.allEngines.map(e => 
                    `<option value="${e}" ${e === status.engine ? 'selected' : ''}>${engineNames[e] || e}</option>`
                ).join('');
                
                subtitle.innerHTML = `转换引擎：<select id="engineSelect" class="engine-select">${options}</select> · 隐私安全`;
                
                document.getElementById('engineSelect').addEventListener('change', async (e) => {
                    const next = e.target.value;
                    const ok = await window.electronAPI.setEngine(next);
                    if (ok) {
                        showToast(`已切换到 ${engineNames[next] || next} 引擎`, 'info');
                    } else {
                        showToast(`切换引擎失败`, 'error');
                        e.target.value = status.engine; // 复原
                    }
                });
            } else if (subtitle) {
                subtitle.textContent = `使用 ${engineName} 引擎 · 高质量转换 · 隐私安全`;
            }
        } else {
            showSetup();
        }

        // 下载按钮
        const downloadBtn = document.getElementById('downloadLoBtn');
        downloadBtn?.addEventListener('click', handleDownloadLO);

        // 下载进度
        window.electronAPI.onDownloadProgress((data) => {
            const loProgressEl = document.getElementById('loProgress');
            const loProgressText = document.getElementById('loProgressText');
            const loProgressPercent = document.getElementById('loProgressPercent');
            const loProgressFill = document.getElementById('loProgressFill');

            loProgressEl.classList.remove('hidden');
            loProgressText.textContent = data.text;
            loProgressPercent.textContent = data.percent + '%';
            loProgressFill.style.width = data.percent + '%';

            if (data.stage === 'done') {
                setTimeout(() => {
                    setupSection.classList.add('hidden');
                    dropZone.classList.remove('hidden');
                    showToast('转换引擎已就绪！', 'success');
                }, 1000);
            }
        });

        // 这里需要绑定新的排队和停止按钮
        const startConvertBtn = document.getElementById('startConvertBtn');
        const stopConvertBtn = document.getElementById('stopConvertBtn');

        startConvertBtn?.addEventListener('click', startConversionQueue);
        stopConvertBtn?.addEventListener('click', () => { stopRequested = true; });
        newConvertBtn?.addEventListener('click', handleResetQueue);
    }

    async function handleDownloadLO() {
        const btn = document.getElementById('downloadLoBtn');
        btn.disabled = true;
        btn.textContent = '正在准备下载...';

        try {
            await window.electronAPI.downloadLibreOffice();
        } catch (err) {
            showToast('下载失败: ' + err.message, 'error');
            btn.disabled = false;
            btn.innerHTML = '重新下载';
        }
    }

    function showSetup() {
        dropZone.classList.add('hidden');
        setupSection.classList.remove('hidden');
    }

    function handleResetQueue() {
        if (isConverting) {
            showToast('请先停止当前转换', 'error');
            return;
        }
        conversionQueue = [];
        handleReset();
    }

    // ===== Web 模式 =====

    function initWebMode() {
        exportBtn?.addEventListener('click', handleExport);
        resetBtn?.addEventListener('click', handleReset);
    }

    // ===== 通用事件处理 =====

    function handleDropZoneClick(e) {
        if (e.target.tagName === 'LABEL') return; // 让 label 自己触发 file input
        fileInput.click();
    }

    function handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('drag-over');

        const files = Array.from(e.dataTransfer.files).filter(f =>
            /\.(pptx?|ppsx?|pps)$/i.test(f.name)
        );
        if (files.length === 0) {
            showToast('请拖入 PPT 文件', 'error');
            return;
        }

        if (isElectron) {
            handleFilesElectron(files);
        } else {
            handleFileWeb(files[0]);
        }
    }

    function handleFileInputChange(e) {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        if (isElectron) {
            handleFilesElectron(files);
        } else {
            handleFileWeb(files[0]);
        }
    }

    async function handleFilesElectron(files) {
        dropZone.classList.add('hidden');
        resultsSection.classList.remove('hidden');

        // 把文件加入队列
        files.forEach((file) => {
            conversionQueue.push({ file, status: 'pending', error: null, result: null });
        });

        renderQueue();
    }

    function renderQueue() {
        const resultsTitle = document.getElementById('resultsTitle');
        resultsTitle.textContent = `待转换 ${conversionQueue.length} 个文件`;
        fileList.innerHTML = '';

        conversionQueue.forEach((item, i) => {
            const div = document.createElement('div');
            div.className = 'file-item';
            
            let statusHtml = '<span style="color:var(--text-secondary)">⏳ 等待中</span>';
            let statusClass = '';
            
            if (item.status === 'converting') {
                statusHtml = '<span style="color:var(--accent-1)">🔄 转换中...</span>';
                statusClass = 'converting';
            } else if (item.status === 'success') {
                statusHtml = '<span style="color:var(--success)">✅ 完成</span>';
                statusClass = 'success';
            } else if (item.status === 'error') {
                statusHtml = `<span style="color:var(--error)">❌ 失败: ${item.error}</span>`;
                statusClass = 'error';
            } else if (item.status === 'stopped') {
                statusHtml = '<span style="color:var(--text-secondary)">⏹️ 已取消</span>';
            }

            let actionsHtml = '';
            if (item.status === 'success' && item.result) {
                actionsHtml = `
                    <button class="btn btn-sm btn-ghost" onclick="window.electronAPI.openFile('${escapeJs(item.result.pdfPath)}')">打开 PDF</button>
                    <button class="btn btn-sm btn-ghost" onclick="window.electronAPI.openFolder('${escapeJs(item.result.pdfPath)}')">打开文件夹</button>
                `;
            } else if (item.status === 'pending') {
                actionsHtml = `<button class="btn btn-sm btn-ghost btn-danger" onclick="removeQueueItem(${i})">移除</button>`;
            }

            div.innerHTML = `
                <div class="file-item-icon">📊</div>
                <div class="file-item-info">
                    <div class="file-item-name">${escapeHtml(item.file.name)}</div>
                    <div class="file-item-status ${statusClass}">${statusHtml}</div>
                </div>
                <div class="file-item-actions">${actionsHtml}</div>
            `;
            fileList.appendChild(div);
        });

        updateConvertButtons();
    }

    window.removeQueueItem = function(index) {
        if (isConverting) return;
        conversionQueue.splice(index, 1);
        if (conversionQueue.length === 0) {
            handleResetQueue();
        } else {
            renderQueue();
        }
    };

    function updateConvertButtons() {
        const startConvertBtn = document.getElementById('startConvertBtn');
        const stopConvertBtn = document.getElementById('stopConvertBtn');
        const newConvertBtn = document.getElementById('newConvertBtn');
        
        const hasPending = conversionQueue.some(i => i.status === 'pending');
        
        if (isConverting) {
            startConvertBtn.classList.add('hidden');
            stopConvertBtn.classList.remove('hidden');
            newConvertBtn.disabled = true;
        } else {
            stopConvertBtn.classList.add('hidden');
            startConvertBtn.classList.remove('hidden');
            newConvertBtn.disabled = false;
            
            if (!hasPending && conversionQueue.length > 0) {
                startConvertBtn.disabled = true;
                startConvertBtn.textContent = '转换已完成';
            } else {
                startConvertBtn.disabled = conversionQueue.length === 0;
                startConvertBtn.innerHTML = `
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>
                    开始转换
                `;
            }
        }
    }

    async function startConversionQueue() {
        if (isConverting) return;
        
        const pendingItems = conversionQueue.filter(i => i.status === 'pending');
        if (pendingItems.length === 0) return;

        isConverting = true;
        stopRequested = false;
        document.getElementById('engineSelect').disabled = true; // 转换期间禁止切换引擎
        updateConvertButtons();

        let successCount = 0;
        let failCount = 0;

        for (let i = 0; i < conversionQueue.length; i++) {
            if (stopRequested) {
                conversionQueue.filter(item => item.status === 'pending').forEach(item => item.status = 'stopped');
                break;
            }

            const item = conversionQueue[i];
            if (item.status !== 'pending') continue;

            item.status = 'converting';
            renderQueue();

            try {
                const result = await window.electronAPI.convertPptx(item.file.path);
                item.status = 'success';
                item.result = result;
                successCount++;
            } catch (err) {
                item.status = 'error';
                item.error = err.message;
                failCount++;
            }
            renderQueue();
        }

        isConverting = false;
        document.getElementById('engineSelect').disabled = false;
        updateConvertButtons();

        const resultsTitle = document.getElementById('resultsTitle');
        if (stopRequested) {
            resultsTitle.textContent = `转换已终止 (成功: ${successCount}, 失败: ${failCount})`;
            showToast('转换已被用户终止', 'info');
        } else {
            resultsTitle.textContent = `转换完成 — ${successCount} 个成功${failCount > 0 ? `，${failCount} 个失败` : ''}`;
            showToast(`完成！${successCount} 个文件已转换`, successCount > 0 ? 'success' : 'error');
        }
    }

    // ===== Web 模式：转换流程 =====

    async function handleFileWeb(file) {
        if (!file.name.toLowerCase().endsWith('.pptx')) {
            showToast('Web 模式仅支持 .pptx 格式', 'error');
            return;
        }

        currentFileName = file.name;
        showProgress();

        try {
            const parser = new PptxParser();
            const result = await parser.parse(file, updateProgress);
            currentSlideSize = result.slideSize;

            updateProgress('正在渲染幻灯片...', 85);
            const renderer = new SlideRenderer();
            currentSlideElements = [];
            renderArea.innerHTML = '';

            for (const slide of result.slides) {
                const el = renderer.renderSlide(slide, result.slideSize);
                renderArea.appendChild(el);
                currentSlideElements.push(el);
            }

            updateProgress('正在生成预览...', 90);
            await generatePreviews();

            updateProgress('完成！', 100);
            setTimeout(() => {
                showPreview(file.name, result.slides.length);
                showToast(`已解析 ${result.slides.length} 页幻灯片`, 'success');
            }, 400);
        } catch (err) {
            console.error('解析失败:', err);
            showToast('文件解析失败: ' + err.message, 'error');
            handleReset();
        }
    }

    async function generatePreviews() {
        slidePreview.innerHTML = '';
        for (let i = 0; i < currentSlideElements.length; i++) {
            const thumb = document.createElement('div');
            thumb.className = 'slide-thumb';
            try {
                const canvas = await html2canvas(currentSlideElements[i], {
                    scale: 0.5, useCORS: true, allowTaint: true,
                    backgroundColor: '#FFFFFF', width: currentSlideSize.width,
                    height: currentSlideSize.height, logging: false,
                });
                const img = document.createElement('img');
                img.src = canvas.toDataURL('image/jpeg', 0.8);
                img.className = 'slide-thumb-img';
                thumb.appendChild(img);
            } catch (err) {
                const placeholder = document.createElement('div');
                placeholder.className = 'slide-thumb-img';
                placeholder.style.cssText = 'display:flex;align-items:center;justify-content:center;color:#999;font-size:14px;';
                placeholder.textContent = '预览失败';
                thumb.appendChild(placeholder);
            }
            const label = document.createElement('div');
            label.className = 'slide-thumb-label';
            label.textContent = `第 ${i + 1} 页`;
            thumb.appendChild(label);
            slidePreview.appendChild(thumb);
        }
    }

    async function handleExport() {
        if (currentSlideElements.length === 0) return;
        exportBtn.disabled = true;
        const overlay = document.createElement('div');
        overlay.className = 'export-overlay';
        overlay.innerHTML = '<div class="spinner"></div><p id="exportProgress">正在准备导出...</p>';
        document.body.appendChild(overlay);
        try {
            const exporter = new PdfExporter();
            const blob = await exporter.export(currentSlideElements, currentSlideSize, (text) => {
                const el = document.getElementById('exportProgress');
                if (el) el.textContent = text;
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = currentFileName.replace(/\.pptx?$/i, '') + '.pdf';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('PDF 导出成功！', 'success');
        } catch (err) {
            showToast('PDF 导出失败: ' + err.message, 'error');
        } finally {
            overlay.remove();
            exportBtn.disabled = false;
        }
    }

    // ===== 通用 UI =====

    function handleReset() {
        currentSlideElements = [];
        currentSlideSize = null;
        currentFileName = '';
        if (renderArea) renderArea.innerHTML = '';
        if (slidePreview) slidePreview.innerHTML = '';
        fileInput.value = '';
        fileList.innerHTML = '';

        previewSection.classList.add('hidden');
        progressSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        setupSection.classList.add('hidden');
        dropZone.classList.remove('hidden');
    }

    function showProgress() {
        dropZone.classList.add('hidden');
        previewSection.classList.add('hidden');
        progressSection.classList.remove('hidden');
        updateProgress('正在准备...', 0);
    }

    function updateProgress(text, percent) {
        progressText.textContent = text;
        progressPercent.textContent = Math.round(percent) + '%';
        progressFill.style.width = percent + '%';
    }

    function showPreview(name, count) {
        progressSection.classList.add('hidden');
        previewSection.classList.remove('hidden');
        fileName.textContent = name;
        slideCount.textContent = `${count} 页幻灯片`;
    }

    function showToast(message, type = 'info') {
        document.querySelectorAll('.toast').forEach(t => t.remove());
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;
        document.body.appendChild(toast);
        requestAnimationFrame(() => requestAnimationFrame(() => toast.classList.add('show')));
        setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 400); }, 3000);
    }

    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    function escapeJs(str) {
        return str.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    }

})();
