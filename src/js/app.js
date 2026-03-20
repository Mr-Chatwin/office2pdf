(function () {
    'use strict';

    if (!window.__TAURI__) {
        console.error('未在 Tauri 环境中运行');
        document.querySelector('.subtitle').textContent = '运行环境错误，请使用 Tauri 启动';
        return;
    }

    const { invoke } = window.__TAURI__.core;

    // DOM Elements
    const setupSection = document.getElementById('setupSection');
    const dropZone = document.getElementById('dropZone');
    const resultsSection = document.getElementById('resultsSection');
    const fileList = document.getElementById('fileList');
    const fileInputLabel = document.getElementById('fileInputLabel');

    // 状态管理
    let conversionQueue = [];
    let isConverting = false;
    let stopRequested = false;

    // --- 初始化 ---
    async function init() {
        const subtitle = document.querySelector('.subtitle');
        try {
            const status = await invoke('check_engines');
            // status -> { available: true, engine: 'office', all_engines: ['office', 'wps'] }
            
            if (status.available) {
                const engineNames = { office: 'Microsoft Office', wps: 'WPS Office', libreoffice: 'LibreOffice' };
                const engineName = engineNames[status.engine] || status.engine;
                showToast(`已检测到 ${engineName} 转换引擎`, 'success');

                if (status.all_engines && status.all_engines.length > 1) {
                    const options = status.all_engines.map(e => 
                        `<option value="${e}" ${e === status.engine ? 'selected' : ''}>${engineNames[e] || e}</option>`
                    ).join('');
                    
                    subtitle.innerHTML = `转换引擎：<select id="engineSelect" class="engine-select">${options}</select> · 高质量 · 极速`;
                    
                    document.getElementById('engineSelect').addEventListener('change', async (e) => {
                        const next = e.target.value;
                        const ok = await invoke('set_engine', { engine: next });
                        if (ok) {
                            showToast(`已切换到 ${engineNames[next] || next} 引擎`, 'info');
                        } else {
                            showToast(`切换引擎失败`, 'error');
                            e.target.value = status.engine;
                        }
                    });
                } else {
                    subtitle.textContent = `使用 ${engineName} 引擎 · 高质量转换 · 极速`;
                }

                dropZone.classList.remove('hidden');
            } else {
                subtitle.textContent = '缺少转换引擎';
                setupSection.classList.remove('hidden');
            }
        } catch (e) {
            subtitle.textContent = '引擎检测失败';
            showToast('引擎检测出错: ' + e, 'error');
            dropZone.classList.remove('hidden');
        }

        bindEvents();
    }

    function bindEvents() {
        // 使用后台 rfd 原生弹窗
        fileInputLabel.addEventListener('click', async () => {
            if (isConverting) return;
            try {
                const paths = await invoke('select_files');
                if (paths && paths.length > 0) {
                    addFilesToQueue(paths);
                }
            } catch (err) {
                console.error('选择文件出错: ', err);
            }
        });

        // 监听 Tauri 提供的拖拽事件
        window.__TAURI__.event.listen('tauri://drop', (event) => {
            if (isConverting) return;
            const paths = event.payload.paths || event.payload; // 兼容不同版本的 payload 结构
            if (!Array.isArray(paths)) return;
            
            const pptPaths = paths.filter(p => /\.(pptx?|ppsx?|pps)$/i.test(p));
            if (pptPaths.length > 0) {
                addFilesToQueue(pptPaths);
            } else {
                showToast('请拖入有效的 PPT 文件', 'error');
            }
            dropZone.classList.remove('drag-over');
        });
        
        window.__TAURI__.event.listen('tauri://drag-enter', () => dropZone.classList.add('drag-over'));
        window.__TAURI__.event.listen('tauri://drag-leave', () => dropZone.classList.remove('drag-over'));

        document.getElementById('startConvertBtn')?.addEventListener('click', startConversion);
        document.getElementById('stopConvertBtn')?.addEventListener('click', () => { stopRequested = true; });
        document.getElementById('newConvertBtn')?.addEventListener('click', clearQueue);
        document.getElementById('addMoreBtn')?.addEventListener('click', () => fileInputLabel.click());
    }

    function addFilesToQueue(paths) {
        dropZone.classList.add('hidden');
        resultsSection.classList.remove('hidden');

        paths.forEach(p => {
            const name = p.split(/[\\/]/).pop();
            conversionQueue.push({ path: p, name, status: 'pending', error: null, result: null });
        });
        renderQueue();
    }

    function clearQueue() {
        if (isConverting) {
            showToast('请先停止当前转换', 'error');
            return;
        }
        conversionQueue = [];
        resultsSection.classList.add('hidden');
        dropZone.classList.remove('hidden');
    }

    // 供 HTML 内联 onClick 调用的全局函数
    window.removeQueueItem = function(index) {
        if (isConverting) return;
        conversionQueue.splice(index, 1);
        if (conversionQueue.length === 0) clearQueue();
        else renderQueue();
    };

    window.openFile = async function(path) {
        try {
            await invoke('open_file', { path });
        } catch (e) {
            showToast('打开文件失败', 'error');
        }
    };

    window.openFolder = async function(path) {
        try {
            await invoke('open_folder', { path });
        } catch (e) {
            showToast('打开文件夹失败', 'error');
        }
    };

    // --- 界面渲染逻辑 ---
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
                    <button class="btn btn-sm btn-ghost" onclick="window.openFile('${escapeJs(item.result)}')">打开 PDF</button>
                    <button class="btn btn-sm btn-ghost" onclick="window.openFolder('${escapeJs(item.result)}')">打开文件夹</button>
                `;
            } else if (item.status === 'pending') {
                actionsHtml = `<button class="btn btn-sm btn-ghost btn-danger" onclick="removeQueueItem(${i})">移除</button>`;
            }

            div.innerHTML = `
                <div class="file-item-icon">📊</div>
                <div class="file-item-info">
                    <div class="file-item-name">${escapeHtml(item.name)}</div>
                    <div class="file-item-status ${statusClass}">${statusHtml}</div>
                </div>
                <div class="file-item-actions">${actionsHtml}</div>
            `;
            fileList.appendChild(div);
        });

        updateConvertButtons();
    }

    function updateConvertButtons() {
        const startConvertBtn = document.getElementById('startConvertBtn');
        const stopConvertBtn = document.getElementById('stopConvertBtn');
        const newConvertBtn = document.getElementById('newConvertBtn');
        const addMoreBtn = document.getElementById('addMoreBtn');
        
        const hasPending = conversionQueue.some(i => i.status === 'pending');
        
        if (isConverting) {
            startConvertBtn.classList.add('hidden');
            stopConvertBtn.classList.remove('hidden');
            newConvertBtn.disabled = true;
            addMoreBtn.disabled = true;
        } else {
            stopConvertBtn.classList.add('hidden');
            startConvertBtn.classList.remove('hidden');
            newConvertBtn.disabled = false;
            addMoreBtn.disabled = false;
            
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

    async function startConversion() {
        if (isConverting) return;
        
        const pendingItems = conversionQueue.filter(i => i.status === 'pending');
        if (pendingItems.length === 0) return;

        isConverting = true;
        stopRequested = false;
        
        const engineSelect = document.getElementById('engineSelect');
        if (engineSelect) engineSelect.disabled = true;
        
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
                // invoke returns the path to the PDF on success
                const pdfPath = await invoke('convert_pptx', { path: item.path });
                item.status = 'success';
                item.result = pdfPath;
                successCount++;
            } catch (err) {
                item.status = 'error';
                item.error = typeof err === 'string' ? err : err.message || JSON.stringify(err);
                failCount++;
            }
            renderQueue();
        }

        isConverting = false;
        if (engineSelect) engineSelect.disabled = false;
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

    // --- 工具函数 ---
    function showToast(msg, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = msg;
        document.body.appendChild(toast);
        
        toast.offsetHeight; // reflow
        toast.classList.add('show');
        
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 400);
        }, 3000);
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str).replace(/[&<>'"]/g, tag => ({
            '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
        }[tag] || tag));
    }
    
    function escapeJs(str) {
        if (!str) return '';
        return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '\\"');
    }

    init();
})();
