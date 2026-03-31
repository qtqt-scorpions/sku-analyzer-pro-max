document.addEventListener('DOMContentLoaded', () => {
    // --- Shared UI Elements ---
    const toast = document.getElementById('toast');
    const navTabs = document.querySelectorAll('.nav-tab');
    const tabContents = document.querySelectorAll('.tab-content');

    // --- SKU Merger Elements ---
    const skuInput = document.getElementById('skuInput');
    const processBtn = document.getElementById('processBtn');
    const clearBtn = document.getElementById('clearBtn');
    const resultsContainer = document.getElementById('resultsContainer');
    const statsContainer = document.getElementById('stats');
    const downloadBtn = document.getElementById('downloadBtn');

    // --- Status Analyzer Elements ---
    const masterFileInput = document.getElementById('masterFileInput');
    const fileLabel = document.getElementById('fileLabel');
    const fileInfoCard = document.getElementById('fileInfoCard');
    const displayFileName = document.getElementById('displayFileName');
    const displayFileMeta = document.getElementById('displayFileMeta');
    const changeFileBtn = document.getElementById('changeFileBtn');
    const masterDataModeFileBtn = document.getElementById('masterDataModeFileBtn');
    const masterDataModePasteBtn = document.getElementById('masterDataModePasteBtn');
    const masterDataFileArea = document.getElementById('masterDataFileArea');
    const masterDataPasteArea = document.getElementById('masterDataPasteArea');
    const masterDataPasteInput = document.getElementById('masterDataPasteInput');
    const processMasterPasteBtn = document.getElementById('processMasterPasteBtn');

    const statusDataFileInput = document.getElementById('statusDataFileInput');
    const statusFileLabel = document.getElementById('statusFileLabel');
    const statusFileInfoCard = document.getElementById('statusFileInfoCard');
    const stFileName = document.getElementById('stFileName');
    const stFileMeta = document.getElementById('stFileMeta');
    const changeStFileBtn = document.getElementById('changeStFileBtn');

    const modeUploadBtn = document.getElementById('modeUploadBtn');
    const modePasteBtn = document.getElementById('modePasteBtn');
    const statusFileUploadArea = document.getElementById('statusFileUploadArea');
    const statusPasteArea = document.getElementById('statusPasteArea');

    const statusInput = document.getElementById('statusInput');
    const processAnalyzerBtn = document.getElementById('processAnalyzerBtn');
    const clearAnalyzerBtn = document.getElementById('clearAnalyzerBtn');
    const analyzerResultsContainer = document.getElementById('analyzerResultsContainer');
    const downloadAnalyzerBtn = document.getElementById('downloadAnalyzerBtn');

    // --- COO Checker Elements ---
    const cooCsvInput = document.getElementById('cooCsvInput');
    const cooCsvLabel = document.getElementById('cooCsvLabel');
    const cooCsvInfo = document.getElementById('cooCsvInfo');
    const cooCsvName = document.getElementById('cooCsvName');
    const cooCsvMeta = document.getElementById('cooCsvMeta');
    const changeCooCsvBtn = document.getElementById('changeCooCsvBtn');
    const cooSkuPairInput = document.getElementById('cooSkuPairInput');
    const processCooBtn = document.getElementById('processCooBtn');
    const clearCooBtn = document.getElementById('clearCooBtn');
    const cooResultsContainer = document.getElementById('cooResultsContainer');
    const downloadCooBtn = document.getElementById('downloadCooBtn');

    let processedGroups = [];
    let analyzerExportData = [];
    let masterData = new Map(); // OriginalSKU -> NewSKU
    let currentStatusFileContent = null;
    let analyzerMode = 'paste'; // 'upload' or 'paste'

    let masterCOOData = new Map(); // SKU -> Country of Origin
    let cooExportData = [];

    // --- Exclusion Checker Elements ---
    const exclusionCsvInput = document.getElementById('exclusionCsvInput');
    const exclusionCsvLabel = document.getElementById('exclusionCsvLabel');
    const exclusionCsvInfo = document.getElementById('exclusionCsvInfo');
    const exdisplayFileName = document.getElementById('exdisplayFileName');
    const exdisplayFileMeta = document.getElementById('exdisplayFileMeta');
    const changeExclusionCsvBtn = document.getElementById('changeExclusionCsvBtn');
    const exclusionDataModeFileBtn = document.getElementById('exclusionDataModeFileBtn');
    const exclusionDataModePasteBtn = document.getElementById('exclusionDataModePasteBtn');
    const exclusionDataFileArea = document.getElementById('exclusionDataFileArea');
    const exclusionDataPasteArea = document.getElementById('exclusionDataPasteArea');
    const exclusionDataPasteInput = document.getElementById('exclusionDataPasteInput');
    const processExclusionPasteBtn = document.getElementById('processExclusionPasteBtn');
    const exclusionSkuInput = document.getElementById('exclusionSkuInput');
    const processExclusionBtn = document.getElementById('processExclusionBtn');
    const clearExclusionBtn = document.getElementById('clearExclusionBtn');
    const exclusionResultsContainer = document.getElementById('exclusionResultsContainer');
    const downloadExclusionBtn = document.getElementById('downloadExclusionBtn');

    let masterExclusionData = new Map(); // SKU -> { flags }
    let exclusionExportData = [];

    // --- Theme Toggle Logic ---
    const themeToggle = document.getElementById('themeToggle');
    const body = document.body;

    // Check for saved theme
    if (localStorage.getItem('theme') === 'dark') {
        body.classList.add('dark-theme');
        updateThemeIcon();
    }

    themeToggle.addEventListener('click', () => {
        body.classList.toggle('dark-theme');
        const isDark = body.classList.contains('dark-theme');
        localStorage.setItem('theme', isDark ? 'dark' : 'light');
        updateThemeIcon();
    });

    function updateThemeIcon() {
        const isDark = body.classList.contains('dark-theme');
        themeToggle.innerHTML = isDark ? '<i data-lucide="sun"></i>' : '<i data-lucide="moon"></i>';
        lucide.createIcons();
    }

    // --- AI Analyst Selection Logic & UI ---
    const aiToggleConfigBtn = document.getElementById('aiToggleConfigBtn');
    const aiConfigToggleIcon = document.getElementById('aiConfigToggleIcon');
    const aiConfigToggleText = document.getElementById('aiConfigToggleText');
    const aiToggleDbBatchBtn = document.getElementById('aiToggleDbBatchBtn');
    const aiDbBatchToggleIcon = document.getElementById('aiDbBatchToggleIcon');
    const aiDbBatchToggleText = document.getElementById('aiDbBatchToggleText');
    const aiMainGrid = document.getElementById('aiMainGrid');

    let isAiConfigCollapsed = false;
    let isAiDbBatchCollapsed = false;

    if (aiToggleConfigBtn) {
        aiToggleConfigBtn.addEventListener('click', () => {
            isAiConfigCollapsed = !isAiConfigCollapsed;
            aiMainGrid.classList.toggle('collapsed-config', isAiConfigCollapsed);

            if (isAiConfigCollapsed) {
                aiConfigToggleIcon.setAttribute('data-lucide', 'eye');
                aiConfigToggleText.textContent = 'Show Config';
                showToast("AI Configuration hidden.");
            } else {
                aiConfigToggleIcon.setAttribute('data-lucide', 'eye-off');
                aiConfigToggleText.textContent = 'Hide Config';
            }
            lucide.createIcons();
        });
    }

    if (aiToggleDbBatchBtn) {
        aiToggleDbBatchBtn.addEventListener('click', () => {
            isAiDbBatchCollapsed = !isAiDbBatchCollapsed;
            aiMainGrid.classList.toggle('collapsed-db-batch', isAiDbBatchCollapsed);

            if (isAiDbBatchCollapsed) {
                aiDbBatchToggleIcon.setAttribute('data-lucide', 'eye');
                aiDbBatchToggleText.textContent = 'Show DB & Batch';
                showToast("Product Database & Batch Analysis hidden.");
            } else {
                aiDbBatchToggleIcon.setAttribute('data-lucide', 'eye-off');
                aiDbBatchToggleText.textContent = 'Hide DB & Batch';
            }
            lucide.createIcons();
        });
    }

    // --- Tab & Mode Switching Logic ---
    navTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const target = tab.dataset.tab;
            navTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            tabContents.forEach(content => {
                content.id === `${target}Tab` ? content.classList.add('active') : content.classList.remove('active');
            });
            lucide.createIcons();
        });
    });

    modeUploadBtn.addEventListener('click', () => setAnalyzerMode('upload'));
    modePasteBtn.addEventListener('click', () => setAnalyzerMode('paste'));

    function setAnalyzerMode(mode) {
        analyzerMode = mode;
        if (mode === 'upload') {
            modeUploadBtn.classList.add('active');
            modeUploadBtn.style.background = '#3b82f6';
            modeUploadBtn.style.color = 'white';
            modePasteBtn.classList.remove('active');
            modePasteBtn.style.background = 'transparent';
            modePasteBtn.style.color = 'var(--text-dim)';
            statusFileUploadArea.style.display = 'block';
            statusPasteArea.style.display = 'none';
        } else {
            modePasteBtn.classList.add('active');
            modePasteBtn.style.background = '#3b82f6';
            modePasteBtn.style.color = 'white';
            modeUploadBtn.classList.remove('active');
            modeUploadBtn.style.background = 'transparent';
            modeUploadBtn.style.color = 'var(--text-dim)';
            statusFileUploadArea.style.display = 'none';
            statusPasteArea.style.display = 'block';
        }
    }

    // --- SKU Merger Logic ---
    // ... (skipped for brevity, but stays same)
    processBtn.addEventListener('click', () => {
        const input = skuInput.value.trim();
        if (!input) return;
        processSKUs(input);
    });

    clearBtn.addEventListener('click', () => {
        skuInput.value = '';
        resultsContainer.innerHTML = '<div class="empty-state"><p>Paste data above and click "Process" to see results.</p></div>';
        statsContainer.textContent = 'No results yet';
        downloadBtn.style.display = 'none';
        processedGroups = [];
    });

    downloadBtn.addEventListener('click', () => {
        if (processedGroups.length === 0) return;
        downloadCSV(processedGroups);
    });

    function processSKUs(text) {
        const lines = text.split('\n');
        const adj = new Map();
        const allSkus = new Set();
        lines.forEach(line => {
            const skus = line.split(/[\t\s,]+/).map(s => s.trim().replace(/^["']|["']$/g, '')).filter(s => s && s.toLowerCase() !== 'sku1' && s.toLowerCase() !== 'sku2');
            if (skus.length > 0) {
                skus.forEach(sku => {
                    allSkus.add(sku);
                    if (!adj.has(sku)) adj.set(sku, new Set());
                });
                if (skus.length > 1) {
                    for (let i = 0; i < skus.length; i++) {
                        for (let j = i + 1; j < skus.length; j++) {
                            adj.get(skus[i]).add(skus[j]);
                            adj.get(skus[j]).add(skus[i]);
                        }
                    }
                }
            }
        });
        const visited = new Set();
        const components = [];
        allSkus.forEach(sku => {
            if (!visited.has(sku)) {
                const component = [];
                const queue = [sku];
                visited.add(sku);
                while (queue.length > 0) {
                    const current = queue.shift();
                    component.push(current);
                    const neighbors = adj.get(current);
                    if (neighbors) {
                        neighbors.forEach(neighbor => {
                            if (!visited.has(neighbor)) {
                                visited.add(neighbor);
                                queue.push(neighbor);
                            }
                        });
                    }
                }
                components.push(component.sort());
            }
        });
        components.sort((a, b) => a[0].localeCompare(b[0]));
        processedGroups = components;
        downloadBtn.style.display = components.length > 0 ? 'flex' : 'none';
        renderResults(components);
    }

    function renderResults(components) {
        resultsContainer.innerHTML = '';
        if (components.length === 0) {
            resultsContainer.innerHTML = '<div class="empty-state"><p>No SKUs found.</p></div>';
            statsContainer.textContent = '0 groups found';
            return;
        }
        statsContainer.textContent = `${components.length} unique groups found`;
        components.forEach((comp, index) => {
            const card = document.createElement('div');
            card.className = 'result-card';
            card.style.animation = `fadeInUp 0.5s ease-out ${index * 0.05}s both`;
            const skuList = document.createElement('div');
            skuList.className = 'sku-list';
            comp.forEach(sku => {
                const tag = document.createElement('span');
                tag.className = 'sku-tag';
                tag.textContent = sku;
                skuList.appendChild(tag);
            });
            const copyBtn = document.createElement('button');
            copyBtn.className = 'copy-btn';
            copyBtn.innerHTML = '<i data-lucide="copy"></i> Copy Group';
            copyBtn.onclick = () => {
                navigator.clipboard.writeText(comp.join(', ')).then(() => showToast());
            };
            card.appendChild(skuList);
            card.appendChild(copyBtn);
            resultsContainer.appendChild(card);
        });
        lucide.createIcons();
    }

    function downloadCSV(groups) {
        let csvContent = "SKUs\n";
        groups.forEach(group => { csvContent += group.join(",") + "\n"; });
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "consolidation_template.csv";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // --- Helper for Excel/CSV Processing ---
    function processFile(file, callback) {
        const reader = new FileReader();
        const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

        reader.onload = (e) => {
            if (isExcel) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const csvData = XLSX.utils.sheet_to_csv(firstSheet);
                callback(csvData);
            } else {
                callback(e.target.result);
            }
        };

        if (isExcel) reader.readAsArrayBuffer(file);
        else reader.readAsText(file);
    }

    // --- Status Analyzer Logic ---
    masterFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        processFile(file, (content) => handleMasterData(content, file.name));
    });

    masterDataModeFileBtn.addEventListener('click', () => {
        masterDataModeFileBtn.classList.add('active');
        masterDataModePasteBtn.classList.remove('active');
        masterDataFileArea.style.display = 'block';
        masterDataPasteArea.style.display = 'none';
        lucide.createIcons();
    });

    masterDataModePasteBtn.addEventListener('click', () => {
        masterDataModePasteBtn.classList.add('active');
        masterDataModeFileBtn.classList.remove('active');
        masterDataPasteArea.style.display = 'block';
        masterDataFileArea.style.display = 'none';
        lucide.createIcons();
    });

    processMasterPasteBtn.addEventListener('click', () => {
        const content = masterDataPasteInput.value.trim();
        if (!content) {
            showToast("Please paste some data first.");
            return;
        }
        handleMasterData(content, "Pasted Master Data");
        lucide.createIcons();
    });

    function handleMasterData(content, fileName) {
        const lines = content.split(/\r?\n/).filter(line => line.trim());
        if (lines.length === 0) return;

        masterData.clear();
        lines.forEach((line, index) => {
            if (index === 0 && (line.toLowerCase().includes('sku') || line.toLowerCase().includes('original'))) return;
            const parts = line.split(/[,;\t]/);
            if (parts.length >= 2) {
                const original = parts[0].trim().replace(/^["']|["']$/g, '');
                const next = parts[1].trim().replace(/^["']|["']$/g, '');
                if (original) {
                    if (!masterData.has(original)) masterData.set(original, new Set());
                    masterData.get(original).add(next);
                }
            }
        });

        displayFileName.textContent = fileName;
        displayFileMeta.textContent = `${masterData.size.toLocaleString()} records loaded`;

        // UI Flow based on mode
        if (masterDataModeFileBtn.classList.contains('active')) {
            masterDataFileArea.style.display = 'none';
            fileLabel.style.display = 'none';
        }

        fileInfoCard.style.display = 'flex';
        fileInfoCard.classList.add('active');
        lucide.createIcons();
    }

    changeFileBtn.addEventListener('click', () => {
        fileInfoCard.style.display = 'none';
        fileInfoCard.classList.remove('active');

        // Restore based on mode
        if (masterDataModeFileBtn.classList.contains('active')) {
            fileLabel.style.display = 'flex';
            masterDataFileArea.style.display = 'block';
            masterFileInput.value = '';
        } else {
            // Already visible in paste mode
        }

        masterData.clear();
    });

    // Handle Status Data File
    statusDataFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        processFile(file, (content) => {
            currentStatusFileContent = content;
            stFileName.textContent = file.name;
            statusFileLabel.style.display = 'none';
            statusFileInfoCard.classList.add('active');
            lucide.createIcons();
        });
    });

    changeStFileBtn.addEventListener('click', () => {
        statusFileInfoCard.classList.remove('active');
        statusFileLabel.style.display = 'flex';
        statusDataFileInput.value = '';
        currentStatusFileContent = null;
    });

    processAnalyzerBtn.addEventListener('click', () => {
        let input = "";
        if (analyzerMode === 'upload') {
            if (!currentStatusFileContent) {
                showToast("Please upload a source data file or switch to Paste mode.");
                return;
            }
            input = currentStatusFileContent;
        } else {
            input = statusInput.value.trim();
            if (!input) return;
        }

        // Only enforce master data if the input contains consolidation-related keywords
        const IR_CONSO = "internally removed - consolidated/duplicate";
        if (masterData.size === 0 && input.toLowerCase().includes("consolidated")) {
            // We'll proceed but show a warning toast
            showToast("Note: No Consolidation File loaded. Rows requiring it will be skipped.");
        }

        analyzeStatuses(input);
    });

    downloadAnalyzerBtn.addEventListener('click', () => {
        if (analyzerExportData.length === 0) return;
        downloadAnalyzerCSV();
    });

    function downloadAnalyzerCSV() {
        const headers = ["SKU1", "SKU2", "Target", "Resolved?", "Resolution", "Image + Tags Updated?", "Note - Details (TH AIPN - Conso)", "Reason Not Resolved", "Note - Details"];

        let csvContent = headers.join(",") + "\n";

        analyzerExportData.forEach(row => {
            const line = [
                `"${row.sku1 || ''}"`,
                `"${row.sku2 || ''}"`,
                `"${row.resolved === 'Yes' ? row.target : ''}"`,
                `"${row.resolved}"`,
                `"${row.resolved === 'Yes' ? row.resolution : ''}"`,
                `""`, // Image + Tags Updated?
                `"${row.noteAipn !== undefined ? row.noteAipn : (row.resolved === 'Yes' ? row.note : '')}"`, // AIPN Note
                `""`, // Reason Not Resolved
                `"${row.noteDetails !== undefined ? row.noteDetails : (row.resolved === 'No' ? row.note : '')}"`  // Note Details
            ];
            csvContent += line.join(",") + "\n";
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `status_analysis_${new Date().toISOString().slice(0, 10)}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    clearAnalyzerBtn.addEventListener('click', () => {
        statusInput.value = '';
        analyzerResultsContainer.innerHTML = '<div class="empty-state"><p>Upload a master file and paste data to see analysis.</p></div>';
    });

    function analyzeStatuses(text) {
        if (!text) return;

        const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
        if (lines.length === 0) return;

        analyzerResultsContainer.innerHTML = '';
        analyzerExportData = []; // Clear old results
        let validLines = 0;
        let skipCount = 0;

        lines.forEach((line, index) => {
            const lowerLine = line.toLowerCase();

            // Skip headers
            if (lowerLine.includes('sku1') || lowerLine.includes('targetbts') || (lowerLine.includes('sku') && lowerLine.includes('status'))) {
                return;
            }

            // Adaptive Splitting
            let parts = [];
            const separators = [/\t/, /  +/, ',', ';'];

            for (let sep of separators) {
                const trial = line.split(sep).map(s => s.trim().replace(/^["']|["']$/g, ''));
                if (trial.length >= 5) {
                    parts = trial;
                    break;
                }
            }

            if (parts.length < 5) {
                parts = line.split(/[\t]+| {2,}/).map(s => s.trim().replace(/^["']|["']$/g, ''));
            }

            if (parts.length < 5) {
                skipCount++;
                return;
            }

            validLines++;
            const [sku1, sku2, stt1, stt2, targetBTS] = parts;

            let target, source, sttTarget, sttSource;

            // Detect user's Excel format (OriginalSKU, NewSKU, PrStatus, PrMaID, MaName)
            // PrStatus (stt1) is usually a number like '4'
            const isUserExcel = parts.length === 5 && !isNaN(stt1) && parts[4].includes(' ');

            if (isUserExcel) {
                source = sku1;
                target = sku2;
                sttSource = "internally removed - consolidated/duplicate";
                sttTarget = stt1;
            } else if (targetBTS === sku1) {
                target = sku1; sttTarget = stt1;
                source = sku2; sttSource = stt2;
            } else {
                target = sku2; sttTarget = stt2;
                source = sku1; sttSource = stt1;
            }

            const result = generateAnalysis(source, target, sttSource, sttTarget);
            if (result) {
                result.sku1 = sku1;
                result.sku2 = sku2;
                result.export.sku1 = sku1;
                result.export.sku2 = sku2;
                analyzerExportData.push(result.export);
                renderAnalyzerResult(result, index);
            }
        });

        downloadAnalyzerBtn.style.display = analyzerExportData.length > 0 ? 'flex' : 'none';

        if (validLines === 0 && skipCount > 0) {
            analyzerResultsContainer.innerHTML = `
                <div class="result-card" style="border-color: #fecaca; background: #fff1f2;">
                    <div class="sku-list" style="flex-direction: column; align-items: flex-start;">
                        <span class="sku-tag danger">Formatting Error</span>
                        <div class="result-text" style="color: #991b1b;">
                            Found ${skipCount} rows that don't have enough columns. Please ensure you provide all 5 columns: <strong>SKU1, SKU2, Stt1, Stt2, TargetBTS</strong>.
                        </div>
                    </div>
                </div>
            `;
        } else if (skipCount > 0) {
            showToast(`Analyzed ${validLines} lines. Skipped ${skipCount} rows due to column mismatch.`);
        }

        lucide.createIcons();
    }

    function generateAnalysis(source, target, sttSource, sttTarget) {
        const IR_CONSO = "internally removed - consolidated/duplicate";
        const BARRIER = "IR Visual Duplicate Barrier";

        // Neutral statuses (ignored)
        const neutralStatuses = [
            "live product",
            "can be deleted",
            "admins can sell - visual duplicate",
            "4"
        ];

        const s1 = (sttSource || "").toLowerCase().trim();
        const s2 = (sttTarget || "").toLowerCase().trim();

        const isNeutral = (s) => neutralStatuses.includes(s);
        const isConso = (s) => s === IR_CONSO;
        const isBarrier = (s) => !isNeutral(s) && !isConso(s);

        // If both are live/neutral, we don't need to analyze
        if (isNeutral(s1) && isNeutral(s2)) return null;

        const warnings = [];
        const checkMultiple = (sku, status) => {
            if (isConso(status)) {
                const targets = masterData.get(sku);
                if (targets && targets.size > 1) {
                    warnings.push(`SKU ${sku} has multiple targets recorded.`);
                }
            }
        };
        checkMultiple(source, s1);
        checkMultiple(target, s2);

        const getConsoInfo = (orig, next) => {
            const targets = masterData.get(orig);
            return (targets && targets.has(next));
        };

        const getDisplayLabel = (s) => {
            // Capitalize first letter of each word for display
            return s.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
        };

        // --- Logic: Consolidated ---
        if (isConso(s1)) {
            if (getConsoInfo(source, target)) {
                // New logic: Check if target status is NOT Live Product (Status 4 or "live product")
                if (s2 !== "4" && s2 !== "live product") {
                    warnings.push(`Target SKU ${target} is not in Live Product status`);
                }

                const txt = `SKU ${source} was consolidated into SKU ${target}`;
                return {
                    tags: [
                        { text: target, type: 'success' },
                        { text: 'Yes', type: 'success' },
                        { text: 'Consolidated', type: 'success' }
                    ],
                    sentence: txt, copy: txt, warnings,
                    export: { target, resolved: 'Yes', resolution: 'Consolidated', note: txt }
                };
            }
        }

        if (isConso(s1) && isConso(s2)) {
            const mSource = masterData.get(source);
            const mTarget = masterData.get(target);
            if (mSource && mTarget) {
                let commonTarget = null;
                for (let sku of mSource) {
                    if (mTarget.has(sku)) {
                        commonTarget = sku;
                        break;
                    }
                }

                if (commonTarget) {
                    const txt = `Both SKUs were consolidated into SKU ${commonTarget}`;
                    return {
                        tags: [{ text: commonTarget, type: 'success' }, { text: 'Consolidated', type: 'success' }],
                        sentence: txt, copy: txt, warnings,
                        export: { target: commonTarget, resolved: 'Yes', resolution: 'Consolidated', note: txt }
                    };
                } else {
                    const sList = Array.from(mSource).join(', ');
                    const tList = Array.from(mTarget).join(', ');
                    const combinedList = `${sList},${tList}`;
                    
                    warnings.push("No common target");

                    return {
                        tags: [],
                        sentence: "",
                        copy: combinedList,
                        isError: true,
                        warnings,
                        export: { target: '', resolved: '', resolution: '', note: '', noteAipn: combinedList, noteDetails: '' }
                    };
                }
            }
        }

        // --- Logic: Barriers (Supplier Discontinued, Visual Duplicate, and anything else not Neutral/Conso) ---
        if (isBarrier(s1) && isBarrier(s2)) {
            const txt = (s1 === s2)
                ? `Both SKUs are in ${getDisplayLabel(s1)} status`
                : `Source SKU ${source} is in ${getDisplayLabel(s1)} status; Target SKU ${target} is in ${getDisplayLabel(s2)} status`;
            return {
                tags: [{ text: target, type: 'success' }, { text: 'Yes', type: 'success' }, { text: BARRIER, type: 'warning' }],
                sentence: txt, copy: txt, warnings,
                export: { target, resolved: 'Yes', resolution: BARRIER, note: txt }
            };
        }

        if (isBarrier(s2) && isNeutral(s1)) {
            const txt = `Target SKU ${target} is in ${getDisplayLabel(s2)} status`;
            return {
                tags: [{ text: 'No', type: 'danger' }],
                sentence: txt, copy: txt, warnings,
                export: { target, resolved: 'No', resolution: BARRIER, note: txt }
            };
        }

        if (isBarrier(s1)) {
            const txt = `Source SKU ${source} is in ${getDisplayLabel(s1)} status`;
            return {
                tags: [{ text: target, type: 'success' }, { text: 'Yes', type: 'success' }, { text: BARRIER, type: 'warning' }],
                sentence: txt, copy: txt, warnings,
                export: { target, resolved: 'Yes', resolution: BARRIER, note: txt }
            };
        }

        return null;
    }

    function renderAnalyzerResult(result, index) {
        const card = document.createElement('div');
        card.className = 'result-card';
        card.style.animation = `fadeInUp 0.5s ease-out ${index * 0.03}s both`;

        if (result.isError) {
            card.style.borderColor = '#ef4444';
            card.style.borderWidth = '2px';
            card.style.background = '#fff1f2';
        }

        const leftContent = document.createElement('div');
        leftContent.className = 'sku-list';
        leftContent.style.flexDirection = 'column';
        leftContent.style.alignItems = 'flex-start';

        const tagContainer = document.createElement('div');
        tagContainer.className = 'sku-list';
        tagContainer.style.marginBottom = '0.5rem';

        if (result.sku1 && result.sku2) {
            const span = document.createElement('span');
            span.className = 'sku-tag';
            span.textContent = `${result.sku1},${result.sku2}`;
            tagContainer.appendChild(span);
        }

        result.tags.forEach(tag => {
            const span = document.createElement('span');
            span.className = `sku-tag ${tag.type}`;
            span.textContent = tag.text;
            tagContainer.appendChild(span);
        });

        const textDiv = document.createElement('div');
        textDiv.className = 'result-text';
        textDiv.textContent = result.sentence;

        leftContent.appendChild(tagContainer);
        leftContent.appendChild(textDiv);

        if (result.warnings && result.warnings.length > 0) {
            result.warnings.forEach(w => {
                const warnDiv = document.createElement('div');
                warnDiv.className = 'analyzer-warning';
                warnDiv.innerHTML = `<i data-lucide="alert-triangle"></i> ${w}`;
                leftContent.appendChild(warnDiv);
            });
        }

        const copyBtn = document.createElement('button');
        copyBtn.className = 'copy-btn';
        copyBtn.innerHTML = '<i data-lucide="copy"></i> Copy Note';
        copyBtn.onclick = () => {
            navigator.clipboard.writeText(result.copy).then(showToast);
        };

        card.appendChild(leftContent);
        card.appendChild(copyBtn);
        analyzerResultsContainer.appendChild(card);
    }

    // --- Option Matcher Elements ---
    const matcherCsvInput = document.getElementById('matcherCsvInput');
    const matcherCsvInfo = document.getElementById('matcherCsvInfo');
    const matcherCsvName = document.getElementById('matcherCsvName');
    const matcherCsvMeta = document.getElementById('matcherCsvMeta');
    const changeMatcherCsvBtn = document.getElementById('changeMatcherCsvBtn');
    const matcherCsvLabel = document.getElementById('matcherCsvLabel');

    const skuPairInput = document.getElementById('skuPairInput');
    const skuAInput = document.getElementById('skuAInput');
    const skuBInput = document.getElementById('skuBInput');
    const skuALabel = document.getElementById('skuALabel');
    const skuBLabel = document.getElementById('skuBLabel');
    // Removed old mode toggle buttons and areas
    const skuASelectArea = document.getElementById('skuASelectArea');
    const skuBSelectArea = document.getElementById('skuBSelectArea');
    const matcherSkuASearch = document.getElementById('matcherSkuASearch');
    const matcherSkuBSearch = document.getElementById('matcherSkuBSearch');
    const matcherDataModeFileBtn = document.getElementById('matcherDataModeFileBtn');
    const matcherDataModePasteBtn = document.getElementById('matcherDataModePasteBtn');
    const matcherDataFileArea = document.getElementById('matcherDataFileArea');
    const matcherDataPasteArea = document.getElementById('matcherDataPasteArea');
    const matcherDataPasteInput = document.getElementById('matcherDataPasteInput');
    const processMatcherPasteBtn = document.getElementById('processMatcherPasteBtn');

    let imageRows = [];
    let selectedPartA = null;
    let selectedPartB = null;
    let matcherSortCol = 'PrSKU';
    let matcherSortDir = 'asc';
    let allPartsA = new Set();
    let allPartsB = new Set();
    // Removed matcherInputMode
    let fullRawOptionsA = []; // Array of {name, part, sku}
    let fullRawOptionsB = [];
    let matcherSearchQueryA = '';
    let matcherSearchQueryB = '';

    // --- Option Matcher Logic ---
    matcherCsvInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        processFile(file, (content) => handleSupplierData(content, file.name));
    });

    matcherDataModeFileBtn.addEventListener('click', () => {
        matcherDataModeFileBtn.classList.add('active');
        matcherDataModePasteBtn.classList.remove('active');
        matcherDataFileArea.style.display = 'block';
        matcherDataPasteArea.style.display = 'none';
        lucide.createIcons();
    });

    matcherDataModePasteBtn.addEventListener('click', () => {
        matcherDataModePasteBtn.classList.add('active');
        matcherDataModeFileBtn.classList.remove('active');
        matcherDataPasteArea.style.display = 'block';
        matcherDataFileArea.style.display = 'none';
        lucide.createIcons();
    });

    processMatcherPasteBtn.addEventListener('click', () => {
        const content = matcherDataPasteInput.value.trim();
        if (!content) {
            showToast("Please paste some data first.");
            return;
        }
        handleSupplierData(content, "Pasted Data");
        lucide.createIcons();
    });

    function handleSupplierData(content, fileName) {
        const lines = content.split(/\r?\n/).filter(line => line.trim());
        if (lines.length === 0) return;

        const getParts = (line) => {
            if (line.includes('\t')) return line.split('\t').map(s => s.trim().replace(/^["']|["']$/g, ''));
            return line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.trim().replace(/^["']|["']$/g, ''));
        };

        let colMap = { sku: -1, part: -1, supplier: -1, name: -1, desc: -1 };
        let headerIdx = -1;

        for (let i = 0; i < Math.min(lines.length, 50); i++) {
            const parts = getParts(lines[i]);
            parts.forEach((p, pIdx) => {
                const s = p.toLowerCase().trim();
                if (!s) return;
                if (s === 'prsku' || s === 'sku' || s === 'wayfair listing') colMap.sku = pIdx;
                if (s.includes('manufacturer part number') || s === 'mpn' || s === 'part number' || s === 'actualmanufacturerpartnumber') colMap.part = pIdx;
                if (s.includes('supplierid') || s === 'supplier id') colMap.supplier = pIdx;
                if (s === 'optionname' || s.includes('option name')) colMap.name = pIdx;
                if (s === 'description' || s === 'product name') colMap.desc = pIdx;
            });
            if (colMap.sku !== -1 && colMap.part !== -1) {
                headerIdx = i;
                break;
            }
        }

        if (headerIdx === -1) {
            showToast("Error: Missing PrSKU or Part Number columns.");
            return;
        }

        imageRows = lines.slice(headerIdx + 1).map(line => {
            const parts = getParts(line);
            const row = {
                PrSKU: (parts[colMap.sku] || '').trim(),
                ActualManufacturerPartNumber: (parts[colMap.part] || '').trim(),
                SupplierID: (parts[colMap.supplier] || '').trim(),
                OptionName: (parts[colMap.name] || parts[colMap.desc] || '').trim()
            };
            return row;
        }).filter(r => r.PrSKU && r.ActualManufacturerPartNumber);

        const uniquePrSkus = new Set(imageRows.map(r => r.PrSKU).filter(s => s));
        if (uniquePrSkus.size < 2) {
            showToast("Error: Data must contain at least two different PrSKU values.");
            matcherCsvName.textContent = "Invalid Data";
            matcherCsvMeta.textContent = `Only ${uniquePrSkus.size} unique PrSKU(s) found`;
            matcherCsvInfo.style.background = "#fef2f2";
            matcherCsvInfo.style.borderColor = "#fecaca";

            // Show info card but keep inputs based on mode
            matcherCsvInfo.style.display = 'flex';
            matcherCsvInfo.classList.add('active');

            if (matcherDataModeFileBtn.classList.contains('active')) {
                matcherDataFileArea.style.display = 'none';
            }
            return;
        } else {
            matcherCsvInfo.style.background = "#f0fdf4";
            matcherCsvInfo.style.borderColor = "#bbf7d0";
        }

        matcherCsvName.textContent = fileName;
        matcherCsvMeta.textContent = `${imageRows.length.toLocaleString()} records loaded`;

        // UI Flow based on mode
        if (matcherDataModeFileBtn.classList.contains('active')) {
            matcherDataFileArea.style.display = 'none';
            matcherCsvLabel.style.display = 'none';
        }

        matcherCsvInfo.style.display = 'flex';
        matcherCsvInfo.classList.add('active');
        processCombinedMatcherInput();
        lucide.createIcons();
    }

    changeMatcherCsvBtn.addEventListener('click', () => {
        matcherCsvInfo.style.display = 'none';
        matcherCsvInfo.classList.remove('active');

        // Restore based on mode
        if (matcherDataModeFileBtn.classList.contains('active')) {
            matcherCsvLabel.style.display = 'flex';
            matcherDataFileArea.style.display = 'block';
            matcherCsvInput.value = '';
        } else {
            // In paste mode, it's already visible
        }

        imageRows = [];
        updateMatcherResults();
    });

    matcherSkuASearch.addEventListener('input', (e) => {
        matcherSearchQueryA = e.target.value.toLowerCase().trim();
        renderOptions('A', fullRawOptionsA);
    });

    matcherSkuBSearch.addEventListener('input', (e) => {
        matcherSearchQueryB = e.target.value.toLowerCase().trim();
        renderOptions('B', fullRawOptionsB);
    });

    skuPairInput.addEventListener('input', processCombinedMatcherInput);
    skuAInput.addEventListener('input', processCombinedMatcherInput);
    skuBInput.addEventListener('input', processCombinedMatcherInput);

    // Give time for text to be inserted before processing on paste
    skuPairInput.addEventListener('paste', () => setTimeout(processCombinedMatcherInput, 0));
    skuAInput.addEventListener('paste', () => setTimeout(processCombinedMatcherInput, 0));
    skuBInput.addEventListener('paste', () => setTimeout(processCombinedMatcherInput, 0));

    skuPairInput.addEventListener('blur', () => {
        const text = skuPairInput.value.trim();
        const skus = text.split(/[\s\t\n,]+/).filter(s => s.length > 0);
        if (skus.length > 0) skuPairInput.value = skus.join(', ');
    });

    const getOptionsForSku = (sku) => {
        if (!sku || imageRows.length === 0) return [];
        const searchSku = sku.toLowerCase().trim();
        const matchingRows = imageRows.filter(r => r.PrSKU.toLowerCase() === searchSku);
        const seenParts = new Set();
        const options = [];

        matchingRows.forEach(row => {
            const part = row.ActualManufacturerPartNumber;
            if (part && !seenParts.has(part.toLowerCase().trim())) {
                seenParts.add(part.toLowerCase().trim());
                options.push({ name: row.OptionName, part: part, sku: row.PrSKU });
            }
        });
        return options;
    };

    function processCombinedMatcherInput() {
        // Find inputs
        const skuText = skuPairInput.value.trim();
        const skus = skuText.split(/[\s\t\n,]+/).filter(s => s.length > 0);
        const skuAFilter = skus.length > 0 ? skus[0].toLowerCase() : null;
        const skuBFilters = skus.length > 1 ? skus.slice(1).map(s => s.toLowerCase()) : [];

        let manualOptionsA = parseManualOptions(skuAInput.value);
        let manualOptionsB = parseManualOptions(skuBInput.value);

        fullRawOptionsA = [];
        fullRawOptionsB = [];

        // Update Labels based on input
        if (skuAFilter) {
            skuALabel.innerHTML = `<i data-lucide="type"></i> ${skuAFilter.toUpperCase()} Options`;
        } else {
            skuALabel.innerHTML = `<i data-lucide="type"></i> SKU A Options`;
        }

        if (skuBFilters.length > 0) {
            const bText = skuBFilters.join(', ').toUpperCase();
            skuBLabel.innerHTML = `<i data-lucide="type"></i> ${bText} Options`;
        } else {
            skuBLabel.innerHTML = `<i data-lucide="type"></i> SKU B Options`;
        }
        lucide.createIcons();

        // Logic for SKU A Side
        if (skuAFilter) {
            let skuAOptions = getOptionsForSku(skuAFilter);
            if (manualOptionsA.length > 0) {
                // Intersect SKU A parts with manual typed parts 
                const manualPartsA = new Set(manualOptionsA.map(o => o.part.toLowerCase()));
                fullRawOptionsA = skuAOptions.filter(o => manualPartsA.has(o.part.toLowerCase()));
            } else {
                fullRawOptionsA = skuAOptions;
            }
        } else {
            // No SKU A, only manual options inputted
            fullRawOptionsA = getDetailedManualOptions(manualOptionsA);
        }

        // Logic for SKU B Side
        if (skuBFilters.length > 0) {
            let skuBOptions = [];
            let seenPartsB = new Set();
            skuBFilters.forEach(s => {
                const opts = getOptionsForSku(s);
                opts.forEach(o => {
                    const lowPart = o.part.toLowerCase().trim();
                    if (!seenPartsB.has(lowPart)) {
                        seenPartsB.add(lowPart);
                        skuBOptions.push(o);
                    }
                });
            });

            if (manualOptionsB.length > 0) {
                const manualPartsB = new Set(manualOptionsB.map(o => o.part.toLowerCase()));
                fullRawOptionsB = skuBOptions.filter(o => manualPartsB.has(o.part.toLowerCase()));
            } else {
                fullRawOptionsB = skuBOptions;
            }
        } else {
            fullRawOptionsB = getDetailedManualOptions(manualOptionsB);
        }

        allPartsA = new Set(fullRawOptionsA.map(o => o.part.toLowerCase().trim()));
        allPartsB = new Set(fullRawOptionsB.map(o => o.part.toLowerCase().trim()));

        // Check if selected parts are still valid
        const currentA = selectedPartA ? selectedPartA.toLowerCase().trim() : null;
        const currentB = selectedPartB ? selectedPartB.toLowerCase().trim() : null;

        if (fullRawOptionsA.length === 1) selectedPartA = fullRawOptionsA[0].part;
        else if (currentA && !allPartsA.has(currentA)) selectedPartA = null;

        if (fullRawOptionsB.length === 1) selectedPartB = fullRawOptionsB[0].part;
        else if (currentB && !allPartsB.has(currentB)) selectedPartB = null;

        renderOptions('A', fullRawOptionsA);
        renderOptions('B', fullRawOptionsB);
        updateMatcherResults();
    }
    
    // Attaches known Supplier Data (if active) to manual options
    function getDetailedManualOptions(manualOptions) {
        if (manualOptions.length === 0) return [];
        if (imageRows.length === 0) return manualOptions; // Fallback if no DB
        
        const enhanced = [];
        const seenParts = new Set();
        
        manualOptions.forEach(m => {
            const targetLower = m.part.toLowerCase();
            // Try to find it in the DB to grab Name and real SKU
            const matchingRow = imageRows.find(r => r.ActualManufacturerPartNumber.toLowerCase() === targetLower);
            
            if (!seenParts.has(targetLower)) {
                seenParts.add(targetLower);
                if (matchingRow) {
                    enhanced.push({
                        name: m.name || matchingRow.OptionName,
                        part: matchingRow.ActualManufacturerPartNumber,
                        sku: matchingRow.PrSKU
                    });
                } else {
                    enhanced.push(m);
                }
            }
        });
        
        return enhanced;
    }

    function parseManualOptions(text) {
        if (!text.trim()) return [];

        const lines = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
        const result = [];
        const commonHeaders = ['manufacturer part numbers', 'options', 'list', 'description', 'part number', 'part #'];

        lines.forEach(line => {
            const lower = line.toLowerCase();
            if (commonHeaders.includes(lower) || commonHeaders.some(h => lower === h + ':')) return;

            if (line.includes(',')) {
                const parts = line.split(',');
                const partNumber = parts.pop().trim();
                const optionName = parts.join(',').trim();
                if (partNumber) result.push({ name: optionName, part: partNumber, sku: 'Manual Option' });
            } else {
                result.push({ name: '', part: line, sku: 'Manual Option' });
            }
        });
        return result;
    }

    function renderOptions(type, options) {
        const area = type === 'A' ? skuASelectArea : skuBSelectArea;
        const query = type === 'A' ? matcherSearchQueryA : matcherSearchQueryB;
        const otherParts = type === 'A' ? allPartsB : allPartsA;

        area.innerHTML = '';

        if (options.length === 0) {
            area.innerHTML = '<p class="empty-text">Paste options to start</p>';
            return;
        }

        let filtered = options;

        // Apply Search
        if (query) {
            filtered = filtered.filter(opt =>
                opt.part.toLowerCase().includes(query) ||
                (opt.name && opt.name.toLowerCase().includes(query))
            );
        }

        if (filtered.length === 0) {
            area.innerHTML = query
                ? `<p class="empty-text">No matches for "${query}"</p>`
                : '<p class="empty-text">No shared part numbers</p>';
            return;
        }

        filtered.forEach((opt, idx) => {
            const lowPart = opt.part.toLowerCase().trim();
            const isShared = allPartsA.has(lowPart) && allPartsB.has(lowPart);
            const item = document.createElement('label');
            const isSelected = (type === 'A' ? selectedPartA : selectedPartB) === opt.part;
            item.className = `option-item ${isShared ? 'shared-option' : ''} ${isSelected ? 'selected' : ''}`;
            item.innerHTML = `
                <input type="radio" name="sku${type}Option" value="${opt.part}" ${isSelected ? 'checked' : ''}>
                <div class="option-info">
                    <div style="display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap;">
                        <span class="option-name">${opt.name || 'Option (Part Only)'}</span>
                        ${isShared ? '<span class="shared-badge">Same Part Number</span>' : ''}
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
                        <span class="option-part">${opt.part}</span>
                        <span style="font-size: 0.7rem; color: var(--text-dim); opacity: 0.7;">${opt.sku}</span>
                    </div>
                </div>
            `;

            item.querySelector('input').addEventListener('change', (e) => {
                const allItems = area.querySelectorAll('.option-item');
                allItems.forEach(i => i.classList.remove('selected'));
                item.classList.add('selected');

                if (type === 'A') selectedPartA = e.target.value;
                else selectedPartB = e.target.value;

                updateMatcherResults();
            });

            area.appendChild(item);
        });
    }

    function updateMatcherResults() {
        if (imageRows.length === 0 || (!selectedPartA && !selectedPartB)) {
            matcherResultsContainer.innerHTML = `
                <div class="empty-state">
                    <p>${imageRows.length === 0 ? 'Upload CSV' : 'Select an option'} to see filtered data.</p>
                </div>
            `;
            return;
        }

        const partA = selectedPartA ? selectedPartA.toLowerCase().trim() : null;
        const partB = selectedPartB ? selectedPartB.toLowerCase().trim() : null;

        // Collect all input SKUs for highlighting
        const inputSkus = new Set();
        const text = skuPairInput.value.trim();
        if (text) {
            text.split(/[\s\t\n,]+/).filter(s => s.length > 0).forEach(s => inputSkus.add(s.toLowerCase().trim()));
        }

        let filtered = imageRows.filter(row => {
            const rowPart = row.ActualManufacturerPartNumber.toLowerCase().trim();
            const rowSku = row.PrSKU.toLowerCase().trim();

            // Strictly only show SKUs that were in the input area (if any input SKUs exist)
            if (inputSkus.size > 0 && !inputSkus.has(rowSku)) return false;

            return (partA && rowPart === partA) || (partB && rowPart === partB);
        });

        if (filtered.length === 0) {
            matcherResultsContainer.innerHTML = `
                <div class="empty-state">
                    <p>No matches found in CSV for selected part numbers.</p>
                </div>
            `;
            return;
        }

        // --- Multi-Part Duplicate Detection ---
        const supplierIdsA = new Set(
            imageRows.filter(r => partA && r.ActualManufacturerPartNumber.toLowerCase().trim() === partA)
                .map(r => r.SupplierID)
        );
        const supplierIdsB = new Set(
            imageRows.filter(r => partB && r.ActualManufacturerPartNumber.toLowerCase().trim() === partB)
                .map(r => r.SupplierID)
        );
        const commonSupplierIds = new Set([...supplierIdsA].filter(id => supplierIdsB.has(id)));

        // --- Sorting ---
        filtered.sort((a, b) => {
            const skuA = String(a.PrSKU).toLowerCase();
            const skuB = String(b.PrSKU).toLowerCase();

            // Prioritize input SKUs in sorting
            const aIsInput = inputSkus.has(skuA);
            const bIsInput = inputSkus.has(skuB);
            if (aIsInput && !bIsInput) return -1;
            if (!aIsInput && bIsInput) return 1;

            let valA = a[matcherSortCol] || '';
            let valB = b[matcherSortCol] || '';

            if (!isNaN(valA) && !isNaN(valB) && valA !== '' && valB !== '') {
                return matcherSortDir === 'asc' ? Number(valA) - Number(valB) : Number(valB) - Number(valA);
            }

            return matcherSortDir === 'asc'
                ? String(valA).localeCompare(String(valB))
                : String(valB).localeCompare(String(valA));
        });

        const getSortIcon = (col) => {
            if (matcherSortCol !== col) return '<i data-lucide="chevrons-up-down" class="sort-icon-dim"></i>';
            return matcherSortDir === 'asc'
                ? '<i data-lucide="chevron-up" class="sort-icon-active"></i>'
                : '<i data-lucide="chevron-down" class="sort-icon-active"></i>';
        };

        let tableHtml = `
            <table class="matcher-table">
                <thead>
                    <tr>
                        <th class="sortable-header" data-col="PrSKU">
                            <div class="header-inner">PrSKU ${getSortIcon('PrSKU')}</div>
                        </th>
                        <th class="sortable-header" data-col="SupplierID">
                            <div class="header-inner">SupplierID ${getSortIcon('SupplierID')}</div>
                        </th>
                        <th>ActualManufacturerPartNumber</th>
                    </tr>
                </thead>
                <tbody>
                    ${filtered.map(row => {
            const rowPart = (row.ActualManufacturerPartNumber || '').toLowerCase().trim();
            const partA = (selectedPartA || '').toLowerCase().trim();
            const partB = (selectedPartB || '').toLowerCase().trim();
            const isDuplicate = commonSupplierIds.has(row.SupplierID) && selectedPartA && selectedPartB;

            // "Same Part Number" means this row's part number matches BOTH of our selected parts
            const isSamePart = rowPart !== '' && partA !== '' && partB !== '' && rowPart === partA && rowPart === partB;

            return `
                        <tr class="${isDuplicate || isSamePart ? 'row-highlight' : ''}">
                            <td>
                                <strong>${row.PrSKU || '-'}</strong>
                            </td>
                            <td class="${isDuplicate ? 'duplicate-cell' : ''}">
                                ${row.SupplierID || '-'}
                                ${isDuplicate ? '<span class="match-badge">Match Found</span>' : ''}
                            </td>
                            <td>
                                <code class="option-part">${row.ActualManufacturerPartNumber || '-'}</code>
                            </td>
                        </tr>
                        `;
        }).join('')}
                </tbody>
            </table>
        `;

        matcherResultsContainer.innerHTML = tableHtml;
        lucide.createIcons();

        matcherResultsContainer.querySelectorAll('.sortable-header').forEach(th => {
            th.addEventListener('click', () => {
                const col = th.dataset.col;
                if (matcherSortCol === col) {
                    matcherSortDir = matcherSortDir === 'asc' ? 'desc' : 'asc';
                } else {
                    matcherSortCol = col;
                    matcherSortDir = 'asc';
                }
                updateMatcherResults();
            });
        });
    }

    // --- COO Checker Logic ---
    cooCsvInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        showToast("Reading file...");
        const reader = new FileReader();
        const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

        reader.onload = (e) => {
            try {
                masterCOOData.clear();
                let sheetsFound = 0;

                if (isExcel) {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    sheetsFound = workbook.SheetNames.length;

                    workbook.SheetNames.forEach(sheetName => {
                        const sheet = workbook.Sheets[sheetName];
                        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                        if (rawRows.length === 0) return;

                        let colMap = { coo: -1, sku1: -1, sku2: -1 };
                        let lastHeaderRow = -1;

                        // Scan first 100 rows for headers independently
                        for (let i = 0; i < Math.min(rawRows.length, 100); i++) {
                            const row = rawRows[i];
                            if (!row || !Array.isArray(row)) continue;

                            row.forEach((cell, cellIdx) => {
                                const s = String(cell || "").toLowerCase().trim();
                                if (!s) return;

                                if (s.includes('country of origin') || s === 'coo' || (s.includes('country') && s.includes('origin'))) {
                                    if (colMap.coo === -1) {
                                        colMap.coo = cellIdx;
                                        lastHeaderRow = Math.max(lastHeaderRow, i);
                                    }
                                }
                                if (s.includes('wayfair listing')) {
                                    if (colMap.sku1 === -1) {
                                        colMap.sku1 = cellIdx;
                                        lastHeaderRow = Math.max(lastHeaderRow, i);
                                    }
                                }
                                if (s.includes('manufacturer part number')) {
                                    if (colMap.sku2 === -1) {
                                        colMap.sku2 = cellIdx;
                                        lastHeaderRow = Math.max(lastHeaderRow, i);
                                    }
                                }
                            });
                        }

                        if (colMap.coo !== -1) {
                            for (let i = lastHeaderRow + 1; i < rawRows.length; i++) {
                                const row = rawRows[i];
                                if (!row) continue;
                                const sku1 = colMap.sku1 !== -1 ? row[colMap.sku1] : null;
                                const sku2 = colMap.sku2 !== -1 ? row[colMap.sku2] : null;
                                let cooVal = (row[colMap.coo] === undefined || row[colMap.coo] === null) ? "" : String(row[colMap.coo]).trim();

                                const addVal = (sku, val) => {
                                    if (!sku) return;
                                    const s = String(sku).trim();
                                    if (!s) return;
                                    if (!masterCOOData.has(s)) masterCOOData.set(s, new Set());
                                    masterCOOData.get(s).add(val);
                                };

                                if (sku1) addVal(sku1, cooVal);
                                if (sku2) addVal(sku2, cooVal);
                            }
                        }
                    });
                } else {
                    const content = e.target.result;
                    const lines = content.split(/\r?\n/).filter(line => line.trim());
                    if (lines.length > 0) {
                        const getParts = (line) => {
                            if (line.includes('\t')) return line.split('\t').map(s => s.trim().replace(/^["']|["']$/g, ''));
                            return line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.trim().replace(/^["']|["']$/g, ''));
                        };

                        let colMap = { coo: -1, sku1: -1, sku2: -1 };
                        let lastHeaderLine = -1;

                        for (let i = 0; i < Math.min(lines.length, 100); i++) {
                            const parts = getParts(lines[i]);
                            parts.forEach((part, partIdx) => {
                                const s = String(part || "").toLowerCase().trim();
                                if (!s) return;

                                if (s.includes('country of origin') || s === 'coo' || (s.includes('country') && s.includes('origin'))) {
                                    if (colMap.coo === -1) {
                                        colMap.coo = partIdx;
                                        lastHeaderLine = Math.max(lastHeaderLine, i);
                                    }
                                }
                                if (s.includes('wayfair listing')) {
                                    if (colMap.sku1 === -1) {
                                        colMap.sku1 = partIdx;
                                        lastHeaderLine = Math.max(lastHeaderLine, i);
                                    }
                                }
                                if (s.includes('manufacturer part number')) {
                                    if (colMap.sku2 === -1) {
                                        colMap.sku2 = partIdx;
                                        lastHeaderLine = Math.max(lastHeaderLine, i);
                                    }
                                }
                            });
                        }

                        if (colMap.coo !== -1) {
                            lines.slice(lastHeaderLine + 1).forEach(line => {
                                const parts = getParts(line);
                                const sku1 = colMap.sku1 !== -1 ? parts[colMap.sku1] : null;
                                const sku2 = colMap.sku2 !== -1 ? parts[colMap.sku2] : null;
                                let cooVal = (parts[colMap.coo] === undefined || parts[colMap.coo] === null) ? "" : String(parts[colMap.coo]).trim();

                                const addVal = (sku, val) => {
                                    if (!sku) return;
                                    const s = String(sku).trim();
                                    if (!s) return;
                                    if (!masterCOOData.has(s)) masterCOOData.set(s, new Set());
                                    masterCOOData.get(s).add(val);
                                };

                                if (sku1) addVal(sku1, cooVal);
                                if (sku2) addVal(sku2, cooVal);
                            });
                        }
                    }
                }

                cooCsvName.textContent = file.name;
                cooCsvMeta.textContent = `${masterCOOData.size.toLocaleString()} SKUs mapped`;
                cooCsvLabel.style.display = 'none';
                cooCsvInfo.classList.add('active');

                showToast(`Success! Found ${masterCOOData.size.toLocaleString()} SKUs across ${sheetsFound || 1} sheet(s).`);
                lucide.createIcons();
            } catch (err) {
                showToast(`Error: ${err.message || 'Check console'}`);
            }
        };

        if (isExcel) reader.readAsArrayBuffer(file);
        else reader.readAsText(file);
    });

    changeCooCsvBtn.addEventListener('click', () => {
        cooCsvInfo.classList.remove('active');
        cooCsvLabel.style.display = 'flex';
        cooCsvInput.value = '';
        masterCOOData.clear();
    });

    clearCooBtn.addEventListener('click', () => {
        cooSkuPairInput.value = '';
        cooResultsContainer.innerHTML = '<div class="empty-state"><p>Upload data and paste SKUs to see differences.</p></div>';
        downloadCooBtn.style.display = 'none';
        cooExportData = [];
    });

    processCooBtn.addEventListener('click', () => {
        const text = cooSkuPairInput.value.trim();
        if (!text) return;

        const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
        cooExportData = [];

        lines.forEach(line => {
            // Split by tab, comma, or spaces
            const skus = line.split(/[\t,]| {2,}/).map(s => s.trim().replace(/^["']|["']$/g, '')).filter(s => s.length > 0);
            if (skus.length >= 2) {
                const sku1 = skus[0];
                const sku2 = skus[1];

                const cooSet1 = masterCOOData.get(sku1) || new Set();
                const cooSet2 = masterCOOData.get(sku2) || new Set();

                const coos1 = Array.from(cooSet1);
                const coos2 = Array.from(cooSet2);

                let resultRow = {
                    sku1,
                    sku2,
                    skuCombined: `${sku1},${sku2}`,
                    target: "",
                    resolved: "", // Blank for normal pairs
                    resolution: "",
                    imageUpdated: "",
                    aipnNote: "",
                    reason: "",
                    note: ""
                };

                if (coos1.length > 1 || coos2.length > 1) {
                    resultRow.resolved = "No";
                    resultRow.reason = "Different Country of Origin";
                    resultRow.note = "Variant Multi Country";
                } else {
                    const val1 = coos1[0] || null;
                    const val2 = coos2[0] || null;

                    if (val1 === null && val2 === null) {
                        // skip
                    } else if (val1 === val2) {
                        // skip
                    } else {
                        // Difference found
                        resultRow.resolved = "No";
                        resultRow.reason = "Different Country of Origin";
                        const showNull = (v) => v === null ? "Null" : v;
                        resultRow.note = `${showNull(val1)} / ${showNull(val2)}`;
                    }
                }
                cooExportData.push(resultRow);
            }
        });

        renderCooResults();
    });

    function renderCooResults() {
        if (cooExportData.length === 0) {
            cooResultsContainer.innerHTML = '<div class="empty-state"><p>No differences found or invalid input.</p></div>';
            downloadCooBtn.style.display = 'none';
            return;
        }

        let tableHtml = `
            <table class="matcher-table">
                <thead>
                    <tr>
                        <th>SKU1</th>
                        <th>SKU2</th>
                        <th>Status</th>
                        <th>Note</th>
                        <th style="width: 50px;"></th>
                    </tr>
                </thead>
                <tbody>
        `;

        cooExportData.forEach((row, idx) => {
            const statusTag = row.resolved ? `<span class="sku-tag danger">${row.resolved}</span>` : "";
            const copyNoteJs = `navigator.clipboard.writeText('${row.note.replace(/'/g, "\\'")}').then(() => { const t = document.getElementById('toast'); t.textContent = 'Copied to clipboard'; t.classList.add('show'); setTimeout(() => t.classList.remove('show'), 3000); })`;
            tableHtml += `
                <tr>
                    <td><code class="option-part">${row.sku1}</code></td>
                    <td><code class="option-part">${row.sku2}</code></td>
                    <td>${statusTag}</td>
                    <td><span class="text-dim" style="font-size: 0.85rem;">${row.note}</span></td>
                    <td>
                        <button class="icon-action-btn" onclick="${copyNoteJs}" title="Copy Note" style="background:transparent; border:none; cursor:pointer; color:var(--text-dim); transition:color 0.2s;">
                            <i data-lucide="copy" style="width:16px; height:16px;"></i>
                        </button>
                    </td>
                </tr>
            `;
        });

        tableHtml += `</tbody></table>`;
        cooResultsContainer.innerHTML = tableHtml;
        downloadCooBtn.style.display = 'flex';
        lucide.createIcons();
    }

    downloadCooBtn.addEventListener('click', () => {
        if (cooExportData.length === 0) return;

        // Use proper CSV formatting with quotes for all fields to prevent shifting
        const headers = ["SKU1", "SKU2", "SKU1,SKU2", "Target", "Resolved?", "Resolution", "Image + Tags Updated?", "Note - Details (TH AIPN - Conso)", "Reason Not Resolved", "Note - Details"];
        let csvContent = "\uFEFF"; // BOM for Excel
        csvContent += headers.map(h => `"${h}"`).join(",") + "\n";

        cooExportData.forEach(row => {
            const line = [
                row.sku1,
                row.sku2,
                row.skuCombined,
                row.target,
                row.resolved,
                row.resolution,
                row.imageUpdated,
                row.aipnNote,
                row.reason,
                row.note
            ];
            // Quote all values and join
            csvContent += line.map(val => `"${String(val || "").replace(/"/g, '""')}"`).join(",") + "\n";
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `coo_comparison_${new Date().toISOString().slice(0, 10)}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    // --- Tag Matcher Elements ---
    const tagCsvInput = document.getElementById('tagCsvInput');
    const tagCsvLabel = document.getElementById('tagCsvLabel');
    const tagCsvInfo = document.getElementById('tagCsvInfo');
    const tagFileStatusArea = document.getElementById('tagFileStatusArea');
    const tagCsvName = document.getElementById('tagCsvName');
    const tagCsvMeta = document.getElementById('tagCsvMeta');
    const changeTagCsvBtn = document.getElementById('changeTagCsvBtn');
    const tagSkuPairInput = document.getElementById('tagSkuPairInput');
    const clearTagBtn = document.getElementById('clearTagBtn');
    const tagResultsContainer = document.getElementById('tagResultsContainer');
    const tagSkuASelectArea = document.getElementById('tagSkuASelectArea');
    const tagSkuBSelectArea = document.getElementById('tagSkuBSelectArea');
    const tagSkuAHeader = document.getElementById('tagSkuAHeader');
    const tagSkuBHeader = document.getElementById('tagSkuBHeader');
    const tagSkuASearch = document.getElementById('tagSkuASearch');
    const tagSkuBSearch = document.getElementById('tagSkuBSearch');
    const tagFilterOverlapBtn = document.getElementById('tagFilterOverlapBtn');
    const tagToggleInputsBtn = document.getElementById('tagToggleInputsBtn');
    const tagToggleIcon = document.getElementById('tagToggleIcon');
    const tagToggleText = document.getElementById('tagToggleText');
    const tagUploadersRow = document.getElementById('tagUploadersRow');
    const tagToggleFilterBtn = document.getElementById('tagToggleFilterBtn');
    const tagFilterToggleIcon = document.getElementById('tagFilterToggleIcon');
    const tagFilterToggleText = document.getElementById('tagFilterToggleText');
    const tagToggleDatabaseBtn = document.getElementById('tagToggleDatabaseBtn');
    const tagDatabaseToggleIcon = document.getElementById('tagDatabaseToggleIcon');
    const tagDatabaseToggleText = document.getElementById('tagDatabaseToggleText');
    const tagSection1 = document.getElementById('tagSection1');
    const tagSectionDesc = document.getElementById('tagSectionDesc');
    const tagSection2 = document.getElementById('tagSection2');

    // New: Page Elements for Description/AI-style integration
    const tagDescInput = document.getElementById('tagDescInput');
    const tagDescLabel = document.getElementById('tagDescLabel');
    const tagDescStatusArea = document.getElementById('tagDescStatusArea');
    const tagDescFileName = document.getElementById('tagDescFileName');
    const tagDescFileMeta = document.getElementById('tagDescFileMeta');
    const changeTagDescBtn = document.getElementById('changeTagDescBtn');

    const tagToggleDescViewBtn = document.getElementById('tagToggleDescViewBtn');
    const tagDescToggleText = document.getElementById('tagDescToggleText');
    const tagDescComparisonSection = document.getElementById('tagDescComparisonSection');
    const tagDescComparisonResults = document.getElementById('tagDescComparisonResults');

    const tagOptionMapInput = document.getElementById('tagOptionMapInput');
    const tagOptionMapLabel = document.getElementById('tagOptionMapLabel');
    const tagOptionMapStatusArea = document.getElementById('tagOptionMapStatusArea');
    const tagOptionMapFileName = document.getElementById('tagOptionMapFileName');
    const tagOptionMapFileMeta = document.getElementById('tagOptionMapFileMeta');
    const changeTagOptionMapBtn = document.getElementById('changeTagOptionMapBtn');
    const tagSectionOptionMap = document.getElementById('tagSectionOptionMap');
    const tagToggleOptionMapBtn = document.getElementById('tagToggleOptionMapBtn');
    const tagOptionMapToggleIcon = document.getElementById('tagOptionMapToggleIcon');
    const tagOptionMapToggleText = document.getElementById('tagOptionMapToggleText');

    let masterTagData = new Map(); // SKU -> Map<MPN, { tagName: value }>
    let tagCatalogData = new Map(); // SKU -> { name, desc, bullets }
    let masterOptionMapData = new Map(); // SKU -> Map<MPN, constructedName>
    let tagExportData = [];
    let selectedTagPartA = null;
    let selectedTagPartB = null;
    let currentTagSkus = [];
    let currentTagFilter = 'all';
    let fullOptionsA = [];
    let fullOptionsB = [];
    let tagSearchQueryA = '';
    let tagSearchQueryB = '';
    let tagOverlapOnly = false;
    let tagFilesLoadedCount = 0;
    let tagLoadedFilesList = [];
    let tagDescLoadedFilesList = [];
    let tagOptionMapLoadedFilesList = [];
    let isTagInputsCollapsed = false;
    let isTagFilterCollapsed = false;
    let isTagDatabaseCollapsed = false;
    let isTagDescViewVisible = false;

    // --- Tag Matcher Logic ---
    tagToggleInputsBtn.addEventListener('click', () => {
        isTagInputsCollapsed = !isTagInputsCollapsed;
        if (tagUploadersRow) {
            tagUploadersRow.classList.toggle('collapsed', isTagInputsCollapsed);
        }

        if (isTagInputsCollapsed) {
            tagToggleIcon.setAttribute('data-lucide', 'chevron-down');
            tagToggleText.textContent = 'Restore Setup';
            showToast("Setup collapsed.");
        } else {
            tagToggleIcon.setAttribute('data-lucide', 'chevron-up');
            tagToggleText.textContent = 'Collapse Setup';
        }
        lucide.createIcons();
    });

    tagToggleFilterBtn.addEventListener('click', () => {
        isTagFilterCollapsed = !isTagFilterCollapsed;
        if (tagFilterSetupCard) {
            tagFilterSetupCard.style.display = isTagFilterCollapsed ? 'none' : 'block';
        }
        if (isTagFilterCollapsed) {
            tagFilterToggleIcon.setAttribute('data-lucide', 'eye');
            tagFilterToggleText.textContent = 'Show Filter';
        } else {
            tagFilterToggleIcon.setAttribute('data-lucide', 'eye-off');
            tagFilterToggleText.textContent = 'Hide Filter';
        }
        lucide.createIcons();
    });

    tagToggleDatabaseBtn.addEventListener('click', () => {
        isTagDatabaseCollapsed = !isTagDatabaseCollapsed;
        if (tagDescComparisonSection) {
            tagDescComparisonSection.style.display = isTagDatabaseCollapsed ? 'none' : 'block';
        }

        if (isTagDatabaseCollapsed) {
            tagDatabaseToggleIcon.setAttribute('data-lucide', 'eye');
            tagDatabaseToggleText.textContent = 'Show Product';
        } else {
            tagDatabaseToggleIcon.setAttribute('data-lucide', 'eye-off');
            tagDatabaseToggleText.textContent = 'Hide Product';
        }
        lucide.createIcons();
    });
    
    tagCsvInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        showToast(`Processing ${files.length} file(s)...`);

        // We don't clear masterTagData here so users can append files
        // masterTagData.clear(); 

        let totalSkusBefore = masterTagData.size;
        let filesProcessed = 0;

        for (const file of files) {
            await new Promise((resolve) => {
                const reader = new FileReader();
                const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

                reader.onload = (e) => {
                    try {
                        let skuParsedInFile = 0;
                        if (isExcel) {
                            const data = new Uint8Array(e.target.result);
                            const workbook = XLSX.read(data, { type: 'array' });

                            workbook.SheetNames.forEach(sheetName => {
                                const sheet = workbook.Sheets[sheetName];
                                const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                                if (rawRows.length < 2) return;

                                let skuIdx = -1, mpnIdx = -1, attrIdx = 7, headers = [], headerRowIdx = -1, foundHeader = false;

                                for (let i = 0; i < Math.min(rawRows.length, 80); i++) {
                                    const row = rawRows[i];
                                    if (!row || !Array.isArray(row)) continue;
                                    const normalized = Array.from(row).map(cell => String(cell || "").toLowerCase().replace(/\s+/g, ' ').trim());
                                    let tempSkuIdx = normalized.findIndex(s => s && (s.includes('listing') || s.includes('prsku') || s === 'sku'));
                                    let tempMpnIdx = normalized.findIndex(s => s && (s.includes('manufacturer') || s.includes('mpn') || s.includes('part number')));
                                    if (tempMpnIdx === -1) tempMpnIdx = normalized.findIndex(s => s && s.includes('part'));

                                    if (tempSkuIdx !== -1 && tempMpnIdx !== -1) {
                                        headerRowIdx = i;
                                        headers = Array.from(row).map(h => String(h || "").trim());
                                        skuIdx = tempSkuIdx;
                                        mpnIdx = tempMpnIdx;
                                        foundHeader = true;

                                        // Look ahead up to 3 rows to see if there is a human-readable header row (no '::' strings)
                                        for (let walk = 1; walk <= 3; walk++) {
                                            if (i + walk < rawRows.length) {
                                                const nextRow = rawRows[i + walk];
                                                if (nextRow && nextRow[skuIdx]) {
                                                    const testSku = String(nextRow[skuIdx]).toLowerCase();
                                                    // If this next row is clearly still a header (contains sku/listing/etc)
                                                    if (testSku.includes('sku') || testSku.includes('listing')) {
                                                        const cleanHeaders = Array.from(nextRow).map(h => String(h || "").trim());
                                                        // Only use it if it's less "code-like", meaning fewer '::'
                                                        const currentCodes = headers.filter(h => h.includes('::')).length;
                                                        const nextCodes = cleanHeaders.filter(h => h.includes('::')).length;
                                                        if (nextCodes < currentCodes) {
                                                            headers = cleanHeaders;
                                                            // We deliberately DO NOT advance headerRowIdx, because the original code
                                                            // correctly skips header rows during data extraction checking `sku.includes('sku')`
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        for (let j = 0; j < i; j++) {
                                            const topRow = rawRows[j];
                                            if (!topRow || !Array.isArray(topRow)) continue;
                                            const findAttrIdx = Array.from(topRow).findIndex(cell => cell && String(cell).toLowerCase().includes('attributes'));
                                            if (findAttrIdx !== -1) { attrIdx = findAttrIdx; break; }
                                        }
                                        break;
                                    }
                                }

                                if (foundHeader) {
                                    for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
                                        const row = rawRows[i];
                                        if (!row || !row[skuIdx]) continue;
                                        const sku = String(row[skuIdx]).trim();
                                        const mpn = String(row[mpnIdx] || "").trim();
                                        if (!sku || !mpn || sku.toLowerCase().includes('listing') || sku.toLowerCase() === 'sku' || sku.length < 3) continue;

                                        if (!masterTagData.has(sku)) masterTagData.set(sku, new Map());
                                        if (!masterTagData.get(sku).has(mpn)) masterTagData.get(sku).set(mpn, {});

                                        const tagObj = masterTagData.get(sku).get(mpn);
                                        headers.forEach((header, hIdx) => {
                                            if (hIdx < attrIdx || !header) return;
                                            const val = (row[hIdx] === undefined || row[hIdx] === null) ? "" : String(row[hIdx]).trim();
                                            tagObj[header] = val;
                                        });
                                        skuParsedInFile++;
                                    }
                                }
                            });
                        } else {
                            const content = e.target.result;
                            const lines = content.split(/\r?\n/).filter(line => line.trim());
                            if (lines.length > 0) {
                                const getParts = (line) => {
                                    if (line.includes('\t')) return line.split('\t').map(s => s.trim().replace(/^["']|["']$/g, ''));
                                    return line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.trim().replace(/^["']|["']$/g, ''));
                                };
                                let skuIdx = -1, mpnIdx = -1, attrIdx = 7, headers = [], headerIdx = -1, foundHeader = false;
                                for (let i = 0; i < Math.min(lines.length, 50); i++) {
                                    const parts = getParts(lines[i]);
                                    const normalized = parts.map(p => String(p || "").toLowerCase().replace(/\s+/g, ' ').trim());
                                    let tempSkuIdx = normalized.findIndex(s => s && (s.includes('listing') || s.includes('prsku') || s === 'sku'));
                                    let tempMpnIdx = normalized.findIndex(s => s && (s.includes('manufacturer') || s.includes('mpn') || s.includes('part number')));
                                    if (tempSkuIdx !== -1 && tempMpnIdx !== -1) {
                                        headerIdx = i; 
                                        headers = parts.map(p => String(p || "").trim()); 
                                        skuIdx = tempSkuIdx; 
                                        mpnIdx = tempMpnIdx; 
                                        foundHeader = true; 

                                        // Look ahead up to 3 rows for better headers (CSV version)
                                        for (let walk = 1; walk <= 3; walk++) {
                                            if (i + walk < lines.length) {
                                                const nextParts = getParts(lines[i + walk]);
                                                if (nextParts && nextParts[skuIdx]) {
                                                    const testSku = String(nextParts[skuIdx]).toLowerCase();
                                                    if (testSku.includes('sku') || testSku.includes('listing')) {
                                                        const cleanHeaders = nextParts.map(h => String(h || "").trim());
                                                        const currentCodes = headers.filter(h => h.includes('::')).length;
                                                        const nextCodes = cleanHeaders.filter(h => h.includes('::')).length;
                                                        if (nextCodes < currentCodes) {
                                                            headers = cleanHeaders;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        for (let j = 0; j < i; j++) {
                                            const p = getParts(lines[j]);
                                            if (!p) continue;
                                            const fIdx = p.findIndex(cell => cell && String(cell).toLowerCase().includes('attributes'));
                                            if (fIdx !== -1) { attrIdx = fIdx; break; }
                                        }
                                        break;
                                    }
                                }
                                if (foundHeader) {
                                    lines.slice(headerIdx + 1).forEach(line => {
                                        const parts = getParts(line);
                                        if (!parts[skuIdx] || !parts[mpnIdx]) return;
                                        const sku = String(parts[skuIdx]).trim();
                                        const mpn = String(parts[mpnIdx]).trim();
                                        if (!sku || !mpn || sku.toLowerCase().includes('listing') || sku.toLowerCase() === 'sku' || sku.length < 3) return;
                                        if (!masterTagData.has(sku)) masterTagData.set(sku, new Map());
                                        if (!masterTagData.get(sku).has(mpn)) masterTagData.get(sku).set(mpn, {});
                                        const tagObj = masterTagData.get(sku).get(mpn);
                                        headers.forEach((header, hIdx) => {
                                            if (hIdx < attrIdx || !header) return;
                                            const val = (parts[hIdx] === undefined || parts[hIdx] === null) ? "" : String(parts[hIdx]).trim();
                                            tagObj[header] = val;
                                        });
                                        skuParsedInFile++;
                                    });
                                }
                            }
                        }
                        filesProcessed++;
                        tagFilesLoadedCount++;
                        tagLoadedFilesList.push(file.name);
                        resolve();
                    } catch (err) {
                        console.error(`Error processing ${file.name}:`, err);
                        resolve();
                    }
                };

                if (isExcel) reader.readAsArrayBuffer(file);
                else reader.readAsText(file);
            });
        }

        if (masterTagData.size === 0) {
            showToast("Alert: No SKU/MPN columns found in any of the files.");
            tagCsvInfo.classList.remove('active');
            tagCsvInfo.style.display = 'none';
            tagCsvLabel.style.display = 'flex';
        } else {
            // Show file list line by line if multiple, or single name
            if (tagLoadedFilesList.length > 1) {
                tagCsvName.innerHTML = tagLoadedFilesList.map(name => `<div class="file-name-pill">${name}</div>`).join('');
            } else {
                tagCsvName.textContent = tagLoadedFilesList[0];
            }

            tagCsvMeta.textContent = `${masterTagData.size.toLocaleString()} SKUs parsed across ${tagFilesLoadedCount} file(s)`;
            tagCsvLabel.style.display = 'none';
            tagFileStatusArea.style.display = 'flex';
            tagCsvInfo.style.display = 'flex'; // Fix visibility
            showToast(`Total: ${masterTagData.size.toLocaleString()} SKUs from ${tagFilesLoadedCount} file(s).`);
            lookupTagOptions();
        }
        lucide.createIcons();
    });

    changeTagCsvBtn.addEventListener('click', () => {
        masterTagData.clear();
        tagFilesLoadedCount = 0;
        tagLoadedFilesList = [];
        
        tagCsvInput.value = '';
        tagCsvInfo.style.display = 'none';
        tagFileStatusArea.style.display = 'none';
        tagCsvLabel.style.display = 'flex';

        // Partial UI reset for results since data is gone
        tagResultsContainer.innerHTML = '<div class="empty-state"><p>Upload tag data and select options to see comparison.</p></div>';
        tagSkuASelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
        tagSkuBSelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
        tagSkuAHeader.textContent = 'Select SKU A Option';
        tagSkuBHeader.textContent = 'Select SKU B Option';

        showToast("Tag data reset.");
        lucide.createIcons();
    });

    tagOptionMapInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        showToast(`Processing ${files.length} variant mapping file(s)...`);

        let totalMappings = 0;
        let filesProcessed = [];

        for (const file of files) {
            await new Promise((resolve) => {
                const reader = new FileReader();
                const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

                reader.onload = (ev) => {
                    try {
                        const data = isExcel ? new Uint8Array(ev.target.result) : ev.target.result;
                        let rawRows = [];

                        if (isExcel) {
                            const workbook = XLSX.read(data, { type: 'array' });
                            rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
                        } else {
                            const lines = data.split(/\r?\n/).filter(l => l.trim());
                            rawRows = lines.map(line => {
                                if (line.includes('\t')) return line.split('\t').map(s => s.trim().replace(/^["']|["']$/g, ''));
                                return line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.trim().replace(/^["']|["']$/g, ''));
                            });
                        }

                        if (rawRows.length < 2) { resolve(); return; }

                        let skuIdx = -1, mpnIdx = -1;
                        // Variant columns: stores column indices for Grouping and Attribute
                        // Format: [ {group: idx, attr: idx}, ... ]
                        let variantPairs = []; 
                        let headerRowIdx = -1;

                        // Detect headers
                        for (let i = 0; i < Math.min(rawRows.length, 50); i++) {
                            const row = rawRows[i];
                            if (!row || !Array.isArray(row)) continue;
                            const normalized = row.map(cell => String(cell || "").toLowerCase().trim());
                            
                            let tempSkuIdx = normalized.findIndex(s => s.includes('wayfair listing') || s === 'sku' || s === 'prsku');
                            let tempMpnIdx = normalized.findIndex(s => s.includes('manufacturer part number') || s === 'mpn' || s === 'part number');

                            if (tempSkuIdx !== -1 && tempMpnIdx !== -1) {
                                headerRowIdx = i;
                                skuIdx = tempSkuIdx;
                                mpnIdx = tempMpnIdx;

                                // Look for variant column pairs
                                for (let j = 0; j < row.length - 1; j++) {
                                    const colName = normalized[j];
                                    if (colName.includes('variant grouping')) {
                                        // Usually followed by 'variant attribute ... name on site'
                                        // We look for the next column if it looks like an attribute column
                                        const nextColName = normalized[j + 1];
                                        if (nextColName.includes('variant attribute') && nextColName.includes('name on site')) {
                                            variantPairs.push({ group: j, attr: j + 1 });
                                        }
                                    }
                                }
                                break;
                            }
                        }

                        if (headerRowIdx !== -1) {
                            const headers = rawRows[headerRowIdx];
                            for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
                                const row = rawRows[i];
                                if (!row || !row[skuIdx] || !row[mpnIdx]) continue;

                                const sku = String(row[skuIdx]).trim();
                                const mpn = String(row[mpnIdx]).trim();
                                if (!sku || !mpn || sku.toLowerCase().includes('listing') || sku.toLowerCase() === 'sku') continue;

                                // Construct Option Name string: Group1: Attr1, Group2: Attr2, ...
                                let parts = new Set();
                                variantPairs.forEach(pair => {
                                    const groupName = String(row[pair.group] || "").trim();
                                    const attrValue = String(row[pair.attr] || "").trim();
                                    if (groupName && attrValue) {
                                        parts.add(`${groupName}: ${attrValue}`);
                                    }
                                });
                                
                                const partArray = Array.from(parts);
                                // Construct the name without Part Number (since it's displayed separately below)
                                // If no variants found, leave it blank (common for single-option SKUs)
                                const constructedName = partArray.length > 0 ? partArray.join(', ') : "";

                                if (!masterOptionMapData.has(sku)) masterOptionMapData.set(sku, new Map());
                                masterOptionMapData.get(sku).set(mpn, constructedName);
                                totalMappings++;
                            }
                        }
                        filesProcessed.push(file.name);
                        tagOptionMapLoadedFilesList.push(file.name);
                        resolve();
                    } catch (err) {
                        console.error("Error parsing variant map:", err);
                        resolve();
                    }
                };

                if (isExcel) reader.readAsArrayBuffer(file);
                else reader.readAsText(file);
            });
        }

        if (tagOptionMapLoadedFilesList.length > 1) {
            tagOptionMapFileName.innerHTML = tagOptionMapLoadedFilesList.map(name => `<div class="file-name-pill">${name}</div>`).join('');
        } else {
            tagOptionMapFileName.textContent = tagOptionMapLoadedFilesList[0];
        }

        tagOptionMapFileMeta.textContent = `${totalMappings.toLocaleString()} variant mappings processed.`;
        tagOptionMapLabel.style.display = 'none';
        tagOptionMapStatusArea.style.display = 'flex';
        tagOptionMapInfo.style.display = 'flex';

        showToast("Option mapping updated successfully.");
        if (currentTagSkus.length > 0) lookupTagOptions();
        lucide.createIcons();
    });

    changeTagOptionMapBtn.addEventListener('click', () => {
        masterOptionMapData.clear();
        tagOptionMapLoadedFilesList = [];
        tagOptionMapInput.value = '';
        tagOptionMapLabel.style.display = 'flex';
        tagOptionMapStatusArea.style.display = 'none';
        showToast("Option mapping cleared.");
        if (currentTagSkus.length > 0) lookupTagOptions();
    });



    // New: Handle Tag Description Database Upload
    tagDescInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        showToast(`Processing ${files.length} description file(s)...`);

        for (const file of files) {
            await new Promise((resolve) => {
                const reader = new FileReader();
                const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

                reader.onload = (ev) => {
                    try {
                        const data = isExcel ? new Uint8Array(ev.target.result) : ev.target.result;
                        let rawRows = [];

                        if (isExcel) {
                            const workbook = XLSX.read(data, { type: 'array' });
                            rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
                        } else {
                            const lines = data.split(/\r?\n/).filter(l => l.trim());
                            rawRows = lines.map(line => {
                                if (line.includes('\t')) return line.split('\t').map(s => s.trim().replace(/^["']|["']$/g, ''));
                                return line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.trim().replace(/^["']|["']$/g, ''));
                            });
                        }

                        if (rawRows.length < 2) { resolve(); return; }
                        let colMap = { sku: -1, part: -1, name: -1, desc: -1, bullets: [] };
                        let bestMatchCount = -1;
                        let headerRowIdx = -1;

                        // Robust header detection (looking at first 30 rows) - Same logic as AI Analyst
                        for (let tIdx = 0; tIdx < Math.min(rawRows.length, 30); tIdx++) {
                            const row = rawRows[tIdx];
                            if (!row || !Array.isArray(row)) continue;
                            let tempMap = { sku: -1, part: -1, name: -1, desc: -1, bullets: [] };
                            let matchCount = 0;

                            row.forEach((cell, idx) => {
                                const s = String(cell || "").toLowerCase().replace(/[^a-z0-9]/g, ' ').trim();
                                if (!s) return;
                                
                                if (s.includes('wayfair listing') || s === 'sku' || s.includes('prsku') || s.includes('listing id')) { 
                                    if (tempMap.sku === -1) { tempMap.sku = idx; matchCount++; }
                                } 
                                else if (s.includes('manufacturer part number') || s.includes('mpn') || s.includes('part number') || s.includes('model number')) { 
                                    if (tempMap.part === -1) { tempMap.part = idx; matchCount++; }
                                } 
                                else if (s.includes('marketing copy') || s.includes('description') || s.includes('copy') || s.includes('listing text') || s.includes('listing description') || s.includes('about the product')) { 
                                    if (tempMap.desc === -1) { tempMap.desc = idx; matchCount++; }
                                } 
                                else if (s.includes('bullet') || s.includes('feature')) { 
                                    if (!tempMap.bullets.includes(idx)) { tempMap.bullets.push(idx); matchCount++; }
                                } 
                                else if (s.includes('product name') || s.includes('title') || s.includes('item name') || s.includes('product title')) {
                                    if (tempMap.name === -1) { tempMap.name = idx; matchCount++; }
                                }
                            });

                            if (tempMap.sku !== -1 || tempMap.part !== -1) {
                                if (matchCount > bestMatchCount) {
                                    bestMatchCount = matchCount;
                                    headerRowIdx = tIdx;
                                    colMap = tempMap;
                                }
                            }
                        }

                        if (headerRowIdx !== -1) {
                            for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
                                const row = rawRows[i];
                                if (!row) continue;
                                const skuValue = colMap.sku !== -1 ? String(row[colMap.sku] || "").trim().toLowerCase() : "";
                                const partValue = colMap.part !== -1 ? String(row[colMap.part] || "").trim().toLowerCase() : "";
                                if (!skuValue && !partValue) continue;

                                const entry = {
                                    name: colMap.name !== -1 ? String(row[colMap.name] || "").trim() : "",
                                    desc: colMap.desc !== -1 ? String(row[colMap.desc] || "").trim() : "",
                                    bullets: colMap.bullets.map(idx => String(row[idx] || "").trim()).filter(Boolean)
                                };
                                
                                // Strict indexing logic
                                const hasData = entry.name || entry.desc || entry.bullets.length > 0;
                                if (hasData) {
                                    // ONLY save the SKU-level data if it's the specific "Wayfair SKU" master row
                                    if (skuValue && partValue === 'wayfair sku') {
                                        tagCatalogData.set(skuValue, entry);
                                    }
                                    
                                    // For other rows, index by the specific part number
                                    if (partValue && partValue !== 'wayfair sku') {
                                        tagCatalogData.set(partValue, entry);
                                    }
                                }
                            }
                        }
                        tagDescLoadedFilesList.push(file.name);
                        resolve();
                    } catch (err) {
                        console.error("Error loading descriptions:", err);
                        resolve();
                    }
                };

                if (isExcel) reader.readAsArrayBuffer(file);
                else reader.readAsText(file);
            });
        }

        if (tagDescLoadedFilesList.length > 1) {
            tagDescFileName.innerHTML = tagDescLoadedFilesList.map(name => `<div class="file-name-pill">${name}</div>`).join('');
        } else {
            tagDescFileName.textContent = tagDescLoadedFilesList[0];
        }

        tagDescFileMeta.textContent = `${tagCatalogData.size.toLocaleString()} SKUs/MPNs records processed.`;
        tagDescLabel.style.display = 'none';
        tagDescStatusArea.style.display = 'flex';
        tagDescInfo.style.display = 'flex';

        isTagDescViewVisible = true;
        tagDescComparisonSection.style.display = 'block';

        showToast("Description records updated.");
        if (currentTagSkus.length > 0) renderTagDescView(currentTagSkus[0], currentTagSkus[1]);
        lucide.createIcons();
    });


    changeTagDescBtn.addEventListener('click', () => {
        tagCatalogData.clear();
        tagDescLoadedFilesList = [];
        tagDescInput.value = '';
        tagDescLabel.style.display = 'flex';
        tagDescStatusArea.style.display = 'none';
        tagDescComparisonSection.style.display = 'none';
        showToast("Description records cleared.");
    });

    // New: Remove old toggle listener since we're making it always show when data exists


    function renderTagDescView(sku1, sku2) {
        if (!tagCatalogData || tagCatalogData.size === 0) return;
        tagDescComparisonResults.innerHTML = '';
        
        // Find best data match: try SKU first (case insensitive), then selected parts
        const findData = (sku, part) => {
            const s = (sku || "").trim().toLowerCase();
            const p = (part || "").trim().toLowerCase();
            if (s && tagCatalogData.has(s)) return tagCatalogData.get(s);
            if (p && tagCatalogData.has(p)) return tagCatalogData.get(p);
            return null;
        };


        const data1 = findData(sku1, selectedTagPartA);
        const data2 = findData(sku2, selectedTagPartB);

        const renderCol = (data, identifier) => {
            if (!data) return `<div class="glass-card"><p class="empty-text">No description data found for SKU/MPN: ${identifier || '?'}</p></div>`;

            return `
                <div class="glass-card" style="padding: 1.5rem;">
                    <h3 style="font-size: 1.1rem; color: var(--accent-main); margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 1px solid var(--card-border);">SKU ${identifier}</h3>
                    <div class="mb-1">
                        <label style="font-weight: 800; font-size: 0.8rem; color: var(--text-dim); text-transform: uppercase;">Product Name</label>
                        <p style="font-size: 0.95rem; line-height: 1.5; font-weight: 500;">${data.name || 'N/A'}</p>
                    </div>
                    <div class="mb-1">
                        <label style="font-weight: 800; font-size: 0.8rem; color: var(--text-dim); text-transform: uppercase;">Description</label>
                        <p style="font-size: 0.9rem; line-height: 1.6; padding-right: 5px;">${data.desc || 'N/A'}</p>
                    </div>
                    <div>
                        <label style="font-weight: 800; font-size: 0.8rem; color: var(--text-dim); text-transform: uppercase;">Feature Bullets</label>
                        <ul style="padding-left: 1.25rem; margin-top: 0.5rem; font-size: 0.85rem;">
                            ${data.bullets.map(b => `<li style="margin-bottom: 0.4rem; line-height: 1.4;">${b}</li>`).join('') || '<li>N/A</li>'}
                        </ul>
                    </div>
                </div>
            `;
        };

        tagDescComparisonResults.innerHTML = renderCol(data1, sku1) + renderCol(data2, sku2);
    }




    clearTagBtn.addEventListener('click', () => {
        tagSkuPairInput.value = '';
        tagResultsContainer.innerHTML = '<div class="empty-state"><p>Upload tag data and select options to see comparison.</p></div>';
        tagDescComparisonResults.innerHTML = '<div class="empty-state" style="grid-column: span 2;"><p>Select options to see description comparison.</p></div>';
        tagSkuASelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
        tagSkuBSelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
        tagSkuAHeader.textContent = 'Select SKU A Option';
        tagSkuBHeader.textContent = 'Select SKU B Option';
        tagSkuASearch.value = '';
        tagSkuBSearch.value = '';
        tagSearchQueryA = '';
        tagSearchQueryB = '';
        tagOverlapOnly = false;
        tagFilterOverlapBtn.checked = false;
        tagExportData = [];
        selectedTagPartA = null;
        selectedTagPartB = null;
        fullOptionsA = [];
        fullOptionsB = [];
        currentTagSkus = [];
        lucide.createIcons();
    });

    tagSkuPairInput.addEventListener('input', lookupTagOptions);

    // Initial load for saved filters
    const savedTagFilters = localStorage.getItem('tag_matcher_filters');
    if (savedTagFilters && tagFilterInternalInput) {
        tagFilterInternalInput.value = savedTagFilters;
    }

    if (tagSaveFilterBtn) {
        tagSaveFilterBtn.addEventListener('click', () => {
            const filters = tagFilterInternalInput.value;
            localStorage.setItem('tag_matcher_filters', filters);
            showToast("Filter tags saved to browser!");
        });
    }

    if (tagClearFilterBtn) {
        tagClearFilterBtn.addEventListener('click', () => {
            tagFilterInternalInput.value = '';
            localStorage.removeItem('tag_matcher_filters');
            performTagComparison();
            showToast("Filters cleared.");
        });
    }

    tagFilterInternalInput.addEventListener('input', performTagComparison);

    tagSkuASearch.addEventListener('input', (e) => {
        tagSearchQueryA = e.target.value.toLowerCase().trim();
        renderTagOptions('A', fullOptionsA);
    });

    tagSkuBSearch.addEventListener('input', (e) => {
        tagSearchQueryB = e.target.value.toLowerCase().trim();
        renderTagOptions('B', fullOptionsB);
    });

    tagFilterOverlapBtn.addEventListener('change', (e) => {
        tagOverlapOnly = e.target.checked;
        renderTagOptions('A', fullOptionsA);
        renderTagOptions('B', fullOptionsB);
    });

    // Filter Listeners
    document.querySelectorAll('.filter-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentTagFilter = btn.dataset.filter;
            renderTagResults(selectedTagPartA, selectedTagPartB);
        });
    });

    function lookupTagOptions() {
        // We don't return early if size is 0 because we might have manual options
        const text = tagSkuPairInput.value.trim();
        if (!text) {
            tagSkuASelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
            tagSkuBSelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
            tagSkuAHeader.textContent = 'Select SKU A Option';
            tagSkuBHeader.textContent = 'Select SKU B Option';
            return;
        }

        const skus = text.split(/[\s\t\n,]+/).map(s => s.trim().replace(/^["']|["']$/g, '')).filter(s => s.length > 0);
        if (skus.length < 1) return;

        const skuA = skus[0];
        const skuB = skus[1] || null;

        currentTagSkus = [skuA, skuB].filter(Boolean);

        if (isTagDescViewVisible) {
            renderTagDescView(skuA, skuB);
        }

        // Update Labels
        tagSkuAHeader.textContent = `Select SKU ${skuA} Option`;
        if (skuB) {
            tagSkuBHeader.textContent = `Select SKU ${skuB} Option`;
        } else {
            tagSkuBHeader.textContent = 'Select SKU B Option';
        }

        const getMergedOptions = (sku) => {
            const fromFile = masterTagData.get(sku) || new Map();
            const fromMap = masterOptionMapData.get(sku) || new Map();
            const results = new Map(); // Part -> Name

            // 1. Add from Variant Map (Highest Priority for Option Name)
            fromMap.forEach((name, mpn) => {
                results.set(mpn, name);
            });

            // 2. Add from tag file (Only if not already in results)
            fromFile.forEach((tags, mpn) => {
                if (!results.has(mpn)) {
                    // Try to find a human name in tags
                    const name = tags['Option Name'] || tags['Product Name'] || tags['Option'] || "Option (File)";
                    results.set(mpn, name);
                }
            });

            return Array.from(results.entries()).map(([part, name]) => ({ part, name }));
        };

        fullOptionsA = getMergedOptions(skuA);
        fullOptionsB = skuB ? getMergedOptions(skuB) : [];

        // --- Calculate Stats for Headers ---
        const partsA = new Set(fullOptionsA.map(o => o.part.toLowerCase().trim()));
        const partsB = new Set(fullOptionsB.map(o => o.part.toLowerCase().trim()));

        const sharedCountA = fullOptionsA.filter(o => partsB.has(o.part.toLowerCase().trim())).length;
        const sharedCountB = fullOptionsB.filter(o => partsA.has(o.part.toLowerCase().trim())).length;

        // Update Headers with Stats
        const formatHeader = (sku, total, shared) => {
            return `
                <div style="display: flex; flex-direction: column; gap: 0.4rem; margin-bottom: 0.5rem;">
                    <div style="display: flex; align-items: center; gap: 0.5rem;">
                        <span style="font-size: 1.15rem; font-weight: 800; color: var(--text-main);">Select SKU ${sku} Option</span>
                    </div>
                    <div style="display: flex; gap: 0.6rem; flex-wrap: wrap;">
                        <span class="sku-tag" style="background: var(--primary-light); color: var(--primary); font-size: 0.85rem; font-weight: 700; padding: 0.2rem 0.6rem; border-radius: 6px;">
                            ${total} options
                        </span>
                        <span class="sku-tag" style="background: var(--success-light); color: var(--success); font-size: 0.85rem; font-weight: 700; padding: 0.2rem 0.6rem; border-radius: 6px;">
                            ${shared} shared part#
                        </span>
                    </div>
                </div>
            `;
        };

        tagSkuAHeader.innerHTML = formatHeader(skuA, fullOptionsA.length, sharedCountA);
        if (skuB) {
            tagSkuBHeader.innerHTML = formatHeader(skuB, fullOptionsB.length, sharedCountB);
        } else {
            tagSkuBHeader.textContent = 'Select SKU B Option';
        }

        // Auto-select if single option and nothing selected yet
        if (fullOptionsA.length === 1 && !selectedTagPartA) selectedTagPartA = fullOptionsA[0].part;
        if (fullOptionsB.length === 1 && !selectedTagPartB) selectedTagPartB = fullOptionsB[0].part;

        renderTagOptions('A', fullOptionsA);
        renderTagOptions('B', fullOptionsB);

        if (selectedTagPartA && selectedTagPartB) performTagComparison();
    }

    function renderTagOptions(type, options) {
        const area = type === 'A' ? tagSkuASelectArea : tagSkuBSelectArea;
        const currentSku = currentTagSkus[type === 'A' ? 0 : 1] || "Unknown";
        const query = type === 'A' ? tagSearchQueryA : tagSearchQueryB;

        area.innerHTML = '';

        let filtered = options;

        // Apply "Common Parts Only" filter
        if (tagOverlapOnly && fullOptionsA.length > 0 && fullOptionsB.length > 0) {
            const partsA = new Set(fullOptionsA.map(o => o.part.toLowerCase().trim()));
            const partsB = new Set(fullOptionsB.map(o => o.part.toLowerCase().trim()));
            filtered = options.filter(opt => {
                const p = opt.part.toLowerCase().trim();
                return partsA.has(p) && partsB.has(p);
            });
        }

        if (query) {
            filtered = filtered.filter(opt =>
                opt.part.toLowerCase().includes(query) ||
                opt.name.toLowerCase().includes(query)
            );
        }

        if (filtered.length === 0) {
            if (tagOverlapOnly && options.length > 0) {
                area.innerHTML = `<p class="empty-text">No shared part numbers found${query ? ` for "${query}"` : ''}</p>`;
            } else {
                area.innerHTML = query
                    ? `<p class="empty-text">No matches for "${query}"</p>`
                    : '<p class="empty-text">No options found</p>';
            }
            return;
        }

        const partsA = new Set(fullOptionsA.map(o => o.part.toLowerCase().trim()));
        const partsB = new Set(fullOptionsB.length > 0 ? fullOptionsB.map(o => o.part.toLowerCase().trim()) : []);

        filtered.forEach(opt => {
            const mpn = opt.part;
            const name = opt.name;
            const isShared = partsA.has(mpn.toLowerCase().trim()) && partsB.has(mpn.toLowerCase().trim());
            const item = document.createElement('label');
            const isSelected = (type === 'A' ? selectedTagPartA : selectedTagPartB) === mpn;
            item.className = `option-item ${isSelected ? 'selected' : ''} ${isShared ? 'shared-option' : ''}`;
            item.innerHTML = `
                <input type="radio" name="tagSku${type}Option" value="${mpn}" ${isSelected ? 'checked' : ''}>
                <div class="option-info" style="display: flex; flex-direction: column; gap: 0.4rem; padding: 0.25rem 0;">
                    <div style="display: flex; align-items: center; gap: 0.75rem;">
                        <span class="option-name" style="font-weight: 700; color: var(--text-main); font-size: 0.95rem;">${name || 'Variant Option'}</span>
                        ${isShared ? '<span class="shared-badge">Same Part#</span>' : ''}
                    </div>
                    <div style="display: flex; align-items: center; gap: 0.75rem;">
                        <code class="option-part" style="font-family: var(--font-mono); font-size: 0.85rem; color: var(--primary); font-weight: 500; letter-spacing: -0.01em;">${mpn}</code>
                        <span style="font-size: 0.65rem; color: var(--text-dim); opacity: 0.4; font-weight: 500;">(${currentSku})</span>
                    </div>
                </div>
            `;

            item.querySelector('input').addEventListener('change', (e) => {
                const allItems = area.querySelectorAll('.option-item');
                allItems.forEach(i => i.classList.remove('selected'));
                item.classList.add('selected');

                if (type === 'A') selectedTagPartA = e.target.value;
                else selectedTagPartB = e.target.value;

                performTagComparison();
            });

            area.appendChild(item);
        });
    }



    function humanizeTagName(name) {
        if (!name) return "";
        let clean = name;

        // Handle format: core::attribute_name::123
        if (name.includes('::')) {
            const parts = name.split('::');
            // Usually the middle part is the identifier
            clean = parts[1] || parts[0];
        }

        // Remove attribute_ prefix
        if (clean.startsWith('attribute_')) {
            clean = clean.replace('attribute_', '');
        }

        // Split CamelCase or underscores
        // We only lowercase/title-case if we actually find machine-like patterns
        if (clean.includes('_') || (/[a-z][A-Z]/.test(clean)) || clean.includes('::')) {
            clean = clean.replace(/([a-z])([A-Z])/g, '$1 $2')
                .replace(/_/g, ' ')
                .toLowerCase();
            
            // Title Case
            return clean.split(' ')
                .map(word => word.charAt(0).toUpperCase() + word.slice(1))
                .join(' ');
        }

        return clean;
    }



    function performTagComparison() {
        if (!selectedTagPartA || !selectedTagPartB) return;

        const sku1 = currentTagSkus[0];
        const sku2 = currentTagSkus[1];

        const data1 = masterTagData.get(sku1)?.get(selectedTagPartA);
        const data2 = masterTagData.get(sku2)?.get(selectedTagPartB);

        if (!data1 || !data2) return;

        tagExportData = [];
        const excludedTags = [
            'total errors', 'first error', 'wayfair listing',
            'manufacturer part number', 'manufacturer part id',
            'su part number', 'supplier part number', 'existing', 'md5hash',
            'store', 'class id', 'class name',
            'attributes why', 'attributes why additional', 'is attributes edited'
        ];
        
        const userFilters = tagFilterInternalInput.value.split('\n')
            .map(t => t.trim().toLowerCase())
            .filter(Boolean);

        const allTagsList = Array.from(new Set([...Object.keys(data1), ...Object.keys(data2)]))
            .filter(tag => {
                if (!tag) return false;
                const lower = tag.toLowerCase().trim();
                const human = humanizeTagName(tag).toLowerCase();
                
                // If user entered filters, prioritize showing them even if they're in standard exclusion list
                if (userFilters.length > 0 && userFilters.some(f => lower.includes(f) || human.includes(f))) {
                    return true;
                }

                return !excludedTags.some(ex => lower.includes(ex));
            });

        // Points 2: De-duplicate by humanized name to avoid visually identical rows 
        // that come from multiple raw keys (e.g. core::width vs attribute_width)
        const seenHumanNames = new Set();
        const allTags = allTagsList.filter(tag => {
            const human = humanizeTagName(tag);
            if (seenHumanNames.has(human)) return false;
            seenHumanNames.add(human);
            return true;
        });

        allTags.forEach(tag => {
            const val1 = data1[tag] || "";
            const val2 = data2[tag] || "";
            const match = String(val1).trim() === String(val2).trim();

            tagExportData.push({
                tag,
                val1,
                val2,
                match
            });
        });

        renderTagResults(selectedTagPartA, selectedTagPartB);
        
        if (isTagDescViewVisible) {
            renderTagDescView(currentTagSkus[0], currentTagSkus[1]);
        }
    }


    function renderTagResults(label1, label2) {
        tagResultsContainer.innerHTML = '';

        if (tagExportData.length === 0) {
            tagResultsContainer.innerHTML = '<div class="empty-state"><p>No tag data found for these options.</p></div>';
            return;
        }

        // Apply filtering logic based on current tag filter
        const filteredData = tagExportData.filter(row => {
            if (currentTagFilter === 'false') return !row.match;
            
            // Only apply user keywords when the "filtered" tab is active
            if (currentTagFilter === 'filtered') {
                const userFilters = tagFilterInternalInput.value.split('\n')
                    .map(t => t.trim().toLowerCase())
                    .filter(Boolean);
                
                if (userFilters.length === 0) return false;

                const rawTag = row.tag.toLowerCase();
                const humanTag = humanizeTagName(row.tag).toLowerCase();
                
                return userFilters.some(f => rawTag.includes(f) || humanTag.includes(f));
            }
            
            return true;
        });

        if (filteredData.length === 0) {
            const filterLabel = {
                'all': 'All', 'false': 'Only False', 'filtered': 'Only Filtered'
            }[currentTagFilter];
            tagResultsContainer.innerHTML = `<div class="empty-state"><p>No rows match the "${filterLabel}" filter.</p></div>`;
            return;
        }

        let tableHtml = `
            <table class="data-table" style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                <thead>
                    <tr>
                        <th style="width: 22%; text-align: left; padding: 1rem; border-bottom: 2px solid var(--card-border); vertical-align: middle;">Tag Name</th>
                        <th style="width: 29%; text-align: left; padding: 1rem; border-bottom: 2px solid var(--card-border); vertical-align: middle;">${label1}</th>
                        <th style="width: 29%; text-align: left; padding: 1rem; border-bottom: 2px solid var(--card-border); vertical-align: middle;">${label2}</th>
                        <th style="width: 20%; text-align: center; padding: 1rem; border-bottom: 2px solid var(--card-border); vertical-align: middle;">Status</th>
                    </tr>
                </thead>
                <tbody>
        `;


        filteredData.forEach(row => {
            const statusClass = row.match ? 'success' : 'danger';
            const statusText = row.match ? 'TRUE' : 'FALSE';
            const rowClass = row.match ? '' : 'style="background: rgba(239, 68, 68, 0.03);"';

            const cleanVal1 = String(row.val1 || '').trim().replace(/^["']|["']$/g, '');
            const cleanVal2 = String(row.val2 || '').trim().replace(/^["']|["']$/g, '');

            tableHtml += `
                <tr ${rowClass}>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); font-weight: 500; font-size: 0.9rem; vertical-align: middle;">
                        ${humanizeTagName(row.tag)}
                    </td>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); font-size: 0.9rem; color: var(--text-main); vertical-align: middle;">
                        ${cleanVal1 || '<span style="color: #cbd5e1;">N/A</span>'}
                    </td>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); font-size: 0.9rem; color: var(--text-main); vertical-align: middle;">
                        ${cleanVal2 || '<span style="color: #cbd5e1;">N/A</span>'}
                    </td>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); text-align: center; vertical-align: middle;">
                        <span class="sku-tag ${statusClass}">${statusText}</span>
                    </td>
                </tr>
            `;
        });

        tableHtml += `</tbody></table>`;
        tagResultsContainer.innerHTML = tableHtml;
        lucide.createIcons();
    }



    lucide.createIcons();
    // --- Exclusion Checker Logic ---

    exclusionCsvInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        processFile(file, (content) => handleExclusionData(content, file.name));
    });

    exclusionDataModeFileBtn.addEventListener('click', () => {
        exclusionDataModeFileBtn.classList.add('active');
        exclusionDataModePasteBtn.classList.remove('active');
        exclusionDataFileArea.style.display = 'block';
        exclusionDataPasteArea.style.display = 'none';
        lucide.createIcons();
    });

    exclusionDataModePasteBtn.addEventListener('click', () => {
        exclusionDataModePasteBtn.classList.add('active');
        exclusionDataModeFileBtn.classList.remove('active');
        exclusionDataPasteArea.style.display = 'block';
        exclusionDataFileArea.style.display = 'none';
        lucide.createIcons();
    });

    processExclusionPasteBtn.addEventListener('click', () => {
        const content = exclusionDataPasteInput.value.trim();
        if (!content) {
            showToast("Please paste some data first.");
            return;
        }
        handleExclusionData(content, "Pasted Exclusion Data");
        lucide.createIcons();
    });

    function handleExclusionData(content, fileName) {
        const lines = content.split(/\r?\n/).filter(line => line.trim());
        if (lines.length === 0) return;

        const getParts = (line) => {
            if (line.includes('\t')) return line.split('\t').map(s => s.trim().replace(/^["']|["']$/g, ''));
            return line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.trim().replace(/^["']|["']$/g, ''));
        };

        const headers = getParts(lines[0]);
        const colMap = {};
        headers.forEach((h, i) => colMap[h.trim().toLowerCase()] = i);

        const getCol = (name) => {
            const lowerName = name.toLowerCase();
            return colMap[lowerName];
        };

        const skuIdx = getCol('prsku') ?? getCol('sku') ?? getCol('sku1');
        if (skuIdx === undefined) {
            showToast("Error: No SKU column (PrSKU or SKU) found in file.");
            return;
        }

        masterExclusionData.clear();
        lines.slice(1).forEach(line => {
            const parts = getParts(line);
            const sku = (parts[skuIdx] || '').trim();

            const getVal = (name) => {
                const idx = getCol(name);
                return idx !== undefined && (parts[idx] || '').toUpperCase() === 'TRUE';
            };

            if (sku) {
                masterExclusionData.set(sku, {
                    IsKitSKU: getVal('iskitsku'),
                    IsCompositeSKU: getVal('iscompositesku'),
                    Trad_Class_Excl: getVal('trad_class_excl'),
                    Trad_Supplier_Excl: getVal('trad_supplier_excl'),
                    Brand_Auth_Exclusion: getVal('brand_auth_exclusion'),
                    SKU_Exclusion: getVal('sku_exclusion'),
                    Supplier_Part_Exclusion: getVal('supplier_part_exclusion')
                });
            }
        });

        exdisplayFileName.textContent = fileName;
        exdisplayFileMeta.textContent = `${masterExclusionData.size.toLocaleString()} records loaded`;

        // UI Flow based on mode
        if (exclusionDataModeFileBtn.classList.contains('active')) {
            exclusionDataFileArea.style.display = 'none';
            exclusionCsvLabel.style.display = 'none';
        }

        exclusionCsvInfo.style.display = 'flex';
        exclusionCsvInfo.classList.add('active');
        lucide.createIcons();
    }

    changeExclusionCsvBtn.addEventListener('click', () => {
        exclusionCsvInfo.style.display = 'none';
        exclusionCsvInfo.classList.remove('active');

        // Restore based on mode
        if (exclusionDataModeFileBtn.classList.contains('active')) {
            exclusionCsvLabel.style.display = 'flex';
            exclusionDataFileArea.style.display = 'block';
            exclusionCsvInput.value = '';
        } else {
            // Paste area already visible
        }

        masterExclusionData.clear();
    });

    clearExclusionBtn.addEventListener('click', () => {
        exclusionSkuInput.value = '';
        exclusionResultsContainer.innerHTML = '<div class="empty-state"><p>Upload the exclusion database and paste SKUs to see results.</p></div>';
        downloadExclusionBtn.style.display = 'none';
        exclusionExportData = [];
    });

    processExclusionBtn.addEventListener('click', () => {
        const input = exclusionSkuInput.value.trim();
        if (!input) return;
        if (masterExclusionData.size === 0) {
            showToast("Please upload the exclusion database first.");
            return;
        }
        analyzeExclusions(input);
    });

    downloadExclusionBtn.addEventListener('click', () => {
        if (exclusionExportData.length === 0) return;

        const headers = ["SKU1", "SKU2", "Target", "Resolved", "Resolution", "Image + Tags Updated?", "Note - Details (TH AIPN - Conso)", "Reason Not Resolved", "Note - Details"];
        let csvContent = "\uFEFF"; // BOM for Excel
        csvContent += headers.map(h => `"${h}"`).join(",") + "\n";

        exclusionExportData.forEach(row => {
            let line = [];
            if (row.isClean) {
                line = [row.sku1, row.sku2, "", "", "", "", "", "", ""];
            } else {
                line = [
                    row.sku1,
                    row.sku2,
                    "", // Target
                    "No", // Resolved
                    "", // Resolution
                    "", // Image + Tags Updated?
                    "", // Note - Details (TH AIPN - Conso)
                    row.reason,
                    row.note
                ];
            }
            csvContent += line.map(val => `"${String(val || "").replace(/"/g, '""')}"`).join(",") + "\n";
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `exclusion_analysis_${new Date().toISOString().slice(0, 10)}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    function analyzeExclusions(text) {
        const lines = text.split(/\r?\n/).filter(l => l.trim());
        exclusionResultsContainer.innerHTML = '';
        exclusionExportData = [];

        lines.forEach((line, index) => {
            const parts = line.split(/[\t\s,]+/).map(s => s.trim().replace(/^["']|["']$/g, '')).filter(s => s);
            if (parts.length < 2) return;

            const sku1 = parts[0];
            const sku2 = parts[1];

            const data1 = masterExclusionData.get(sku1);
            const data2 = masterExclusionData.get(sku2);

            const result = processExclusionPair(sku1, sku2, data1, data2);
            if (result) {
                exclusionExportData.push({
                    sku1, sku2,
                    reason: result.reason || "",
                    note: result.note || "",
                    isClean: result.isClean || false
                });
                renderExclusionResult(sku1, sku2, result, index);
            }
        });

        downloadExclusionBtn.style.display = exclusionExportData.length > 0 ? 'flex' : 'none';
        lucide.createIcons();
    }

    function processExclusionPair(sku1, sku2, data1, data2) {
        // Logic priority: Right SKU (sku2) determines the main reason.
        // If both are true for different things, note combines both.

        const getIssues = (sku, data) => {
            if (!data) return [];
            const issues = [];
            if (data.IsKitSKU) issues.push({ type: 'Kit', reason: 'SKU is Kit or Composite', note: `SKU ${sku} is a KitPrSKU`, bothNote: 'Both SKUs are KitPrSKU' });
            if (data.IsCompositeSKU) issues.push({ type: 'Composite', reason: 'SKU is Kit or Composite', note: `SKU ${sku} is a CompositePrSKU`, bothNote: 'Both SKUs are CompositePrSKU' });
            if (data.Trad_Class_Excl) issues.push({ type: 'Class', reason: 'Excluded Class', note: 'Both SKUs belong to an Excluded Class (Class ID: )', bothNote: 'Both SKUs belong to an Excluded Class (Class ID: )' });
            if (data.Trad_Supplier_Excl) issues.push({ type: 'Supplier', reason: 'Excluded Supplier', note: '', bothNote: '' });
            if (data.Brand_Auth_Exclusion) issues.push({ type: 'Brand', reason: '', note: '', bothNote: '' });
            if (data.SKU_Exclusion) issues.push({ type: 'SKU', reason: 'Excluded SKU', note: `SKU ${sku} is an Excluded SKU`, bothNote: 'Both SKUs are Excluded SKU' });
            if (data.Supplier_Part_Exclusion) issues.push({ type: 'Part', reason: 'Excluded Supplier Part', note: '', bothNote: '' });
            return issues;
        };

        const issues1 = getIssues(sku1, data1);
        const issues2 = getIssues(sku2, data2);

        if (issues1.length === 0 && issues2.length === 0) {
            return { isClean: true };
        }

        let mainReason = "";
        let finalNote = "";

        // Combine notes and determine reason
        // Rule: SKU nào bên phải thì sẽ tính Reason cho SKU đó

        const allIssues = [];

        // Handle issues2 (Right SKU) first for priority in reason
        if (issues2.length > 0) {
            mainReason = issues2[0].reason;
            issues2.forEach(i2 => {
                // Check if SKU1 has the same issue type
                const commonIssue = issues1.find(i1 => i1.type === i2.type);
                if (commonIssue) {
                    allIssues.push(i2.bothNote);
                    // Remove from issues1 to avoid double counting
                    const idx = issues1.indexOf(commonIssue);
                    issues1.splice(idx, 1);
                } else {
                    allIssues.push(i2.note);
                }
            });
        }

        // Add remaining issues from SKU1
        if (issues1.length > 0) {
            if (!mainReason) mainReason = issues1[0].reason;
            issues1.forEach(i1 => {
                allIssues.push(i1.note);
            });
        }

        finalNote = allIssues.filter(n => n).join('; ');

        return { reason: mainReason, note: finalNote };
    }

    function renderExclusionResult(sku1, sku2, result, index) {
        const card = document.createElement('div');
        card.className = 'result-card';
        card.style.animation = `fadeInUp 0.5s ease-out ${index * 0.03}s both`;

        const leftContent = document.createElement('div');
        leftContent.className = 'sku-list';
        leftContent.style.flexDirection = 'row';
        leftContent.style.alignItems = 'center';
        leftContent.style.flex = '1';
        leftContent.style.gap = '1.5rem';

        const tagContainer = document.createElement('div');
        tagContainer.className = 'sku-list';
        tagContainer.style.margin = '0';
        tagContainer.style.flex = '0 0 auto';

        const combinedTag = document.createElement('span');
        combinedTag.className = 'sku-tag';
        combinedTag.style.display = 'inline-flex';
        combinedTag.style.alignItems = 'center';
        combinedTag.style.gap = '0.25rem';
        combinedTag.style.padding = '0.4rem 0.8rem';
        combinedTag.innerHTML = `
            <span>${sku1}</span>
            <span style="font-weight: bold; color: var(--text-dim); margin-top: -2px;">,</span>
            <span>${sku2}</span>
        `;
        tagContainer.appendChild(combinedTag);

        if (!result.isClean) {
            const reasonBadge = document.createElement('span');
            reasonBadge.className = 'sku-tag danger';
            reasonBadge.textContent = result.reason || 'No Reason Specified';
            tagContainer.appendChild(reasonBadge);

            const noteDiv = document.createElement('div');
            noteDiv.className = 'result-text';
            noteDiv.style.fontWeight = '500';
            noteDiv.style.padding = '0';
            noteDiv.textContent = result.note;

            leftContent.appendChild(tagContainer);
            leftContent.appendChild(noteDiv);
        } else {
            const cleanBadge = document.createElement('span');
            cleanBadge.className = 'sku-tag success';
            cleanBadge.style.background = '#f0fdf4';
            cleanBadge.style.color = '#16a34a';
            cleanBadge.textContent = 'No Exclusions';
            tagContainer.appendChild(cleanBadge);

            leftContent.appendChild(tagContainer);
        }

        card.appendChild(leftContent);

        if (!result.isClean) {
            const copyBtn = document.createElement('button');
            copyBtn.className = 'copy-btn';
            copyBtn.innerHTML = '<i data-lucide="copy"></i> Copy Note';
            copyBtn.onclick = () => {
                navigator.clipboard.writeText(result.note).then(showToast);
            };
            card.appendChild(copyBtn);
        }
        exclusionResultsContainer.appendChild(card);
    }

    // --- Toast & Utility ---
    function showToast(message = "Copied to clipboard!") {
        toast.textContent = message;
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 3000);
    }

    // --- AI Analyst Logic ---
    const aiApiKeyInput = document.getElementById('aiApiKeyInput');
    const toggleApiKeyBtn = document.getElementById('toggleApiKeyBtn');
    const modelCards = document.querySelectorAll('.model-card');

    const aiCsvInput = document.getElementById('aiCsvInput');
    const aiCsvInfo = document.getElementById('aiCsvInfo');
    const aiCsvLabel = document.getElementById('aiCsvLabel');
    const aiCsvName = document.getElementById('aiCsvName');
    const aiCsvMeta = document.getElementById('aiCsvMeta');
    const changeAiCsvBtn = document.getElementById('changeAiCsvBtn');
    const aiPromptInput = document.getElementById('aiPromptInput');
    const aiBatchInput = document.getElementById('aiBatchInput');
    const aiProcessBatchBtn = document.getElementById('aiProcessBatchBtn');
    const aiProgressArea = document.getElementById('aiProgressArea');
    const aiProgressBar = document.getElementById('aiProgressBar');
    const aiProgressCount = document.getElementById('aiProgressCount');
    const aiProgressText = document.getElementById('aiProgressText');
    const aiErrorLogArea = document.getElementById('aiErrorLogArea');
    const aiErrorLogList = document.getElementById('aiErrorLogList');
    const aiQueryInput = document.getElementById('aiQueryInput');
    const aiQueryBtn = document.getElementById('aiQueryBtn');
    const aiQueryClearBtn = document.getElementById('aiQueryClearBtn');
    const aiQueryResult = document.getElementById('aiQueryResult');

    let aiCatalogData = new Map(); // PrSKU (Listing) -> { sku, name, desc, bullets: [] }
    let mpnToPrSku = new Map();     // MPN -> PrSKU (Listing)
    let aiResultsDB = new Map();   // "PrSKU-A|PrSKU-B" -> String (AI Response)
    let isAiProcessing = false;
    let isAiCancelled = false;
    const aiStopBatchBtn = document.getElementById('aiStopBatchBtn');

    if (aiStopBatchBtn) {
        aiStopBatchBtn.addEventListener('click', () => {
            if (isAiProcessing) {
                isAiCancelled = true;
                aiStopBatchBtn.innerHTML = '<i data-lucide="loader" class="spin"></i> Stopping...';
                if (window.lucide) lucide.createIcons();
                showToast("Stopping analysis...");
            }
        });
    }

    // --- API Key & Model Persistence ---
    const STORAGE_KEYS = {
        API_KEY: 'gemini_api_key',
        MODEL: 'gemini_selected_model'
    };

    // Load saved settings
    const savedApiKey = localStorage.getItem(STORAGE_KEYS.API_KEY);
    const savedModel = localStorage.getItem(STORAGE_KEYS.MODEL) || 'gemini-3.1-pro-preview';

    if (savedApiKey) aiApiKeyInput.value = savedApiKey;

    modelCards.forEach(card => {
        if (card.dataset.model === savedModel) {
            modelCards.forEach(c => c.classList.remove('active'));
            card.classList.add('active');
        }

        card.addEventListener('click', () => {
            modelCards.forEach(c => c.classList.remove('active'));
            card.classList.add('active');
            localStorage.setItem(STORAGE_KEYS.MODEL, card.dataset.model);
            showToast(`Selected model: ${card.querySelector('h4').textContent}`);
        });
    });

    aiApiKeyInput.addEventListener('input', (e) => {
        localStorage.setItem(STORAGE_KEYS.API_KEY, e.target.value);
    });

    toggleApiKeyBtn.addEventListener('click', () => {
        const type = aiApiKeyInput.type === 'password' ? 'text' : 'password';
        aiApiKeyInput.type = type;
        toggleApiKeyBtn.innerHTML = type === 'password' ? '<i data-lucide="eye"></i>' : '<i data-lucide="eye-off"></i>';
        lucide.createIcons();
    });

    // --- IndexedDB Persistence Engine ---
    const DB_NAME = 'SKUAnalyzerAI';
    const STORE_NAME = 'analysis_results';

    function initIndexedDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, 1);
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME);
                }
            };
            request.onsuccess = (e) => resolve(e.target.result);
            request.onerror = (e) => reject(e.target.error);
        });
    }

    async function saveResultToDB(key, value) {
        const db = await initIndexedDB();
        const tx = db.transaction(STORE_NAME, 'readwrite');
        tx.objectStore(STORE_NAME).put(value, key);
        return tx.complete;
    }

    async function loadResultsFromDB() {
        const db = await initIndexedDB();
        return new Promise((resolve) => {
            const tx = db.transaction(STORE_NAME, 'readonly');
            const store = tx.objectStore(STORE_NAME);
            const request = store.getAll();
            const keysRequest = store.getAllKeys();

            request.onsuccess = () => {
                keysRequest.onsuccess = () => {
                    const results = new Map();
                    request.result.forEach((val, i) => {
                        results.set(keysRequest.result[i], val);
                    });
                    resolve(results);
                };
            };
        });
    }

    // Load existing results on startup
    loadResultsFromDB().then(savedResults => {
        aiResultsDB = savedResults;
        console.log(`Loaded ${aiResultsDB.size} cached AI results from permanent storage.`);
    });

    // Default Prompt
    const defaultAiPrompt = `Persona: Professional Wayfair Data Specialist
You are a precise, expert analyst specializing in Wayfair product listings. Speak in clear, professional English. Your role is to compare TWO SKUs (products) from Wayfair data and calculate the similarity percentage. Always refer to the products by their actual SKU IDs provided in the input, prefixed with "SKU " and in ALL UPPERCASE (e.g., if ID is cbzr1097, write SKU CBZR1097).

Output a single similarity score (0-100%) based on identical or near-identical content. Differences in color, finish, dimensions, or weight count as "option variants" (not core differences). Flag true separators only if they impact:
- Design/Style
- Material
- Functions
- Patterns
- Quantity

For every distinct pair found in the data, generate a separate analysis.

## Step-by-Step Process
1. **Compare Sections**: Analyze word-for-word similarity per section (e.g., 90% match if phrasing identical except minor words).
2. **Calculate Overall %**: Average the 3 sections' similarity (e.g., Names 100%, Desc 80%, Bullets 90% = 90% total). Explain math briefly.
3. **Differences**:
   - **Core Separators** (if any): List with evidence.
   - **Option Variants** (if any): List differences.
   - If identical: State "No separating differences; potential duplicates."

## Output Format (Always Use This Template)
**Similarity Score: [X]%**

**Section Breakdown:**
| Section | Similarity % |
|---------|--------------|
| Product Name | [X]% |
| Description | [X]% |
| Feature Bullets | [X]% |

**Core Separators:** [List or "None"]

**Option Variants:** [List or "None"]

**Recommendation:** [Merge/Duplicate/Variants/Separate + 1-sentence reason + actual SKU IDs in "SKU [ID_UPPERCASE]" format].

Example: If names match 100%, desc 85% (minor rephrase), bullets 95% (same features), variants in color: Score 93%. Variants: Color. Recommendation: Variants under one parent.`;
    aiPromptInput.value = defaultAiPrompt;

    // Load AI Catalog
    aiCsvInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        showToast("Reading AI database...");
        const reader = new FileReader();
        const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

        reader.onload = (ev) => {
            try {
                aiCatalogData.clear();

                const processRawRows = (rawRows) => {
                    if (rawRows.length < 2) return;
                    let colMap = { sku: -1, part: -1, name: -1, desc: -1, bullets: [] };
                    let bestMatchCount = -1;
                    let headerRowIdx = -1;

                    // Scan first 15 rows for the best header row (robust detection)
                    for (let i = 0; i < Math.min(rawRows.length, 15); i++) {
                        const row = rawRows[i];
                        if (!row || !Array.isArray(row)) continue;
                        
                        let tempMap = { sku: -1, part: -1, name: -1, desc: -1, bullets: [] };
                        let matchCount = 0;
                        
                        row.forEach((cell, idx) => {
                            const s = String(cell || "").toLowerCase().replace(/\s+/g, ' ').trim();
                            if (!s) return;
                            
                            // Essential: Listing ID
                            if (s.includes('listing') || s.includes('prsku') || s === 'sku') { 
                                if (tempMap.sku === -1) { tempMap.sku = idx; matchCount++; }
                            } 
                            // Essential: MPN
                            else if (s.includes('manufacturer') || s.includes('mpn') || s.includes('part number')) { 
                                if (tempMap.part === -1) { tempMap.part = idx; matchCount++; }
                            } 
                            // Marketing Description
                            else if (s.includes('marketing copy') || s.includes('description') || s.includes('copy')) { 
                                if (tempMap.desc === -1) { tempMap.desc = idx; matchCount++; }
                            } 
                            // Feature Bullets (Can be multiple columns)
                            else if (s.includes('bullet') || s.includes('feature')) { 
                                if (!tempMap.bullets.includes(idx)) {
                                    tempMap.bullets.push(idx); 
                                    matchCount++;
                                }
                            } 
                            // Product Name
                            else if (s === 'product name') {
                                if (tempMap.name === -1) { tempMap.name = idx; matchCount++; }
                            }
                            else if ((s.includes('name') || s.includes('title') || s.includes('product')) && !s.includes('name on site')) { 
                                if (tempMap.name === -1) { tempMap.name = idx; matchCount++; }
                            }
                        });

                        // Prioritize rows that have BOTH essential columns
                        if (tempMap.sku !== -1 && tempMap.part !== -1) {
                            // Weight essential matches higher or just use total count if essentials are met
                            if (matchCount > bestMatchCount) {
                                bestMatchCount = matchCount;
                                headerRowIdx = i;
                                colMap = tempMap;
                            }
                        }
                    }

                    if (headerRowIdx === -1) {
                        console.error("AI Analyst - Error: Required columns (SKU/Listing and MPN/Part Number) not found in the first 15 rows.");
                        showToast("Error: Could not find required headers in Excel file.");
                    } else {
                        console.log("AI Analyst - Selected Header Row", headerRowIdx, "with", bestMatchCount, "mappings:", colMap);
                        let fullDataCount = 0;
                        for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
                            const row = rawRows[i];
                            if (!row || row[colMap.sku] === undefined) continue;
                            
                            const prSkuRaw = String(row[colMap.sku]).trim();
                            const prSku = prSkuRaw.toLowerCase();
                            if (!prSku) continue;

                            const partNumRaw = String(row[colMap.part] || "").trim();
                            const partNum = partNumRaw.replace(/\s+/g, ' ').toLowerCase();

                            // Mapping actual MPNs to the parent listing ID
                            if (partNumRaw && partNum !== "wayfair sku") {
                                mpnToPrSku.set(partNum, prSku);
                            }

                            // If this row is the "Parent" (Wayfair SKU row), extract marketing content
                            if (partNum === "wayfair sku" || partNum.includes("wayfair sku")) {
                                let allBullets = [];
                                colMap.bullets.forEach(bIdx => {
                                    const raw = String(row[bIdx] || "").trim().replace(/^["']+|["']+$/g, '');
                                    if (raw) {
                                        // Split by newline in case a cell contains multiple bullets
                                        const lines = raw.split(/\r?\n/).map(l => l.trim().replace(/^[-*•]\s*/, '')).filter(l => l);
                                        allBullets = allBullets.concat(lines);
                                    }
                                });
                                // Uniquify bullets
                                const uniqueBullets = [...new Set(allBullets)]
                                    .filter(b => b.toLowerCase() !== "text" && b.toLowerCase() !== "conditional");

                                aiCatalogData.set(prSku, {
                                    sku: prSkuRaw, // Keep original casing for display
                                    name: colMap.name !== -1 ? String(row[colMap.name] || "").trim().replace(/^["']+|["']+$/g, '') : "",
                                    desc: colMap.desc !== -1 ? String(row[colMap.desc] || "").trim().replace(/^["']+|["']+$/g, '') : "",
                                    bullets: uniqueBullets
                                });
                                fullDataCount++;
                            }
                        }
                        console.log(`AI Analyst - Extracted content for ${fullDataCount} listings. Mapped ${mpnToPrSku.size} MPNs.`);
                    }
                };

                if (isExcel) {
                    const data = new Uint8Array(ev.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    workbook.SheetNames.forEach(sheetName => {
                        const sheet = workbook.Sheets[sheetName];
                        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                        processRawRows(rawRows);
                    });
                } else {
                    showToast("Only Excel files (.xlsx) are supported for AI Database due to formatting.");
                    return;
                }

                aiCsvName.textContent = file.name;
                aiCsvMeta.textContent = `${aiCatalogData.size.toLocaleString()} SKUs extracted`;
                aiCsvLabel.style.display = 'none';
                aiCsvInfo.style.display = 'flex';
                showToast(`Success! Extracted ${aiCatalogData.size} target SKUs.`);
            } catch (err) {
                showToast(`Error processing file: ${err.message}`);
                console.error(err);
            }
        };

        if (isExcel) reader.readAsArrayBuffer(file);
        else reader.readAsText(file);
    });

    changeAiCsvBtn.addEventListener('click', () => {
        aiCsvInfo.style.display = 'none';
        aiCsvLabel.style.display = 'flex';
        aiCsvInput.value = '';
        aiCatalogData.clear();
    });

    // Generate pair key
    const getPairKey = (s1, s2) => {
        const sorted = [s1.toLowerCase().trim(), s2.toLowerCase().trim()].sort();
        return sorted.join('|');
    };

    // Call Gemini API
    async function callGeminiApi(promptText) {
        const apiKey = aiApiKeyInput.value.trim();
        const activeModelCard = document.querySelector('.model-card.active');
        const model = activeModelCard ? activeModelCard.dataset.model : 'gemini-3.1-pro-preview';

        if (!apiKey) {
            throw new Error("Missing API Key. Please provide one in the Configuration section.");
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

        const payload = {
            contents: [{ parts: [{ text: promptText }] }],
            generationConfig: { temperature: 0.2 }
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errData = await response.json();
            throw new Error(errData.error?.message || "API Error");
        }

        const data = await response.json();
        return data.candidates[0].content.parts[0].text;
    }

    // Process Batch
    aiProcessBatchBtn.addEventListener('click', async () => {
        if (isAiProcessing) return;

        if (aiCatalogData.size === 0) {
            showToast("Please upload the Product Database first.");
            return;
        }

        const text = aiBatchInput.value.trim();
        if (!text) {
            showToast("Please paste some SKU pairs to analyze");
            return;
        }

        const lines = text.split(/\r?\n/).filter(l => l.trim());
        const pairsToProcess = [];
        lines.forEach(line => {
            const skus = line.split(/[\t\s,]+/).map(s => s.trim().replace(/^["']|["']$/g, '')).filter(s => s);
            if (skus.length >= 2) {
                pairsToProcess.push({ sku1: skus[0], sku2: skus[1] });
            }
        });

        if (pairsToProcess.length === 0) {
            showToast("No valid pairs found to analyze.");
            return;
        }

        // --- NEW: Check for existing results to avoid accidental costs ---
        let existingPairsCount = 0;
        pairsToProcess.forEach(pair => {
            const s1 = mpnToPrSku.get(pair.sku1.toLowerCase().trim()) || pair.sku1.toLowerCase().trim();
            const s2 = mpnToPrSku.get(pair.sku2.toLowerCase().trim()) || pair.sku2.toLowerCase().trim();
            if (aiResultsDB.has(getPairKey(s1, s2))) {
                existingPairsCount++;
            }
        });

        let forceReanalyze = false;
        if (existingPairsCount > 0) {
            forceReanalyze = confirm(`Phát hiện ${existingPairsCount} cặp SKU đã được phân tích trước đó.\n\nBạn có muốn PHÂN TÍCH LẠI (Re-analyze) chúng không?\n- Chọn 'OK': Sẽ phân tích lại tất cả (Tốn thêm tiền/token AI).\n- Chọn 'Cancel': Sẽ bỏ qua các cặp cũ, chỉ phân tích các cặp mới.`);
        }
        // -----------------------------------------------------------------

        isAiProcessing = true;
        isAiCancelled = false;
        aiProcessBatchBtn.innerHTML = '<i data-lucide="loader" class="spin"></i> Processing...';
        if (aiStopBatchBtn) {
            aiStopBatchBtn.style.display = 'flex';
            aiStopBatchBtn.innerHTML = '<i data-lucide="x-octagon"></i> Stop';
        }
        if (window.lucide) lucide.createIcons();
        aiProgressArea.style.display = 'block';
        aiProgressText.textContent = "Analyzing...";
        aiErrorLogArea.style.display = 'none';
        aiErrorLogList.innerHTML = '';
        updateAiProgress(0, pairsToProcess.length);
        let completed = 0;

        const systemPrompt = aiPromptInput.value.trim();

        for (const pair of pairsToProcess) {
            if (isAiCancelled) {
                break;
            }

            const input1 = pair.sku1.toLowerCase().trim();
            const input2 = pair.sku2.toLowerCase().trim();

            // Resolve to Listing IDs (PrSKUs)
            const s1 = mpnToPrSku.get(input1) || input1;
            const s2 = mpnToPrSku.get(input2) || input2;

            const key = getPairKey(s1, s2);

            const data1 = aiCatalogData.get(s1);
            const data2 = aiCatalogData.get(s2);

            if (!data1 || !data2) {
                if (isAiCancelled) break; 
                const missing = [];
                if (!data1) missing.push(input1);
                if (!data2) missing.push(input2);

                aiResultsDB.set(key, `**[Error]** Some SKUs not found in the Database.\nMissing Info for: ${missing.join(', ')}`);
                aiErrorLogArea.style.display = 'block';
                const li = document.createElement('li');
                li.innerHTML = `<strong>${input1} - ${input2}:</strong> Missing Data`;
                aiErrorLogList.appendChild(li);
                completed++;
                updateAiProgress(completed, pairsToProcess.length);
                continue;
            }

            const cleanText = (str) => String(str || "").replace(/\s{2,}/g, ' ').trim();

            const promptText = `
${systemPrompt.trim()}

### INPUT DATA:
SKU A ID: ${data1.sku}
Product Name: ${cleanText(data1.name)}
Description: ${cleanText(data1.desc)}
Feature Bullets:
${data1.bullets.map(b => '- ' + cleanText(b)).join('\n')}

SKU B ID: ${data2.sku}
Product Name: ${cleanText(data2.name)}
Description: ${cleanText(data2.desc)}
Feature Bullets:
${data2.bullets.map(b => '- ' + cleanText(b)).join('\n')}

REMINDER: In your response, replace "[SKU_ID_1_UPPER]" with "SKU ${data1.sku.toUpperCase()}" and "[SKU_ID_2_UPPER]" with "SKU ${data2.sku.toUpperCase()}". Throughout the analysis, always prefix SKU IDs with "SKU " and write them in ALL CAPS.
`.replace(/\n{3,}/g, '\n\n').trim();



            // Skip if exists unless forceReanalyze is true
            if (aiResultsDB.has(key) && !forceReanalyze) {
                completed++;
                updateAiProgress(completed, pairsToProcess.length);
                continue;
            }

            try {
                await new Promise(r => setTimeout(r, 1000));
                if (isAiCancelled) break;
                const responseText = await callGeminiApi(promptText);
                if (isAiCancelled) break;
                aiResultsDB.set(key, responseText);
                await saveResultToDB(key, responseText); // Save to permanent storage
            } catch (err) {
                if (isAiCancelled) break;
                const errMsg = err.message || "Unknown API error";
                aiErrorLogArea.style.display = 'block';
                const li = document.createElement('li');
                li.innerHTML = `<strong>${input1} - ${input2}:</strong> ${errMsg}`;
                aiErrorLogList.appendChild(li);

                if (errMsg.includes("API_KEY_INVALID")) {
                    showToast("Invalid API Key. Terminating batch.");
                    break;
                }
            }

            completed++;
            updateAiProgress(completed, pairsToProcess.length);
        }

        aiProgressText.textContent = isAiCancelled ? "Stopped" : "Completed";
        isAiProcessing = false;
        aiProcessBatchBtn.innerHTML = '<i data-lucide="bot"></i> Analyze with AI';
        if (aiStopBatchBtn) aiStopBatchBtn.style.display = 'none';
        showToast(isAiCancelled ? "Analysis stopped." : "Batch analysis completed.");
        if (isAiCancelled) isAiCancelled = false;
        if (window.lucide) lucide.createIcons();
    });

    function updateAiProgress(completed, total) {
        aiProgressCount.textContent = `${completed} / ${total}`;
        aiProgressBar.style.width = `${(completed / total) * 100}%`;
    }

    // Query Result
    aiQueryBtn.addEventListener('click', () => {
        const text = aiQueryInput.value.trim();
        if (!text) return;

        // Take the first pair found
        const skus = text.split(/[\t\s,]+/).map(s => s.trim()).filter(s => s);
        if (skus.length < 2) {
            aiQueryResult.innerHTML = '<p class="empty-text">Please enter a valid pair (e.g. SKU1, SKU2)</p>';
            return;
        }

        const input1 = skus[0].toLowerCase().trim();
        const input2 = skus[1].toLowerCase().trim();

        const s1 = mpnToPrSku.get(input1) || input1;
        const s2 = mpnToPrSku.get(input2) || input2;

        const key = getPairKey(s1, s2);
        const d1 = aiCatalogData.get(s1);
        const d2 = aiCatalogData.get(s2);

        aiQueryResult.innerHTML = `<div style="margin-bottom:15px; display:flex; gap:0.5rem; align-items:center;">
             <span class="sku-tag uppercase-exception" style="background:#8b5cf6;color:white;font-size:1rem;padding:0.4rem 0.8rem;">${input1}</span> 
             <span style="font-weight:bold;color:var(--text-dim);">VS</span> 
             <span class="sku-tag uppercase-exception" style="background:#ec4899;color:white;font-size:1rem;padding:0.4rem 0.8rem;">${input2}</span>
        </div>`;

        const getOrigText = (section, origData) => {
            if (!origData) return "<em style='color:#94a3b8;'>[Data not loaded]</em>";
            if (section.toLowerCase().includes('name')) return origData.name || "";
            if (section.toLowerCase().includes('desc')) return origData.desc || "";
            if (section.toLowerCase().includes('bullet')) {
                if (!origData.bullets || origData.bullets.length === 0) return "";
                return "- " + origData.bullets.join("<br/>- ");
            }
            return "";
        };

        if (aiResultsDB.has(key)) {
            const rawText = aiResultsDB.get(key);

            let htmlChunks = [];
            const lines = rawText.split('\n');
            let inTable = false;
            let tableHtml = "";

            lines.forEach(line => {
                if (line.trim().startsWith('|')) {
                    if (!inTable) { 
                        inTable = true; 
                        tableHtml = "<div style='overflow-x: auto; margin: 15px 0;'><table style='border-collapse: collapse; width: 100%; font-size: 0.85rem;'><tbody>"; 
                        // Override header row completely
                        if (line.toLowerCase().includes('section')) {
                            tableHtml += `<tr>
                                <td style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt);">Section</td>
                                <td style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt);">SKU ${input1.toUpperCase()} Text</td>
                                <td style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt);">SKU ${input2.toUpperCase()} Text</td>
                                <td style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt);">Similarity %</td>
                            </tr>`;
                            return;
                        }
                    }
                    if (line.includes('---------')) return; // skip separator

                    const rowCells = line.split('|').map(s => s.trim()).filter((s, i, arr) => i > 0 && i < arr.length - 1);
                    const cleanCells = rowCells.map(c => c.trim().replace(/^["']+|["']+$/g, ''));
                    
                    if (cleanCells.length === 2 && !line.toLowerCase().includes('section')) {
                        // Dynamically inject the 2 middle columns using local data
                        const sectionName = cleanCells[0];
                        const simScore = cleanCells[1];
                        
                        tableHtml += `<tr>
                            <td style="border: 1px solid var(--card-border); padding: 8px;">${sectionName}</td>
                            <td style="border: 1px solid var(--card-border); padding: 8px;">${getOrigText(sectionName, d1)}</td>
                            <td style="border: 1px solid var(--card-border); padding: 8px;">${getOrigText(sectionName, d2)}</td>
                            <td style="border: 1px solid var(--card-border); padding: 8px; text-align: center; font-weight: 500;">${simScore}</td>
                        </tr>`;
                    } else if (cleanCells.length > 2 && !line.toLowerCase().includes('section')) {
                         // Fallback in case old cached data with 4 columns is queried
                         tableHtml += "<tr>" + cleanCells.map(c => `<td style="border: 1px solid var(--card-border); padding: 8px;">${c}</td>`).join('') + "</tr>";
                    }
                } else {
                    if (inTable) {
                        inTable = false;
                        htmlChunks.push(tableHtml + "</tbody></table></div>");
                    }
                    htmlChunks.push(line);
                }
            });
            if (inTable) htmlChunks.push(tableHtml + "</tbody></table></div>");

            let formatted = htmlChunks.join('\n')
                .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
                .replace(/\*(.*?)\*/g, '<em>$1</em>');

            const resultDiv = document.createElement('div');
            resultDiv.innerHTML = formatted;
            aiQueryResult.appendChild(resultDiv);

        } else {
            const sections = ["Product Name", "Description", "Feature Bullets"];
            let tableHtml = `<p class="empty-text" style="color:#ef4444; font-weight:600; text-align:left; margin-bottom:10px;">Chưa được phân tích (Not yet analyzed).</p>
                <div style='overflow-x: auto; margin: 15px 0;'>
                <table style='border-collapse: collapse; width: 100%; font-size: 0.85rem;'>
                <thead>
                    <tr>
                        <th style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt); text-align:left;">Section</th>
                        <th style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt); text-align:left;">SKU ${input1.toUpperCase()} Text</th>
                        <th style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt); text-align:left;">SKU ${input2.toUpperCase()} Text</th>
                        <th style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt); text-align:left;">Similarity %</th>
                    </tr>
                </thead>
                <tbody>`;
            
            sections.forEach(sec => {
                tableHtml += `<tr>
                    <td style="border: 1px solid var(--card-border); padding: 8px; font-weight: bold; background: var(--bg-alt);">${sec}</td>
                    <td style="border: 1px solid var(--card-border); padding: 8px;">${getOrigText(sec, d1)}</td>
                    <td style="border: 1px solid var(--card-border); padding: 8px;">${getOrigText(sec, d2)}</td>
                    <td style="border: 1px solid var(--card-border); padding: 8px; text-align:center;">-</td>
                </tr>`;
            });

            tableHtml += "</tbody></table></div>";
            const resultDiv = document.createElement('div');
            resultDiv.innerHTML = tableHtml;
            aiQueryResult.appendChild(resultDiv);
        }
    });

    aiQueryClearBtn.addEventListener('click', () => {
        aiQueryInput.value = '';
        aiQueryResult.innerHTML = '<p class="empty-text">Result will appear here</p>';
    });

    // --- Export/Import Logic ---
    const aiExportDBBtn = document.getElementById('aiExportDBBtn');
    const aiImportDBBtn = document.getElementById('aiImportDBBtn');
    const aiImportFile = document.getElementById('aiImportFile');
    const aiClearDBBtn = document.getElementById('aiClearDBBtn');

    aiExportDBBtn.addEventListener('click', () => {
        if (aiResultsDB.size === 0) {
            showToast("Database is empty. Nothing to export.");
            return;
        }
        const data = Object.fromEntries(aiResultsDB);
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `sku_ai_results_backup_${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        showToast("Database exported successfully!");
    });

    aiImportDBBtn.addEventListener('click', () => aiImportFile.click());

    aiClearDBBtn.addEventListener('click', async () => {
        if (aiResultsDB.size === 0) {
            showToast("Database is already empty.");
            return;
        }

        if (confirm(`Are you sure you want to CLEAR all ${aiResultsDB.size} analysis results? This cannot be undone unless you have a backup.`)) {
            const db = await initIndexedDB();
            const tx = db.transaction(STORE_NAME, 'readwrite');
            await tx.objectStore(STORE_NAME).clear();
            aiResultsDB.clear();
            showToast("Database cleared successfully!");
            aiQueryResult.innerHTML = '<p class="empty-text">Result will appear here</p>';
        }
    });

    aiImportFile.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = async (ev) => {
            try {
                const data = JSON.parse(ev.target.result);
                let importedCount = 0;
                let overwrittenCount = 0;
                for (const [key, value] of Object.entries(data)) {
                    if (aiResultsDB.has(key)) overwrittenCount++;
                    aiResultsDB.set(key, value);
                    await saveResultToDB(key, value);
                    importedCount++;
                }
                showToast(`Imported ${importedCount} results (${overwrittenCount} overwritten). Total cached: ${aiResultsDB.size}`);
                aiImportFile.value = ''; // Reset
            } catch (err) {
                showToast("Error: Invalid JSON file.");
                console.error(err);
            }
        };
        reader.readAsText(file);
    });

    // --- Toast & Utility ---
    function showToast(message = "Copied to clipboard!") {
        toast.textContent = message;
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 3000);
    }

    // Initialize Icons
    if (window.lucide) {
        lucide.createIcons();
    }
});

