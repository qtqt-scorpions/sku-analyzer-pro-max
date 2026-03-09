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
    let analyzerMode = 'upload'; // 'upload' or 'paste'

    let masterCOOData = new Map(); // SKU -> Country of Origin
    let cooExportData = [];

    // --- Exclusion Checker Elements ---
    const exclusionCsvInput = document.getElementById('exclusionCsvInput');
    const exclusionCsvLabel = document.getElementById('exclusionCsvLabel');
    const exclusionCsvInfo = document.getElementById('exclusionCsvInfo');
    const exdisplayFileName = document.getElementById('exdisplayFileName');
    const exdisplayFileMeta = document.getElementById('exdisplayFileMeta');
    const changeExclusionCsvBtn = document.getElementById('changeExclusionCsvBtn');
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
            const skus = line.split(/[\t\s,]+/).map(s => s.trim()).filter(s => s && s.toLowerCase() !== 'sku1' && s.toLowerCase() !== 'sku2');
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

        processFile(file, (content) => {
            const lines = content.split(/\r?\n/);
            masterData.clear();
            lines.forEach((line, index) => {
                if (index === 0) return;
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
            displayFileName.textContent = file.name;
            displayFileMeta.textContent = `${masterData.size.toLocaleString()} records loaded`;
            fileLabel.style.display = 'none';
            fileInfoCard.classList.add('active');
            lucide.createIcons();
        });
    });

    changeFileBtn.addEventListener('click', () => {
        fileInfoCard.classList.remove('active');
        fileLabel.style.display = 'block';
        masterFileInput.value = '';
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
        const headers = ["Target", "Resolved?", "Resolution", "Image + Tags Updated?", "Note - Details (TH AIPN - Conso)", "Reason Not Resolved", "Note - Details"];

        let csvContent = headers.join(",") + "\n";

        analyzerExportData.forEach(row => {
            const line = [
                `"${row.resolved === 'Yes' ? row.target : ''}"`,
                `"${row.resolved}"`,
                `"${row.resolved === 'Yes' ? row.resolution : ''}"`,
                `""`, // Image + Tags Updated?
                `"${row.resolved === 'Yes' ? row.note : ''}"`, // AIPN Note
                `""`, // Reason Not Resolved
                `"${row.resolved === 'No' ? row.note : ''}"`  // Note Details
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
                    const txt = `Source [${sList}], Target [${tList}] - No common destination found`;
                    return {
                        tags: [{ text: 'No', type: 'danger' }, { text: 'No common target', type: 'danger' }],
                        sentence: "Both SKUs don't have the same new target",
                        copy: txt, isError: true, warnings,
                        export: { target, resolved: 'No', resolution: '', note: "Both SKUs don't have the same new target" }
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
    const matcherModeLookupBtn = document.getElementById('matcherModeLookupBtn');
    const matcherModePasteBtn = document.getElementById('matcherModePasteBtn');
    const matcherLookupArea = document.getElementById('matcherLookupArea');
    const matcherPasteArea = document.getElementById('matcherPasteArea');
    const skuASelectArea = document.getElementById('skuASelectArea');
    const skuBSelectArea = document.getElementById('skuBSelectArea');
    const matcherSkuASearch = document.getElementById('matcherSkuASearch');
    const matcherSkuBSearch = document.getElementById('matcherSkuBSearch');

    let imageRows = [];
    let selectedPartA = null;
    let selectedPartB = null;
    let matcherSortCol = 'PrSKU';
    let matcherSortDir = 'asc';
    let allPartsA = new Set();
    let allPartsB = new Set();
    let matcherInputMode = 'lookup'; // 'lookup' or 'paste'
    let fullRawOptionsA = []; // Array of {name, part, sku}
    let fullRawOptionsB = [];
    let matcherSearchQueryA = '';
    let matcherSearchQueryB = '';

    // --- Option Matcher Logic ---
    matcherCsvInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        processFile(file, (content) => {
            const lines = content.split(/\r?\n/).filter(line => line.trim());
            if (lines.length === 0) return;

            const getParts = (line) => {
                if (line.includes('\t')) return line.split('\t').map(s => s.trim().replace(/^["']|["']$/g, ''));
                return line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.trim().replace(/^["']|["']$/g, ''));
            };

            // --- Better Header Detection ---
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

            // --- Validation: Check for at least two different PrSKU values ---
            const uniquePrSkus = new Set(imageRows.map(r => r.PrSKU).filter(s => s));
            if (uniquePrSkus.size < 2) {
                showToast("Error: CSV must contain at least two different PrSKU values.");
                matcherCsvName.textContent = "Invalid CSV";
                matcherCsvMeta.textContent = `Only ${uniquePrSkus.size} unique PrSKU(s) found`;
                matcherCsvInfo.style.background = "#fef2f2";
                matcherCsvInfo.style.borderColor = "#fecaca";
                return;
            } else {
                matcherCsvInfo.style.background = "#f0fdf4";
                matcherCsvInfo.style.borderColor = "#bbf7d0";
            }

            matcherCsvName.textContent = file.name;
            matcherCsvMeta.textContent = `${imageRows.length.toLocaleString()} records loaded`;
            matcherCsvLabel.style.display = 'none';
            matcherCsvInfo.style.display = 'flex';
            matcherCsvInfo.classList.add('active');
            lookupOptionsFromCsv();
        });
    });

    changeMatcherCsvBtn.addEventListener('click', () => {
        matcherCsvInfo.style.display = 'none';
        matcherCsvInfo.classList.remove('active');
        matcherCsvLabel.style.display = 'flex';
        matcherCsvInput.value = '';
        imageRows = [];
        updateMatcherResults();
    });

    matcherModeLookupBtn.addEventListener('click', () => setMatcherInputMode('lookup'));
    matcherModePasteBtn.addEventListener('click', () => setMatcherInputMode('paste'));

    matcherSkuASearch.addEventListener('input', (e) => {
        matcherSearchQueryA = e.target.value.toLowerCase().trim();
        renderOptions('A', fullRawOptionsA);
    });

    matcherSkuBSearch.addEventListener('input', (e) => {
        matcherSearchQueryB = e.target.value.toLowerCase().trim();
        renderOptions('B', fullRawOptionsB);
    });

    function setMatcherInputMode(mode) {
        matcherInputMode = mode;
        if (mode === 'lookup') {
            matcherModeLookupBtn.classList.add('active');
            matcherModePasteBtn.classList.remove('active');
            matcherLookupArea.style.display = 'block';
            matcherPasteArea.style.display = 'none';
            lookupOptionsFromCsv();
        } else {
            matcherModePasteBtn.classList.add('active');
            matcherModeLookupBtn.classList.remove('active');
            matcherLookupArea.style.display = 'none';
            matcherPasteArea.style.display = 'block';
            processManualOptions();
        }
        lucide.createIcons();
    }

    skuAInput.addEventListener('input', processManualOptions);
    skuBInput.addEventListener('input', processManualOptions);

    function processManualOptions() {
        if (matcherInputMode !== 'paste') return;

        fullRawOptionsA = parseManualOptions(skuAInput.value);
        fullRawOptionsB = parseManualOptions(skuBInput.value);

        allPartsA = new Set(fullRawOptionsA.map(o => o.part.toLowerCase().trim()));
        allPartsB = new Set(fullRawOptionsB.map(o => o.part.toLowerCase().trim()));

        // Auto-select if only one option exists
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
                if (partNumber) result.push({ name: optionName, part: partNumber, sku: 'SKU' });
            } else {
                result.push({ name: '', part: line, sku: 'SKU' });
            }
        });
        return result;
    }

    skuPairInput.addEventListener('input', () => {
        if (matcherInputMode === 'lookup') lookupOptionsFromCsv();
    });

    skuPairInput.addEventListener('blur', () => {
        if (matcherInputMode !== 'lookup') return;
        const text = skuPairInput.value.trim();
        const skus = text.split(/[\s\t\n,]+/).filter(s => s.length > 0);
        if (skus.length > 0) {
            skuPairInput.value = skus.join(', ');
        }
    });

    skuPairInput.addEventListener('paste', () => {
        if (matcherInputMode !== 'lookup') return;
        setTimeout(() => {
            const text = skuPairInput.value.trim();
            const skus = text.split(/[\s\t\n,]+/).filter(s => s.length > 0);
            if (skus.length > 0) {
                skuPairInput.value = skus.join(', ');
                lookupOptionsFromCsv();
            }
        }, 0);
    });

    const getOptionsForSku = (sku) => {
        if (!sku) return [];
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

    function lookupOptionsFromCsv() {
        if (imageRows.length === 0) return;

        const text = skuPairInput.value.trim();
        if (!text) {
            allPartsA = new Set();
            allPartsB = new Set();
            selectedPartA = null;
            selectedPartB = null;
            renderOptions('A', []);
            renderOptions('B', []);
            updateMatcherResults();
            return;
        }

        const skus = text.split(/[\s\t\n,]+/).filter(s => s.length > 0);
        if (skus.length < 1) return;

        const skuA = skus[0];
        const skusB = skus.slice(1);

        fullRawOptionsA = getOptionsForSku(skuA);
        fullRawOptionsB = [];
        const seenPartsB = new Set();

        skusB.forEach(s => {
            const opts = getOptionsForSku(s);
            opts.forEach(o => {
                const lowPart = o.part.toLowerCase().trim();
                if (!seenPartsB.has(lowPart)) {
                    seenPartsB.add(lowPart);
                    fullRawOptionsB.push(o);
                }
            });
        });

        allPartsA = new Set(fullRawOptionsA.map(o => o.part.toLowerCase().trim()));
        allPartsB = new Set(fullRawOptionsB.map(o => o.part.toLowerCase().trim()));

        // Check if selected parts are still valid
        const allPartsInFile = new Set(imageRows.map(r => r.ActualManufacturerPartNumber.toLowerCase().trim()));

        if (fullRawOptionsA.length === 1) selectedPartA = fullRawOptionsA[0].part;
        else if (selectedPartA && !allPartsInFile.has(selectedPartA.toLowerCase().trim())) selectedPartA = null;

        if (fullRawOptionsB.length === 1) selectedPartB = fullRawOptionsB[0].part;
        else if (selectedPartB && !allPartsInFile.has(selectedPartB.toLowerCase().trim())) selectedPartB = null;

        renderOptions('A', fullRawOptionsA);
        renderOptions('B', fullRawOptionsB);
        updateMatcherResults();
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

            // Strictly only show SKUs that were in the input area (only for Lookup Mode)
            if (matcherInputMode === 'lookup' && !inputSkus.has(rowSku)) return false;

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
            const skus = line.split(/[\t,]| {2,}/).map(s => s.trim()).filter(s => s.length > 0);
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
                    </tr>
                </thead>
                <tbody>
        `;

        cooExportData.forEach(row => {
            const statusTag = row.resolved ? `<span class="sku-tag danger">${row.resolved}</span>` : "";
            tableHtml += `
                <tr>
                    <td class="option-part">${row.sku1}</td>
                    <td class="option-part">${row.sku2}</td>
                    <td>${statusTag}</td>
                    <td><span class="text-dim" style="font-size: 0.85rem;">${row.note}</span></td>
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

        const headers = ["SKU1", "SKU2", "SKU1,SKU2", "Target", "Resolved?", "Resolution", "Image + Tags Updated?", "Note - Details (TH AIPN - Conso)", "Reason Not Resolved", "Note - Details"];
        let csvContent = "\uFEFF"; // Byte Order Mark for Excel UTF-8 support
        csvContent += headers.join(",") + "\n";

        cooExportData.forEach(row => {
            const line = [
                `"${row.sku1}"`,
                `"${row.sku2}"`,
                `"${row.skuCombined}"`,
                `"${row.target}"`,
                `"${row.resolved}"`,
                `"${row.resolution}"`,
                `"${row.imageUpdated}"`,
                `"${row.aipnNote}"`,
                `"${row.reason}"`,
                `"${row.note}"`
            ];
            csvContent += line.join(",") + "\n";
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
    const manualTagOptionA = document.getElementById('manualTagOptionA');
    const manualTagOptionB = document.getElementById('manualTagOptionB');
    const manualTagOptionALabel = document.getElementById('manualTagOptionALabel');
    const manualTagOptionBLabel = document.getElementById('manualTagOptionBLabel');
    const tagSkuAHeader = document.getElementById('tagSkuAHeader');
    const tagSkuBHeader = document.getElementById('tagSkuBHeader');
    const tagSkuASearch = document.getElementById('tagSkuASearch');
    const tagSkuBSearch = document.getElementById('tagSkuBSearch');
    const tagFilterOverlapBtn = document.getElementById('tagFilterOverlapBtn');
    const tagToggleInputsBtn = document.getElementById('tagToggleInputsBtn');
    const tagMatcherSidebar = document.getElementById('tagMatcherSidebar');
    const tagToggleIcon = document.getElementById('tagToggleIcon');
    const tagToggleText = document.getElementById('tagToggleText');

    let masterTagData = new Map(); // SKU -> Map<MPN, { tagName: value }>
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
    let isTagInputsCollapsed = false;

    // --- Tag Matcher Logic ---
    tagToggleInputsBtn.addEventListener('click', () => {
        isTagInputsCollapsed = !isTagInputsCollapsed;
        tagMatcherSidebar.classList.toggle('collapsed', isTagInputsCollapsed);

        if (isTagInputsCollapsed) {
            tagToggleIcon.setAttribute('data-lucide', 'chevron-down');
            tagToggleText.textContent = 'Restore Setup';
            showToast("Setup collapsed for analysis view.");
        } else {
            tagToggleIcon.setAttribute('data-lucide', 'chevron-up');
            tagToggleText.textContent = 'Collapse Setup';
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
                                        headerIdx = i; headers = parts; skuIdx = tempSkuIdx; mpnIdx = tempMpnIdx; foundHeader = true; break;
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
        tagFileStatusArea.style.display = 'none';
        tagCsvLabel.style.display = 'flex';
        tagCsvInput.value = '';
        masterTagData.clear();
        tagFilesLoadedCount = 0;
        tagLoadedFilesList = [];
    });

    clearTagBtn.addEventListener('click', () => {
        tagSkuPairInput.value = '';
        tagResultsContainer.innerHTML = '<div class="empty-state"><p>Upload tag data and select options to see comparison.</p></div>';
        tagSkuASelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
        tagSkuBSelectArea.innerHTML = '<p class="empty-text">Enter SKUs to see options</p>';
        manualTagOptionA.value = '';
        manualTagOptionB.value = '';
        manualTagOptionALabel.textContent = 'SKU A Options';
        manualTagOptionBLabel.textContent = 'SKU B Options';
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
        masterTagData.clear();
        tagFilesLoadedCount = 0;
        tagLoadedFilesList = [];
    });

    tagSkuPairInput.addEventListener('input', lookupTagOptions);
    manualTagOptionA.addEventListener('input', lookupTagOptions);
    manualTagOptionB.addEventListener('input', lookupTagOptions);

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
            manualTagOptionALabel.textContent = 'SKU A Options';
            manualTagOptionBLabel.textContent = 'SKU B Options';
            tagSkuAHeader.textContent = 'Select SKU A Option';
            tagSkuBHeader.textContent = 'Select SKU B Option';
            return;
        }

        const skus = text.split(/[\s\t\n,]+/).filter(s => s.length > 0);
        if (skus.length < 1) return;

        const skuA = skus[0];
        const skuB = skus[1] || null;

        currentTagSkus = [skuA, skuB].filter(Boolean);

        // Update Labels
        manualTagOptionALabel.textContent = `${skuA} Options`;
        tagSkuAHeader.textContent = `Select SKU ${skuA} Option`;
        if (skuB) {
            manualTagOptionBLabel.textContent = `${skuB} Options`;
            tagSkuBHeader.textContent = `Select SKU ${skuB} Option`;
        } else {
            manualTagOptionBLabel.textContent = 'SKU B Options';
            tagSkuBHeader.textContent = 'Select SKU B Option';
        }

        const getMergedOptions = (sku, manualText) => {
            const fromFile = masterTagData.get(sku) || new Map();
            const results = new Map(); // Part -> Name

            // 1. Add from file
            fromFile.forEach((tags, mpn) => {
                // Try to find a human name in tags
                const name = tags['Option Name'] || tags['Product Name'] || tags['Option'] || "Option (File)";
                results.set(mpn, name);
            });

            // 2. Add from manual (Line-based)
            const manualLines = manualText.split(/\n/).filter(line => line.trim());
            manualLines.forEach(line => {
                // Heuristic: Last part after comma/tab/pipe is usually the part number
                const parts = line.split(/[,|\t]+/).map(p => p.trim()).filter(p => p.length > 0);
                if (parts.length >= 2) {
                    const mpn = parts.pop();
                    const name = parts.join(', ');
                    results.set(mpn, name);
                } else if (parts.length === 1) {
                    // Just a part number
                    if (!results.has(parts[0])) {
                        results.set(parts[0], "Manual Entry");
                    }
                }
            });

            return Array.from(results.entries()).map(([part, name]) => ({ part, name }));
        };

        fullOptionsA = getMergedOptions(skuA, manualTagOptionA.value);
        fullOptionsB = skuB ? getMergedOptions(skuB, manualTagOptionB.value) : [];

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
                <div class="option-info">
                    <div style="display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap;">
                        <span class="option-name">${name}</span>
                        ${isShared ? '<span class="shared-badge">Same Part Number</span>' : ''}
                        <span style="font-size: 0.7rem; color: var(--text-dim); opacity: 0.7; margin-left: auto;">${currentSku}</span>
                    </div>
                    <span class="option-part">${mpn}</span>
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
        clean = clean.replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/_/g, ' ')
            .toLowerCase();

        // Title Case
        return clean.split(' ')
            .map(word => word.charAt(0).toUpperCase() + word.slice(1))
            .join(' ');
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

        const allTags = Array.from(new Set([...Object.keys(data1), ...Object.keys(data2)]))
            .filter(tag => {
                if (!tag) return false;
                const lower = tag.toLowerCase().trim();
                return !excludedTags.some(ex => lower.includes(ex));
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
    }

    function renderTagResults(label1, label2) {
        tagResultsContainer.innerHTML = '';

        if (tagExportData.length === 0) {
            tagResultsContainer.innerHTML = '<div class="empty-state"><p>No tag data found for these options.</p></div>';
            return;
        }

        // Apply filtering
        const filteredData = tagExportData.filter(row => {
            if (currentTagFilter === 'false') return !row.match;
            if (currentTagFilter === 'true') return row.match;
            return true;
        });

        if (filteredData.length === 0) {
            tagResultsContainer.innerHTML = `<div class="empty-state"><p>No rows match the "${currentTagFilter}" filter.</p></div>`;
            return;
        }

        let tableHtml = `
            <table class="data-table" style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr>
                        <th style="width: 35%; text-align: left; padding: 1rem; border-bottom: 2px solid var(--card-border);">Tag Name</th>
                        <th style="width: 25%; text-align: left; padding: 1rem; border-bottom: 2px solid var(--card-border);">${label1}</th>
                        <th style="width: 25%; text-align: left; padding: 1rem; border-bottom: 2px solid var(--card-border);">${label2}</th>
                        <th style="width: 15%; text-align: center; padding: 1rem; border-bottom: 2px solid var(--card-border);">Status</th>
                    </tr>
                </thead>
                <tbody>
        `;

        filteredData.forEach(row => {
            const statusClass = row.match ? 'success' : 'danger';
            const statusText = row.match ? 'TRUE' : 'FALSE';
            const rowClass = row.match ? '' : 'style="background: rgba(239, 68, 68, 0.03);"';

            tableHtml += `
                <tr ${rowClass}>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); font-weight: 500; font-size: 0.9rem;">
                        ${humanizeTagName(row.tag)}
                    </td>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); font-size: 0.9rem; color: var(--text-main);">
                        ${row.val1 || '<span style="color: #cbd5e1;">N/A</span>'}
                    </td>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); font-size: 0.9rem; color: var(--text-main);">
                        ${row.val2 || '<span style="color: #cbd5e1;">N/A</span>'}
                    </td>
                    <td style="padding: 0.75rem 1rem; border-bottom: 1px solid var(--card-border); text-align: center;">
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

        processFile(file, (content) => {
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

            exdisplayFileName.textContent = file.name;
            exdisplayFileMeta.textContent = `${masterExclusionData.size.toLocaleString()} records loaded`;
            exclusionCsvLabel.style.display = 'none';
            exclusionCsvInfo.classList.add('active');
            lucide.createIcons();
        });
    });

    changeExclusionCsvBtn.addEventListener('click', () => {
        exclusionCsvInfo.classList.remove('active');
        exclusionCsvLabel.style.display = 'flex';
        exclusionCsvInput.value = '';
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
        let csvContent = headers.join(",") + "\n";

        exclusionExportData.forEach(row => {
            if (row.isClean) {
                // If clean, only SKU1 and SKU2 are filled, others blank
                const line = [
                    `"${row.sku1}"`,
                    `"${row.sku2}"`,
                    `""`, `""`, `""`, `""`, `""`, `""`, `""`
                ];
                csvContent += line.join(",") + "\n";
            } else {
                const line = [
                    `"${row.sku1}"`,
                    `"${row.sku2}"`,
                    `""`, // Target
                    `"No"`, // Resolved - Tool only does Resolved = No
                    `""`, // Resolution
                    `"No"`, // Image + Tags Updated?
                    `""`, // Note - Details (TH AIPN - Conso)
                    `"${row.reason}"`,
                    `"${row.note}"`
                ];
                csvContent += line.join(",") + "\n";
            }
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
            const parts = line.split(/[\t\s,]+/).map(s => s.trim()).filter(s => s);
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
});

