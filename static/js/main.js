document.addEventListener('DOMContentLoaded', function() {
    const generateButton = document.getElementById('generateButton');
    const macroDescription = document.getElementById('macroDescription');
    const macroList = document.getElementById('macroList');
    const categoryList = document.getElementById('categoryList');
    const previewArea = document.getElementById('previewArea');
    const previewCode = document.getElementById('previewCode');
    const templateSelect = document.getElementById('templateSelect');
    const templateId = document.getElementById('templateId');
    const categorySelect = document.getElementById('categorySelect');
    const macroTypeRadios = document.getElementsByName('macroType');

    let currentCategory = 'all';
    let showPublicOnly = false;

    // カテゴリーの定義
    const categories = {
        'TEMPLATE': 'テンプレート',
        'AI_GENERATED': 'AI生成',
        'DATA_PROCESSING': 'データ処理',
        'FORMATTING': 'フォーマット',
        'CALCULATION': '計算',
        'AUTOMATION': '自動化',
        'REPORTING': 'レポート',
        'CUSTOM': 'カスタム'
    };

    // バージョン履歴を表示
    async function showVersionHistory(macroId) {
        try {
            const response = await fetch(`/macros/${macroId}/versions`);
            if (response.ok) {
                const versions = await response.json();
                const versionList = versions.map(version => `
                    <div class="version-item mb-2">
                        <div class="d-flex justify-content-between align-items-center">
                            <span>バージョン ${version.version_number}</span>
                            <small class="text-muted">${new Date(version.created_at).toLocaleDateString()}</small>
                        </div>
                        <pre class="mt-2"><code>${version.content}</code></pre>
                    </div>
                `).join('');

                previewArea.style.display = 'block';
                previewCode.innerHTML = `
                    <h4>バージョン履歴</h4>
                    ${versionList}
                `;
            }
        } catch (error) {
            console.error('バージョン履歴の取得に失敗:', error);
        }
    }

    // マクロの共有設定を更新
    async function toggleShare(macroId, isPublic) {
        try {
            const response = await fetch(`/macros/${macroId}/share`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ is_public: isPublic })
            });

            if (response.ok) {
                loadMacros();
            } else {
                throw new Error('共有設定の更新に失敗しました');
            }
        } catch (error) {
            alert(error.message);
        }
    }

    // カテゴリー選択肢の初期化
    function initializeCategories() {
        Object.entries(categories).forEach(([value, label]) => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = label;
            categorySelect.appendChild(option);

            const li = document.createElement('li');
            li.className = 'nav-item';
            li.innerHTML = `
                <a class="nav-link" href="#" data-category="${value}">
                    ${label}
                </a>
            `;
            categoryList.appendChild(li);
        });

        // カテゴリーフィルター機能
        categoryList.addEventListener('click', (e) => {
            if (e.target.classList.contains('nav-link')) {
                e.preventDefault();
                const category = e.target.dataset.category;
                currentCategory = category;

                categoryList.querySelectorAll('.nav-link').forEach(link => {
                    link.classList.remove('active');
                });
                e.target.classList.add('active');

                loadMacros();
            }
        });
    }

    // マクロタイプの切り替え
    macroTypeRadios.forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.value === 'template') {
                templateSelect.style.display = 'block';
                macroDescription.disabled = true;
                categorySelect.value = 'TEMPLATE';
                categorySelect.disabled = true;
            } else {
                templateSelect.style.display = 'none';
                macroDescription.disabled = false;
                categorySelect.disabled = false;
            }
        });
    });

    // マクロ生成処理
    generateButton.addEventListener('click', async () => {
        const useAI = document.getElementById('aiGenerate').checked;

        if (useAI && !macroDescription.value) {
            alert('マクロの要件を入力してください');
            return;
        }

        if (!useAI && !templateId.value) {
            alert('テンプレートを選択してください');
            return;
        }

        if (!categorySelect.value) {
            alert('カテゴリーを選択してください');
            return;
        }

        generateButton.disabled = true;
        try {
            const response = await fetch('/generate-macro', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    description: macroDescription.value,
                    use_ai: useAI,
                    template_id: useAI ? null : parseInt(templateId.value),
                    category: categorySelect.value
                })
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'generated_macro.xlsx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);

                loadMacros();
            } else {
                throw new Error('マクロの生成に失敗しました');
            }
        } catch (error) {
            alert(error.message);
        } finally {
            generateButton.disabled = false;
        }
    });

    // マクロのダウンロード
    async function downloadMacro(macroId) {
        try {
            const response = await fetch('/generate-macro', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    template_id: macroId,
                    use_ai: false
                })
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'macro.xlsx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            } else {
                throw new Error('マクロのダウンロードに失敗しました');
            }
        } catch (error) {
            alert(error.message);
        }
    }

    // マクロライブラリの読み込み
    async function loadMacros() {
        try {
            const response = await fetch(`/macros?public=${showPublicOnly}`);
            if (response.ok) {
                const macros = await response.json();

                // テンプレートの選択肢を更新
                const templates = macros.filter(macro => macro.category === 'TEMPLATE');
                templateId.innerHTML = '<option value="">テンプレートを選択してください</option>';
                templates.forEach(template => {
                    const option = document.createElement('option');
                    option.value = template.id;
                    option.textContent = template.title;
                    templateId.appendChild(option);
                });

                // マクロリストを更新
                macroList.innerHTML = '';
                let filteredMacros = macros;
                if (currentCategory !== 'all') {
                    filteredMacros = macros.filter(macro => macro.category === currentCategory);
                }

                filteredMacros.forEach(macro => {
                    const li = document.createElement('li');
                    li.className = 'nav-item macro-item';
                    const date = new Date(macro.created_at).toLocaleDateString();
                    li.innerHTML = `
                        <div class="macro-content">
                            <div class="d-flex justify-content-between align-items-center">
                                <small class="text-muted">${date}</small>
                                <span class="badge bg-${macro.category === 'TEMPLATE' ? 'primary' : 'success'}">${categories[macro.category] || macro.category}</span>
                            </div>
                            <h6 class="macro-title mb-2">${macro.title || '無題のマクロ'}</h6>
                            <p class="macro-description mb-2">${macro.description.split('\n')[0]}</p>
                            <div class="btn-group w-100">
                                <button class="btn btn-sm btn-outline-primary" onclick="event.stopPropagation(); downloadMacro(${macro.id})">
                                    ダウンロード
                                </button>
                                <button class="btn btn-sm btn-outline-info" onclick="event.stopPropagation(); showVersionHistory(${macro.id})">
                                    履歴
                                </button>
                                <button class="btn btn-sm ${macro.is_public ? 'btn-outline-warning' : 'btn-outline-success'}" onclick="event.stopPropagation(); toggleShare(${macro.id}, ${!macro.is_public})">
                                    ${macro.is_public ? '非公開にする' : '共有する'}
                                </button>
                            </div>
                        </div>
                    `;

                    li.addEventListener('click', () => {
                        macroDescription.value = macro.description;
                        if (macro.category === 'TEMPLATE') {
                            document.getElementById('useTemplate').checked = true;
                            templateSelect.style.display = 'block';
                            templateId.value = macro.id;
                            macroDescription.disabled = true;
                        } else {
                            document.getElementById('aiGenerate').checked = true;
                            templateSelect.style.display = 'none';
                            macroDescription.disabled = false;
                        }
                        categorySelect.value = macro.category;
                    });

                    macroList.appendChild(li);
                });
            }
        } catch (error) {
            console.error('マクロの読み込みに失敗しました:', error);
        }
    }

    // 初期化
    initializeCategories();
    loadMacros();

    // グローバル関数の定義
    window.downloadMacro = downloadMacro;
    window.showVersionHistory = showVersionHistory;
    window.toggleShare = toggleShare;
});