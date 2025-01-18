document.addEventListener('DOMContentLoaded', function() {
    const generateButton = document.getElementById('generateButton');
    const macroDescription = document.getElementById('macroDescription');
    const macroList = document.getElementById('macroList');
    const previewArea = document.getElementById('previewArea');
    const previewCode = document.getElementById('previewCode');
    const templateSelect = document.getElementById('templateSelect');
    const templateId = document.getElementById('templateId');
    const macroTypeRadios = document.getElementsByName('macroType');

    // マクロタイプの切り替え
    macroTypeRadios.forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.value === 'template') {
                templateSelect.style.display = 'block';
                macroDescription.disabled = true;
            } else {
                templateSelect.style.display = 'none';
                macroDescription.disabled = false;
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
                    template_id: useAI ? null : parseInt(templateId.value)
                })
            });

            if (response.ok) {
                // Excelファイルのダウンロード
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'generated_macro.xlsx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);

                // ライブラリを更新
                loadMacros();

                // プレビュー表示
                previewArea.style.display = 'block';
                previewCode.textContent = macroDescription.value;
            } else {
                throw new Error('マクロの生成に失敗しました');
            }
        } catch (error) {
            alert(error.message);
        } finally {
            generateButton.disabled = false;
        }
    });

    // マクロライブラリの読み込み
    async function loadMacros() {
        try {
            const response = await fetch('/macros');
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
                macros.forEach(macro => {
                    const li = document.createElement('li');
                    li.className = 'nav-item macro-item';
                    const date = new Date(macro.created_at).toLocaleDateString();
                    li.innerHTML = `
                        <div>
                            <small class="text-muted">${date}</small>
                            <div class="macro-title">${macro.title || '無題のマクロ'}</div>
                            <small class="text-muted">${macro.category === 'TEMPLATE' ? 'テンプレート' : 'AI生成'}</small>
                            <div class="macro-description">${macro.description}</div>
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
                    });

                    macroList.appendChild(li);
                });
            }
        } catch (error) {
            console.error('マクロの読み込みに失敗しました:', error);
        }
    }

    // 初期読み込み
    loadMacros();
});