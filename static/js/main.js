document.addEventListener('DOMContentLoaded', function() {
    const generateButton = document.getElementById('generateButton');
    const macroDescription = document.getElementById('macroDescription');
    const macroList = document.getElementById('macroList');
    const previewArea = document.getElementById('previewArea');
    const previewCode = document.getElementById('previewCode');

    // マクロ生成処理
    generateButton.addEventListener('click', async () => {
        if (!macroDescription.value) {
            alert('マクロの要件を入力してください');
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
                    description: macroDescription.value
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
                macroList.innerHTML = '';
                
                macros.forEach(macro => {
                    const li = document.createElement('li');
                    li.className = 'nav-item macro-item';
                    const date = new Date(macro.created_at).toLocaleDateString();
                    li.innerHTML = `
                        <div>
                            <small class="text-muted">${date}</small>
                            <div>${macro.description}</div>
                        </div>
                    `;
                    
                    li.addEventListener('click', async () => {
                        macroDescription.value = macro.description;
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
