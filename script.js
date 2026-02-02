// グローバル変数：請求書データを保存
let currentInvoiceData = null;
let sealImageBase64 = null;

// 画像をBase64に変換（Promise版）
function loadSealImage() {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = function() {
            try {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                sealImageBase64 = canvas.toDataURL('image/png');
                resolve(sealImageBase64);
            } catch (error) {
                console.warn('画像の変換に失敗しました:', error);
                reject(error);
            }
        };
        img.onerror = function() {
            console.warn('印鑑画像の読み込みに失敗しました。画像ファイル（seal.png）が同じフォルダにあることを確認してください。');
            reject(new Error('画像の読み込みに失敗しました'));
        };
        // CORSを回避するため、crossOriginを設定しない
        img.src = 'seal.png';
    });
}

// 画像を確実に読み込む（再試行機能付き）
async function ensureSealImageLoaded() {
    if (sealImageBase64) {
        return sealImageBase64;
    }
    
    try {
        await loadSealImage();
        return sealImageBase64;
    } catch (error) {
        // 再試行
        try {
            await new Promise(resolve => setTimeout(resolve, 500));
            await loadSealImage();
            return sealImageBase64;
        } catch (retryError) {
            console.warn('画像の読み込みに失敗しました。画像なしで続行します。');
            return null;
        }
    }
}

// ページ読み込み時に画像を読み込む
window.addEventListener('DOMContentLoaded', function() {
    loadSealImage().catch(() => {
        // エラーは無視（後で再試行可能）
    });
});

// 請求書番号を自動生成（日付ベースの連番）
function generateInvoiceNumber(issueDate) {
    // 日付をYYYYMMDD形式に変換
    const dateStr = issueDate.replace(/-/g, '');
    
    // ローカルストレージから今日の連番を取得
    const storageKey = `invoice_counter_${dateStr}`;
    let counter = parseInt(localStorage.getItem(storageKey) || '0', 10);
    
    // 連番をインクリメント
    counter++;
    
    // ローカルストレージに保存
    localStorage.setItem(storageKey, counter.toString());
    
    // 請求書番号を生成（INV-YYYYMMDD-XXX形式）
    const invoiceNumber = `INV-${dateStr}-${String(counter).padStart(3, '0')}`;
    
    return invoiceNumber;
}

// 項目を追加
function addItem() {
    const container = document.getElementById('itemsContainer');
    const itemRow = document.createElement('div');
    itemRow.className = 'item-row';
    
    const today = new Date().toISOString().split('T')[0];
    itemRow.innerHTML = `
        <input type="date" class="item-date" value="${today}" required>
        <input type="text" class="item-description" placeholder="項目名" required>
        <input type="number" class="item-amount" placeholder="金額" min="0" step="1" required>
        <select class="item-tax-rate">
            <option value="0">0%</option>
            <option value="8">8%</option>
            <option value="10" selected>10%</option>
        </select>
        <select class="item-amount-type">
            <option value="exclusive">税抜</option>
            <option value="inclusive" selected>税込</option>
        </select>
        <button type="button" class="remove-item-btn" onclick="removeItem(this)">削除</button>
    `;
    
    container.appendChild(itemRow);
    
    // 最初の項目以外は削除ボタンを表示
    updateRemoveButtons();
}

// 項目を削除
function removeItem(btn) {
    const container = document.getElementById('itemsContainer');
    if (container.children.length > 1) {
        btn.parentElement.remove();
        updateRemoveButtons();
    }
}

// 削除ボタンの表示/非表示を更新
function updateRemoveButtons() {
    const container = document.getElementById('itemsContainer');
    const removeButtons = container.querySelectorAll('.remove-item-btn');
    
    if (container.children.length > 1) {
        removeButtons.forEach(btn => btn.style.display = 'block');
    } else {
        removeButtons.forEach(btn => btn.style.display = 'none');
    }
}

// 請求書を生成
async function generateInvoice() {
    try {
        await generateInvoiceInternal();
    } catch (error) {
        console.error('請求書の生成中にエラーが発生しました:', error);
        alert('請求書の生成中にエラーが発生しました: ' + error.message);
    }
}

// 請求書を生成（内部関数）
async function generateInvoiceInternal() {
    // 基本情報を取得
    const issueDate = document.getElementById('issueDate').value;
    const clientCompanyType = document.getElementById('clientCompanyType').value;
    const clientName = document.getElementById('clientName').value;
    const companyName = '株式会社KASEKI CREATIVE'; // 固定
    
    // 請求書番号を自動生成
    const invoiceNumber = generateInvoiceNumber(issueDate);
    
    // 請求書番号を表示
    const invoiceNumberDisplay = document.querySelector('.invoice-number-display');
    if (invoiceNumberDisplay) {
        invoiceNumberDisplay.textContent = invoiceNumber;
    }
    
    // お客様名を組み立て
    let fullClientName = '';
    if (clientCompanyType && clientCompanyType !== 'その他') {
        fullClientName = clientCompanyType + clientName;
    } else if (clientCompanyType === 'その他') {
        // その他の場合は入力欄に法人種別を入力してもらう想定
        fullClientName = clientName;
    } else {
        fullClientName = clientName;
    }
    
    // バリデーション
    if (!issueDate || !clientName) {
        alert('基本情報をすべて入力してください。');
        return;
    }
    
    // 項目を取得
    const itemRows = document.querySelectorAll('.item-row');
    const items = [];
    
    itemRows.forEach(row => {
        const date = row.querySelector('.item-date').value;
        const description = row.querySelector('.item-description').value;
        const amount = parseFloat(row.querySelector('.item-amount').value);
        const taxRate = parseFloat(row.querySelector('.item-tax-rate').value);
        const amountType = row.querySelector('.item-amount-type').value;
        
        if (date && description && !isNaN(amount) && amount > 0) {
            items.push({ date, description, amount, taxRate, amountType });
        }
    });
    
    if (items.length === 0) {
        alert('少なくとも1つの項目を入力してください。');
        return;
    }
    
    // 計算処理
    let subtotal = 0;
    let totalTax = 0;
    let total = 0;
    
    items.forEach(item => {
        let itemAmount = parseFloat(item.amount);
        let itemTaxRate = parseFloat(item.taxRate) || 0;
        
        if (item.amountType === 'inclusive') {
            // 税込み金額の場合、税抜きに変換
            if (itemTaxRate > 0) {
                // 浮動小数点の精度問題を回避するため、整数演算を使用
                // 税抜金額 = 税込金額 ÷ (1 + 税率/100) を切り捨て
                // 例: 440000 / 1.1 = 400000
                // 整数演算: (440000 * 100) / 110 = 400000
                const taxRatePercent = itemTaxRate; // 10, 8, 0など
                const denominator = 100 + taxRatePercent; // 110, 108, 100など
                // 整数演算で計算してから切り捨て
                item.exclusiveAmount = Math.floor((itemAmount * 100) / denominator);
                // 税金 = 税込金額 - 税抜金額
                item.taxAmount = itemAmount - item.exclusiveAmount;
            } else {
                item.exclusiveAmount = Math.floor(itemAmount);
                item.taxAmount = 0;
            }
        } else {
            // 税抜き金額の場合
            item.exclusiveAmount = Math.floor(itemAmount);
            // 税金 = 税抜金額 × 税率/100 を切り捨て
            // 整数演算で計算してから切り捨て
            item.taxAmount = itemTaxRate > 0 ? Math.floor((item.exclusiveAmount * itemTaxRate) / 100) : 0;
        }
        
        // 各項目の税抜金額と税金を合計に加算
        subtotal += item.exclusiveAmount;
        totalTax += item.taxAmount;
    });
    
    // 合計金額（税込）= 合計税抜金額 + 合計税金
    total = subtotal + totalTax;
    
    // 請求書データを保存
    currentInvoiceData = {
        issueDate,
        invoiceNumber,
        clientName: fullClientName,
        companyName,
        items,
        subtotal,
        totalTax,
        total,
        createdAt: new Date().toISOString() // 作成日時を追加
    };
    
    // 履歴に保存
    saveInvoiceToHistory(currentInvoiceData);
    
    // 画像を確実に読み込む（エラーが発生しても続行）
    let sealImage = null;
    try {
        sealImage = await ensureSealImageLoaded();
    } catch (error) {
        console.warn('画像の読み込みに失敗しましたが、請求書の生成を続行します:', error);
        sealImage = null;
    }
    
    // 日付フォーマット
    const formattedDate = formatDate(issueDate);
    
    // 請求書HTMLを生成
    let invoiceHTML = `
        <div class="invoice-header">
            <h2>請求書</h2>
        </div>
        
        <div class="invoice-info">
            <div class="invoice-info-left">
                <p><strong>${escapeHtml(fullClientName)}様</strong></p>
                <div class="invoice-amount-summary">
                    <p><strong>請求金額:</strong> ¥${formatNumber(total)}</p>
                </div>
            </div>
            <div class="invoice-info-right">
                <div class="company-seal-section">
                    <p><strong>発行元:</strong> ${companyName}</p>
                    ${sealImage ? `<img src="${sealImage}" alt="会社印鑑" class="company-seal">` : ''}
                </div>
                <p><strong>登録番号:</strong> T1120001277681</p>
                <p><strong>発行日:</strong> ${formattedDate}</p>
            </div>
        </div>
        
        <table class="invoice-table">
            <thead>
                <tr>
                    <th>日付</th>
                    <th>項目</th>
                    <th class="amount">料金</th>
                    <th class="amount">税金</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    items.forEach(item => {
        const itemDate = formatDate(item.date);
        invoiceHTML += `
            <tr>
                <td>${escapeHtml(itemDate)}</td>
                <td>${escapeHtml(item.description)}</td>
                <td class="amount">¥${formatNumber(item.exclusiveAmount)}</td>
                <td class="amount">¥${formatNumber(item.taxAmount)}</td>
            </tr>
        `;
    });
    
    invoiceHTML += `
            </tbody>
        </table>
        
        <div class="invoice-total">
            <div class="invoice-total-row">
                <div class="invoice-total-label">合計金額（税抜）:</div>
                <div class="invoice-total-value">¥${formatNumber(subtotal)}</div>
            </div>
            <div class="invoice-total-row">
                <div class="invoice-total-label">税金の合計:</div>
                <div class="invoice-total-value">¥${formatNumber(totalTax)}</div>
            </div>
            <div class="invoice-total-row invoice-total-final">
                <div class="invoice-total-label">合計金額（税込）:</div>
                <div class="invoice-total-value">¥${formatNumber(total)}</div>
            </div>
        </div>
        
        <div class="bank-info">
            <h3>振込先情報</h3>
            <div class="bank-details">
                <div class="bank-detail-item">
                    <p><strong>銀行名:</strong> 京都信用金庫(普通)</p>
                    <p class="furigana">キョウトシンヨウキンコ</p>
                </div>
                <div class="bank-detail-item">
                    <p><strong>支店名:</strong> 寝屋川支店</p>
                    <p class="furigana">ネヤガワシテン</p>
                </div>
                <p><strong>口座番号:</strong> 3035077</p>
                <p><strong>口座名義:</strong> ｶ)ｶｾｷｸﾘｴｲﾃｨﾌﾞ</p>
            </div>
            <p class="bank-note">※振り込み手数料は御社ご負担にてお願い致します。</p>
        </div>
        
        <div class="company-info">
            <p>株式会社 KASEKI CREATIVE</p>
            <p>〒 572-0811</p>
            <p>大阪府寝屋川市 楠根南町 5-15</p>
            <p>☎：080-5718-7502</p>
        </div>
        
        <div class="invoice-number-bottom">
            <p><strong>請求書番号:</strong> ${invoiceNumber}</p>
        </div>
        
        <button type="button" class="payment-received-btn" onclick="generateReceipt()">
            支払いを受けました
        </button>
    `;
    
    // プレビューを表示
    const preview = document.getElementById('invoicePreview');
    preview.innerHTML = invoiceHTML;
    preview.style.display = 'block';
    
    // 印刷ボタンを表示
    document.getElementById('printBtn').style.display = 'block';
    
    // 領収書プレビューを非表示
    document.getElementById('receiptPreview').style.display = 'none';
    
    // スクロール
    preview.scrollIntoView({ behavior: 'smooth' });
}

// 領収書を生成
async function generateReceipt() {
    try {
        await generateReceiptInternal();
    } catch (error) {
        console.error('領収書の生成中にエラーが発生しました:', error);
        alert('領収書の生成中にエラーが発生しました: ' + error.message);
    }
}

// 領収書を生成（内部関数）
async function generateReceiptInternal() {
    if (!currentInvoiceData) {
        alert('先に請求書を生成してください。');
        return;
    }
    
    const data = currentInvoiceData;
    const today = new Date().toISOString().split('T')[0];
    const receiptDate = formatDate(today);
    
    // 領収書番号を生成（請求書番号から自動生成、または日付ベース）
    let receiptNumber = '';
    if (data.invoiceNumber) {
        receiptNumber = data.invoiceNumber.replace('INV', 'REC');
    } else {
        const dateStr = today.replace(/-/g, '');
        receiptNumber = `REC-${dateStr}-001`;
    }
    
    // 画像を確実に読み込む（エラーが発生しても続行）
    let sealImage = null;
    try {
        sealImage = await ensureSealImageLoaded();
    } catch (error) {
        console.warn('画像の読み込みに失敗しましたが、領収書の生成を続行します:', error);
        sealImage = null;
    }
    
    // 領収書HTMLを生成
    let receiptHTML = `
        <div class="receipt-header">
            <h2>領収書</h2>
        </div>
        
        <div class="receipt-info">
            <div class="receipt-info-left">
                <p><strong>${escapeHtml(data.clientName)}様</strong></p>
            </div>
            <div class="receipt-info-right">
                <div class="company-seal-section">
                    <p><strong>発行元:</strong> ${escapeHtml(data.companyName)}</p>
                    ${sealImage ? `<img src="${sealImage}" alt="会社印鑑" class="company-seal">` : ''}
                </div>
                <p><strong>登録番号:</strong> T1120001277681</p>
                <p><strong>発行日:</strong> ${receiptDate}</p>
            </div>
        </div>
        
        <table class="receipt-table">
            <thead>
                <tr>
                    <th>日付</th>
                    <th>項目</th>
                    <th class="amount">料金</th>
                    <th class="amount">税金</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    data.items.forEach(item => {
        const itemDate = formatDate(item.date);
        receiptHTML += `
            <tr>
                <td>${escapeHtml(itemDate)}</td>
                <td>${escapeHtml(item.description)}</td>
                <td class="amount">¥${formatNumber(item.exclusiveAmount)}</td>
                <td class="amount">¥${formatNumber(item.taxAmount)}</td>
            </tr>
        `;
    });
    
    receiptHTML += `
            </tbody>
        </table>
        
        <div class="receipt-total">
            <div class="receipt-total-row">
                <div class="receipt-total-label">合計金額（税抜）:</div>
                <div class="receipt-total-value">¥${formatNumber(data.subtotal)}</div>
            </div>
            <div class="receipt-total-row">
                <div class="receipt-total-label">税金の合計:</div>
                <div class="receipt-total-value">¥${formatNumber(data.totalTax)}</div>
            </div>
            <div class="receipt-total-row receipt-total-final">
                <div class="receipt-total-label">合計金額（税込）:</div>
                <div class="receipt-total-value">¥${formatNumber(data.total)}</div>
            </div>
        </div>
        
        <div class="receipt-note">
            <p class="note-text">上記正に領収いたしました</p>
            <p>但し書き: ${data.items.map(item => escapeHtml(item.description)).join('、')}</p>
        </div>
        
        <div class="receipt-number-bottom">
            <p><strong>領収書番号:</strong> ${receiptNumber}</p>
        </div>
        
        <div class="company-info">
            <p>株式会社 KASEKI CREATIVE</p>
            <p>〒 572-0811</p>
            <p>大阪府寝屋川市 楠根南町 5-15</p>
            <p>☎：080-5718-7502</p>
        </div>
        
        <button type="button" class="print-receipt-btn" onclick="saveReceiptAsPDF()">
            領収書をPDFとして保存
        </button>
    `;
    
    // 領収書プレビューを表示
    const receiptPreview = document.getElementById('receiptPreview');
    receiptPreview.innerHTML = receiptHTML;
    receiptPreview.style.display = 'block';
    
    // スクロール
    receiptPreview.scrollIntoView({ behavior: 'smooth' });
}

// HTMLエスケープ（XSS対策）
function escapeHtml(text) {
    if (text === null || text === undefined) {
        return '';
    }
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

// 日付フォーマット
function formatDate(dateString) {
    if (!dateString) {
        return '';
    }
    try {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) {
            return dateString; // 無効な日付の場合は元の文字列を返す
        }
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}年${month}月${day}日`;
    } catch (error) {
        console.error('日付フォーマットエラー:', error);
        return dateString; // エラー時は元の文字列を返す
    }
}

// 数値フォーマット（カンマ区切り）
function formatNumber(num) {
    if (num === null || num === undefined || isNaN(num)) {
        return '0';
    }
    return Math.round(num).toLocaleString('ja-JP');
}

// 請求書をPDFとして保存
async function saveInvoiceAsPDF() {
    const invoicePreview = document.getElementById('invoicePreview');
    if (!invoicePreview || invoicePreview.style.display === 'none') {
        alert('先に請求書を生成してください。');
        return;
    }
    
    // 請求書番号を取得してファイル名に使用
    const invoiceNumberElement = invoicePreview.querySelector('.invoice-number-bottom p');
    let fileName = '請求書';
    if (invoiceNumberElement) {
        const invoiceNumber = invoiceNumberElement.textContent.replace('請求書番号:', '').trim();
        if (invoiceNumber) {
            fileName = `請求書_${invoiceNumber}`;
        }
    }
    
    // 要素の内容が存在することを確認
    if (!invoicePreview.innerHTML || invoicePreview.innerHTML.trim() === '') {
        alert('請求書の内容が空です。先に請求書を生成してください。');
        return;
    }
    
    // 元のスタイルを保存
    const originalDisplay = invoicePreview.style.display;
    const originalPosition = invoicePreview.style.position || '';
    const originalMargin = invoicePreview.style.margin || '';
    const originalOverflow = invoicePreview.style.overflow || '';
    const originalMaxWidth = invoicePreview.style.maxWidth || '';
    const originalWidth = invoicePreview.style.width || '';
    const originalHeight = invoicePreview.style.height || '';
    const originalPadding = invoicePreview.style.padding || '';
    const originalBorder = invoicePreview.style.border || '';
    const originalBorderRadius = invoicePreview.style.borderRadius || '';
    
    // ボタンの元の表示状態を保存
    const paymentButton = invoicePreview.querySelector('.payment-received-btn');
    const originalButtonDisplay = paymentButton ? paymentButton.style.display : '';
    
    // 親要素のスタイルも保存
    const parentContainer = invoicePreview.parentElement;
    const originalParentMaxWidth = parentContainer ? parentContainer.style.maxWidth || '' : '';
    const originalParentOverflow = parentContainer ? parentContainer.style.overflow || '' : '';
    
    try {
        // ボタンを非表示にする（PDF出力時）
        if (paymentButton) {
            paymentButton.style.display = 'none';
        }
        
        // 親要素のスタイルを調整（幅制限を解除）
        if (parentContainer) {
            parentContainer.style.maxWidth = 'none';
            parentContainer.style.overflow = 'visible';
        }
        
        // A4サイズに合わせて要素の幅を設定
        // A4: 210mm、マージン: 5mm x 2 = 10mm、利用可能幅: 200mm
        // 96 DPIを想定: 1mm = 3.779527559 pixels
        const mmToPx = 3.779527559;
        const a4WidthMm = 210;
        const marginMm = 5;
        const availableWidthMm = a4WidthMm - (marginMm * 2);
        const availableWidthPx = availableWidthMm * mmToPx;
        
        // 要素を確実に表示状態にする
        invoicePreview.style.display = 'block';
        invoicePreview.style.position = 'relative';
        invoicePreview.style.margin = '0';
        invoicePreview.style.overflow = 'visible';
        invoicePreview.style.maxWidth = `${availableWidthPx}px`;
        invoicePreview.style.width = `${availableWidthPx}px`;
        invoicePreview.style.height = 'auto';
        invoicePreview.style.padding = '15px';
        invoicePreview.style.border = 'none';
        invoicePreview.style.borderRadius = '0';
        invoicePreview.style.pageBreakInside = 'avoid';
        
        // 要素を画面内に確実に表示させる（上部に配置）
        window.scrollTo(0, 0);
        invoicePreview.scrollIntoView({ behavior: 'auto', block: 'start' });
        
        // レンダリングを待つ（より長い待機時間）
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // 要素の実際のサイズを取得（幅はA4に合わせて設定済み）
        await new Promise(resolve => setTimeout(resolve, 100)); // スタイル変更後の再レンダリングを待つ
        const rect = invoicePreview.getBoundingClientRect();
        const elementWidth = Math.max(invoicePreview.scrollWidth, invoicePreview.offsetWidth, rect.width);
        const elementHeight = Math.max(invoicePreview.scrollHeight, invoicePreview.offsetHeight, rect.height);
        
        console.log('請求書要素の位置:', rect);
        console.log('請求書要素のサイズ:', elementWidth, 'x', elementHeight);
        
        // A4サイズ（mm）とマージンを考慮した利用可能サイズを計算
        // A4: 210mm x 297mm、マージン: 5mm x 4 = 10mm
        const a4HeightMm = 297;
        const availableHeightMm = a4HeightMm - (marginMm * 2); // 287mm
        const availableHeightPx = availableHeightMm * mmToPx; // 約1086.2px
        
        // 要素がA4ページに収まるようにスケールを計算
        // 幅は既にA4に合わせて設定済みなので、高さに基づいてスケールを計算
        const heightScale = availableHeightPx / elementHeight;
        
        // 要素がA4を超える場合、フォントサイズとパディングを動的に縮小
        // 反復的に縮小して、確実に1ページに収まるようにする
        let fontSizeScale = 1.0;
        let paddingScale = 1.0;
        let currentHeight = elementHeight;
        const maxIterations = 5; // 最大5回まで縮小を試行
        
        for (let iteration = 0; iteration < maxIterations && currentHeight > availableHeightPx; iteration++) {
            // 高さがA4を超える場合、縮小率を計算（余裕を持たせる）
            const scaleFactor = (availableHeightPx / currentHeight) * 0.90; // 90%に縮小して余裕を持たせる
            
            // より積極的に縮小（最小50%まで）
            fontSizeScale = Math.max(scaleFactor, 0.5);
            paddingScale = Math.max(scaleFactor * 0.8, 0.4); // パディングはさらに小さく（最小40%まで）
            
            console.log(`縮小試行 ${iteration + 1}: 現在の高さ=${currentHeight}, スケール=${fontSizeScale}`);
            
            // フォントサイズとパディングを動的に調整
            invoicePreview.style.fontSize = `${fontSizeScale * 100}%`;
            invoicePreview.style.padding = `${paddingScale * 15}px`;
            
            // テーブルのパディングも調整
            const tables = invoicePreview.querySelectorAll('table');
            tables.forEach(table => {
                const cells = table.querySelectorAll('th, td');
                cells.forEach(cell => {
                    const originalPadding = 12;
                    cell.style.padding = `${paddingScale * originalPadding}px`;
                });
            });
            
            // ヘッダーやその他の要素のフォントサイズも調整
            const headers = invoicePreview.querySelectorAll('h2');
            headers.forEach(header => {
                const originalFontSize = 2; // em単位
                header.style.fontSize = `${fontSizeScale * originalFontSize}em`;
            });
            
            // 合計金額部分の縮小率を緩和（重要情報なので最小サイズを保つ）
            // 他の要素よりも縮小率を緩和（最小80%まで）
            const importantScale = Math.max(fontSizeScale, 0.8);
            
            // 会社情報部分の縮小率を緩和（最小70%まで）
            const companyInfoScale = Math.max(fontSizeScale, 0.7);
            
            // お客様名、発行元、登録番号、発行日などの重要情報のフォントサイズを調整
            const invoiceInfo = invoicePreview.querySelector('.invoice-info');
            if (invoiceInfo) {
                // お客様名（.invoice-info-left内の最初のp要素）
                const clientNamePara = invoiceInfo.querySelector('.invoice-info-left > p');
                if (clientNamePara) {
                    const computedStyle = window.getComputedStyle(clientNamePara);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 16;
                        const scaledFontSize = importantScale * currentFontSize;
                        clientNamePara.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                }
                
                // 発行元、登録番号、発行日（.invoice-info-right内のp要素）
                const infoRightParas = invoiceInfo.querySelectorAll('.invoice-info-right p');
                infoRightParas.forEach(p => {
                    const computedStyle = window.getComputedStyle(p);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 14;
                        const scaledFontSize = importantScale * currentFontSize;
                        p.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                });
                
                // 会社印鑑セクション内の発行元
                const companySealSection = invoiceInfo.querySelector('.company-seal-section');
                if (companySealSection) {
                    const companySealParas = companySealSection.querySelectorAll('p');
                    companySealParas.forEach(p => {
                        const computedStyle = window.getComputedStyle(p);
                        const currentFontSize = parseFloat(computedStyle.fontSize);
                        if (currentFontSize > 0) {
                            const minFontSize = 14;
                            const scaledFontSize = importantScale * currentFontSize;
                            p.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                        }
                    });
                }
            }
            
            // 請求金額のフォントサイズも調整（重要なので縮小率を緩和）
            const amountSummary = invoicePreview.querySelector('.invoice-amount-summary');
            if (amountSummary) {
                const amountParagraphs = amountSummary.querySelectorAll('p');
                amountParagraphs.forEach(p => {
                    const originalFontSize = 24; // px単位
                    // 最小18pxを保つ
                    const minFontSize = 18;
                    const scaledFontSize = importantScale * originalFontSize;
                    p.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                });
            }
            
            // 合計金額セクションのフォントサイズも調整（重要なので縮小率を緩和）
            const totalSection = invoicePreview.querySelector('.invoice-total');
            if (totalSection) {
                const totalLabels = totalSection.querySelectorAll('.invoice-total-label');
                const totalValues = totalSection.querySelectorAll('.invoice-total-value');
                totalLabels.forEach(el => {
                    const computedStyle = window.getComputedStyle(el);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 14;
                        const scaledFontSize = importantScale * currentFontSize;
                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                });
                totalValues.forEach(el => {
                    const computedStyle = window.getComputedStyle(el);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 14;
                        const scaledFontSize = importantScale * currentFontSize;
                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                });
                // 最終合計金額はさらに大きく
                const finalTotal = totalSection.querySelector('.invoice-total-final');
                if (finalTotal) {
                    const computedStyle = window.getComputedStyle(finalTotal);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 20;
                        const scaledFontSize = importantScale * currentFontSize;
                        finalTotal.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                }
            }
            
            // 会社情報部分（銀行情報と会社情報）のフォントサイズを調整
            const bankInfo = invoicePreview.querySelector('.bank-info');
            if (bankInfo) {
                const bankInfoElements = bankInfo.querySelectorAll('p, h3, span, div');
                bankInfoElements.forEach(el => {
                    const computedStyle = window.getComputedStyle(el);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 12;
                        const scaledFontSize = companyInfoScale * currentFontSize;
                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                });
            }
            
            const companyInfo = invoicePreview.querySelector('.company-info');
            if (companyInfo) {
                const companyInfoElements = companyInfo.querySelectorAll('p, span, div');
                companyInfoElements.forEach(el => {
                    const computedStyle = window.getComputedStyle(el);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 12;
                        const scaledFontSize = companyInfoScale * currentFontSize;
                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                });
            }
            
            const invoiceNumberBottom = invoicePreview.querySelector('.invoice-number-bottom');
            if (invoiceNumberBottom) {
                const invoiceNumberElements = invoiceNumberBottom.querySelectorAll('p, span, div');
                invoiceNumberElements.forEach(el => {
                    const computedStyle = window.getComputedStyle(el);
                    const currentFontSize = parseFloat(computedStyle.fontSize);
                    if (currentFontSize > 0) {
                        const minFontSize = 12;
                        const scaledFontSize = companyInfoScale * currentFontSize;
                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                    }
                });
            }
            
            // すべてのテキスト要素のフォントサイズも調整
            // ただし、重要情報部分は除外（上で個別に処理済み）
            const allTextElements = invoicePreview.querySelectorAll('p, span, div, td, th');
            allTextElements.forEach(el => {
                // 重要情報部分はスキップ
                if (el.closest('.invoice-amount-summary') || 
                    el.closest('.invoice-total') || 
                    el.closest('.invoice-info') ||
                    el.closest('.bank-info') ||
                    el.closest('.company-info') ||
                    el.closest('.invoice-number-bottom')) {
                    return;
                }
                const computedStyle = window.getComputedStyle(el);
                const currentFontSize = parseFloat(computedStyle.fontSize);
                if (currentFontSize > 0) {
                    el.style.fontSize = `${fontSizeScale * currentFontSize}px`;
                }
            });
            
            // マージンも縮小
            const infoSections = invoicePreview.querySelectorAll('.invoice-info, .invoice-header, .invoice-table, .invoice-total, .bank-info');
            infoSections.forEach(section => {
                const computedStyle = window.getComputedStyle(section);
                const currentMarginBottom = parseFloat(computedStyle.marginBottom) || 0;
                if (currentMarginBottom > 0) {
                    section.style.marginBottom = `${paddingScale * currentMarginBottom}px`;
                }
            });
            
            // 再レンダリングを待つ
            await new Promise(resolve => setTimeout(resolve, 300));
            
            // サイズを再取得
            const newRect = invoicePreview.getBoundingClientRect();
            currentHeight = Math.max(invoicePreview.scrollHeight, invoicePreview.offsetHeight, newRect.height);
            console.log(`調整後の要素の高さ: ${currentHeight}`);
            
            // 十分に縮小できた場合は終了
            if (currentHeight <= availableHeightPx) {
                break;
            }
        }
        
        // スケールは1.5から3.0の範囲で、要素が1ページに収まるように調整
        const optimalScale = Math.min(Math.max(heightScale, 1.5), 3.0);
        
        console.log('利用可能サイズ (px):', availableWidthPx, 'x', availableHeightPx);
        console.log('最終的な要素の高さ:', currentHeight);
        console.log('最終的なフォントサイズスケール:', fontSizeScale);
        console.log('最終的なパディングスケール:', paddingScale);
        console.log('最適なスケール:', optimalScale);
        
        // html2pdfのオプション設定（1ページに収める設定）
        const opt = {
            margin: [5, 5, 5, 5],
            filename: `${fileName}.pdf`,
            image: { 
                type: 'jpeg', 
                quality: 0.98 
            },
            html2canvas: { 
                scale: optimalScale,
                useCORS: true,
                logging: false,
                allowTaint: true,
                backgroundColor: '#ffffff',
                removeContainer: false,
                scrollX: 0,
                scrollY: 0,
                onclone: function(clonedDoc, element) {
                    // クローンされたドキュメントのbodyとhtmlのスタイルを調整
                    const clonedBody = clonedDoc.body;
                    const clonedHtml = clonedDoc.documentElement;
                    if (clonedBody) {
                        clonedBody.style.margin = '0';
                        clonedBody.style.padding = '0';
                        clonedBody.style.background = '#ffffff';
                        clonedBody.style.overflow = 'visible';
                    }
                    if (clonedHtml) {
                        clonedHtml.style.margin = '0';
                        clonedHtml.style.padding = '0';
                        clonedHtml.style.overflow = 'visible';
                    }
                    
                    // コンテナ要素のスタイルを調整
                    const clonedContainer = clonedDoc.querySelector('.container');
                    if (clonedContainer) {
                        clonedContainer.style.margin = '0';
                        clonedContainer.style.padding = '0';
                        clonedContainer.style.maxWidth = 'none';
                        clonedContainer.style.width = 'auto';
                        clonedContainer.style.boxShadow = 'none';
                        clonedContainer.style.borderRadius = '0';
                        clonedContainer.style.background = '#ffffff';
                    }
                    
                    // クローンされた要素のスタイルを調整
                    const clonedElement = clonedDoc.getElementById('invoicePreview');
                    if (clonedElement) {
                        // A4サイズに合わせて幅を設定（元の要素と同じ設定）
                        const clonedMmToPx = 3.779527559;
                        const clonedA4WidthMm = 210;
                        const clonedMarginMm = 5;
                        const clonedAvailableWidthMm = clonedA4WidthMm - (clonedMarginMm * 2);
                        const clonedAvailableWidthPx = clonedAvailableWidthMm * clonedMmToPx;
                        
                        clonedElement.style.display = 'block';
                        clonedElement.style.position = 'relative';
                        clonedElement.style.overflow = 'visible';
                        clonedElement.style.maxWidth = `${clonedAvailableWidthPx}px`;
                        clonedElement.style.width = `${clonedAvailableWidthPx}px`;
                        clonedElement.style.height = 'auto';
                        clonedElement.style.margin = '0';
                        clonedElement.style.padding = '15px';
                        clonedElement.style.visibility = 'visible';
                        clonedElement.style.opacity = '1';
                        clonedElement.style.border = 'none';
                        clonedElement.style.borderRadius = '0';
                        clonedElement.style.background = '#ffffff';
                        // ページ分割を防ぐスタイルを追加
                        clonedElement.style.pageBreakInside = 'avoid';
                        clonedElement.style.breakInside = 'avoid';
                        clonedElement.style.pageBreakAfter = 'avoid';
                        clonedElement.style.breakAfter = 'avoid';
                        clonedElement.style.pageBreakBefore = 'avoid';
                        clonedElement.style.breakBefore = 'avoid';
                        
                        // フォントサイズとパディングのスケールを適用
                        if (fontSizeScale < 1.0 || paddingScale < 1.0) {
                            clonedElement.style.fontSize = `${fontSizeScale * 100}%`;
                            clonedElement.style.padding = `${paddingScale * 15}px`;
                            
                            // テーブルのパディングも調整
                            const clonedTables = clonedElement.querySelectorAll('table');
                            clonedTables.forEach(table => {
                                const cells = table.querySelectorAll('th, td');
                                cells.forEach(cell => {
                                    const originalPadding = 12;
                                    cell.style.padding = `${paddingScale * originalPadding}px`;
                                });
                            });
                            
                            // ヘッダーやその他の要素のフォントサイズも調整
                            const clonedHeaders = clonedElement.querySelectorAll('h2');
                            clonedHeaders.forEach(header => {
                                const originalFontSize = 2; // em単位
                                header.style.fontSize = `${fontSizeScale * originalFontSize}em`;
                            });
                            
                            // 合計金額部分の縮小率を緩和（重要情報なので最小サイズを保つ）
                            // 他の要素よりも縮小率を緩和（最小80%まで）
                            const clonedImportantScale = Math.max(fontSizeScale, 0.8);
                            
                            // 会社情報部分の縮小率を緩和（最小70%まで）
                            const clonedCompanyInfoScale = Math.max(fontSizeScale, 0.7);
                            
                            // お客様名、発行元、登録番号、発行日などの重要情報のフォントサイズを調整
                            const clonedInvoiceInfo = clonedElement.querySelector('.invoice-info');
                            if (clonedInvoiceInfo) {
                                // お客様名（.invoice-info-left内の最初のp要素）
                                const clonedClientNamePara = clonedInvoiceInfo.querySelector('.invoice-info-left > p');
                                if (clonedClientNamePara) {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(clonedClientNamePara);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 16;
                                        const scaledFontSize = clonedImportantScale * currentFontSize;
                                        clonedClientNamePara.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                }
                                
                                // 発行元、登録番号、発行日（.invoice-info-right内のp要素）
                                const clonedInfoRightParas = clonedInvoiceInfo.querySelectorAll('.invoice-info-right p');
                                clonedInfoRightParas.forEach(p => {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(p);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 14;
                                        const scaledFontSize = clonedImportantScale * currentFontSize;
                                        p.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                });
                                
                                // 会社印鑑セクション内の発行元
                                const clonedCompanySealSection = clonedInvoiceInfo.querySelector('.company-seal-section');
                                if (clonedCompanySealSection) {
                                    const clonedCompanySealParas = clonedCompanySealSection.querySelectorAll('p');
                                    clonedCompanySealParas.forEach(p => {
                                        const computedStyle = clonedDoc.defaultView.getComputedStyle(p);
                                        const currentFontSize = parseFloat(computedStyle.fontSize);
                                        if (currentFontSize > 0) {
                                            const minFontSize = 14;
                                            const scaledFontSize = clonedImportantScale * currentFontSize;
                                            p.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                        }
                                    });
                                }
                            }
                            
                            // 請求金額のフォントサイズも調整（重要なので縮小率を緩和）
                            const clonedAmountSummary = clonedElement.querySelector('.invoice-amount-summary');
                            if (clonedAmountSummary) {
                                const clonedAmountParagraphs = clonedAmountSummary.querySelectorAll('p');
                                clonedAmountParagraphs.forEach(p => {
                                    const originalFontSize = 24; // px単位
                                    // 最小18pxを保つ
                                    const minFontSize = 18;
                                    const scaledFontSize = clonedImportantScale * originalFontSize;
                                    p.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                });
                            }
                            
                            // 合計金額セクションのフォントサイズも調整（重要なので縮小率を緩和）
                            const clonedTotalSection = clonedElement.querySelector('.invoice-total');
                            if (clonedTotalSection) {
                                const clonedTotalLabels = clonedTotalSection.querySelectorAll('.invoice-total-label');
                                const clonedTotalValues = clonedTotalSection.querySelectorAll('.invoice-total-value');
                                clonedTotalLabels.forEach(el => {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(el);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 14;
                                        const scaledFontSize = clonedImportantScale * currentFontSize;
                                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                });
                                clonedTotalValues.forEach(el => {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(el);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 14;
                                        const scaledFontSize = clonedImportantScale * currentFontSize;
                                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                });
                                // 最終合計金額はさらに大きく
                                const clonedFinalTotal = clonedTotalSection.querySelector('.invoice-total-final');
                                if (clonedFinalTotal) {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(clonedFinalTotal);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 20;
                                        const scaledFontSize = clonedImportantScale * currentFontSize;
                                        clonedFinalTotal.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                }
                            }
                            
                            // 会社情報部分（銀行情報と会社情報）のフォントサイズを調整
                            const clonedBankInfo = clonedElement.querySelector('.bank-info');
                            if (clonedBankInfo) {
                                const clonedBankInfoElements = clonedBankInfo.querySelectorAll('p, h3, span, div');
                                clonedBankInfoElements.forEach(el => {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(el);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 12;
                                        const scaledFontSize = clonedCompanyInfoScale * currentFontSize;
                                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                });
                            }
                            
                            const clonedCompanyInfo = clonedElement.querySelector('.company-info');
                            if (clonedCompanyInfo) {
                                const clonedCompanyInfoElements = clonedCompanyInfo.querySelectorAll('p, span, div');
                                clonedCompanyInfoElements.forEach(el => {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(el);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 12;
                                        const scaledFontSize = clonedCompanyInfoScale * currentFontSize;
                                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                });
                            }
                            
                            const clonedInvoiceNumberBottom = clonedElement.querySelector('.invoice-number-bottom');
                            if (clonedInvoiceNumberBottom) {
                                const clonedInvoiceNumberElements = clonedInvoiceNumberBottom.querySelectorAll('p, span, div');
                                clonedInvoiceNumberElements.forEach(el => {
                                    const computedStyle = clonedDoc.defaultView.getComputedStyle(el);
                                    const currentFontSize = parseFloat(computedStyle.fontSize);
                                    if (currentFontSize > 0) {
                                        const minFontSize = 12;
                                        const scaledFontSize = clonedCompanyInfoScale * currentFontSize;
                                        el.style.fontSize = `${Math.max(scaledFontSize, minFontSize)}px`;
                                    }
                                });
                            }
                            
                            // すべてのテキスト要素のフォントサイズも調整（クローンされたドキュメントのコンテキストを使用）
                            // ただし、重要情報部分は除外（上で個別に処理済み）
                            const clonedTextElements = clonedElement.querySelectorAll('p, span, div, td, th');
                            clonedTextElements.forEach(el => {
                                // 重要情報部分はスキップ
                                if (el.closest('.invoice-amount-summary') || 
                                    el.closest('.invoice-total') || 
                                    el.closest('.invoice-info') ||
                                    el.closest('.bank-info') ||
                                    el.closest('.company-info') ||
                                    el.closest('.invoice-number-bottom')) {
                                    return;
                                }
                                // クローンされたドキュメントのコンテキストでgetComputedStyleを使用
                                const computedStyle = clonedDoc.defaultView.getComputedStyle(el);
                                const currentFontSize = parseFloat(computedStyle.fontSize);
                                if (currentFontSize > 0) {
                                    el.style.fontSize = `${fontSizeScale * currentFontSize}px`;
                                }
                            });
                            
                            // マージンも縮小
                            const clonedInfoSections = clonedElement.querySelectorAll('.invoice-info, .invoice-header, .invoice-table, .invoice-total, .bank-info');
                            clonedInfoSections.forEach(section => {
                                const computedStyle = clonedDoc.defaultView.getComputedStyle(section);
                                const currentMarginBottom = parseFloat(computedStyle.marginBottom) || 0;
                                if (currentMarginBottom > 0) {
                                    section.style.marginBottom = `${paddingScale * currentMarginBottom}px`;
                                }
                            });
                        }
                        
                        // すべての子要素も確実に表示
                        const allChildren = clonedElement.querySelectorAll('*');
                        allChildren.forEach(child => {
                            child.style.visibility = 'visible';
                            child.style.opacity = '1';
                            // テーブルやその他の要素のスタイルも調整
                            if (child.tagName === 'TABLE') {
                                child.style.width = '100%';
                                child.style.borderCollapse = 'collapse';
                            }
                            // ページ分割を防ぐスタイルを追加
                            child.style.pageBreakInside = 'avoid';
                            child.style.breakInside = 'avoid';
                        });
                    }
                    
                    // ボタンを非表示にする
                    const clonedButton = clonedElement ? clonedElement.querySelector('.payment-received-btn') : null;
                    if (clonedButton) {
                        clonedButton.style.display = 'none';
                    }
                    
                    // フォーム部分を非表示にする
                    const formSections = clonedDoc.querySelectorAll('.form-section, #generateBtn, #printBtn, #addItemBtn, h1');
                    formSections.forEach(el => {
                        el.style.display = 'none';
                    });
                }
            },
            jsPDF: { 
                unit: 'mm', 
                format: 'a4',
                orientation: 'portrait',
                compress: true,
                putOnlyUsedFonts: true,
                floatPrecision: 16
            },
            pagebreak: { 
                mode: 'avoid-all'
            }
        };
        
        // PDFを生成して保存（1ページに収める）
        await html2pdf().set(opt).from(invoicePreview).save();
        
        console.log('PDF保存が完了しました');
    } catch (error) {
        console.error('PDF保存エラー:', error);
        console.error('エラー詳細:', error.stack);
        alert('PDFの保存に失敗しました: ' + error.message + '\nブラウザのコンソールを確認してください。');
    } finally {
        // 親要素のスタイルを戻す
        if (parentContainer) {
            if (originalParentMaxWidth) {
                parentContainer.style.maxWidth = originalParentMaxWidth;
            } else {
                parentContainer.style.maxWidth = '';
            }
            if (originalParentOverflow) {
                parentContainer.style.overflow = originalParentOverflow;
            } else {
                parentContainer.style.overflow = '';
            }
        }
        
        // 元のスタイルに戻す
        invoicePreview.style.display = originalDisplay;
        if (originalPosition) invoicePreview.style.position = originalPosition;
        if (originalMargin) invoicePreview.style.margin = originalMargin;
        if (originalOverflow) invoicePreview.style.overflow = originalOverflow;
        if (originalMaxWidth) invoicePreview.style.maxWidth = originalMaxWidth;
        if (originalWidth) invoicePreview.style.width = originalWidth;
        if (originalHeight) invoicePreview.style.height = originalHeight;
        if (originalPadding) invoicePreview.style.padding = originalPadding;
        if (originalBorder) invoicePreview.style.border = originalBorder;
        if (originalBorderRadius) invoicePreview.style.borderRadius = originalBorderRadius;
        
        // ボタンの表示状態を戻す
        if (paymentButton && originalButtonDisplay) {
            paymentButton.style.display = originalButtonDisplay;
        } else if (paymentButton) {
            paymentButton.style.display = '';
        }
    }
}

// 領収書をPDFとして保存
async function saveReceiptAsPDF() {
    const receiptPreview = document.getElementById('receiptPreview');
    if (!receiptPreview || receiptPreview.style.display === 'none') {
        alert('先に領収書を生成してください。');
        return;
    }
    
    // 領収書番号を取得してファイル名に使用
    const receiptNumberElement = receiptPreview.querySelector('.receipt-number-bottom p');
    let fileName = '領収書';
    if (receiptNumberElement) {
        const receiptNumber = receiptNumberElement.textContent.replace('領収書番号:', '').trim();
        if (receiptNumber) {
            fileName = `領収書_${receiptNumber}`;
        }
    }
    
    // 元のスタイルを保存
    const originalDisplay = receiptPreview.style.display;
    const originalPosition = receiptPreview.style.position || '';
    const originalLeft = receiptPreview.style.left || '';
    const originalTop = receiptPreview.style.top || '';
    const originalMargin = receiptPreview.style.margin || '';
    const originalOverflow = receiptPreview.style.overflow || '';
    const originalMaxWidth = receiptPreview.style.maxWidth || '';
    const originalWidth = receiptPreview.style.width || '';
    const originalHeight = receiptPreview.style.height || '';
    
    try {
        // 要素を確実に表示状態にする
        receiptPreview.style.display = 'block';
        receiptPreview.style.position = 'relative';
        receiptPreview.style.left = 'auto';
        receiptPreview.style.top = 'auto';
        receiptPreview.style.margin = '0';
        receiptPreview.style.overflow = 'visible';
        receiptPreview.style.maxWidth = 'none';
        receiptPreview.style.width = 'auto';
        receiptPreview.style.height = 'auto';
        
        // レンダリングを待つ
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // 要素の実際のサイズを取得
        const rect = receiptPreview.getBoundingClientRect();
        const elementHeight = Math.max(receiptPreview.scrollHeight, receiptPreview.offsetHeight, rect.height);
        const elementWidth = Math.max(receiptPreview.scrollWidth, receiptPreview.offsetWidth, rect.width);
        
        console.log('要素サイズ:', elementWidth, 'x', elementHeight);
        
        // html2pdfのオプション設定
        const opt = {
            margin: [5, 5, 5, 5],
            filename: `${fileName}.pdf`,
            image: { 
                type: 'jpeg', 
                quality: 0.98 
            },
            html2canvas: { 
                scale: 2,
                useCORS: true,
                logging: false,
                allowTaint: true,
                backgroundColor: '#ffffff',
                removeContainer: false,
                onclone: function(clonedDoc) {
                    // クローンされた要素のスタイルを調整
                    const clonedElement = clonedDoc.getElementById('receiptPreview');
                    if (clonedElement) {
                        clonedElement.style.display = 'block';
                        clonedElement.style.position = 'relative';
                        clonedElement.style.overflow = 'visible';
                        clonedElement.style.maxWidth = 'none';
                        clonedElement.style.width = 'auto';
                        clonedElement.style.height = 'auto';
                    }
                }
            },
            jsPDF: { 
                unit: 'mm', 
                format: 'a4',
                orientation: 'portrait',
                compress: true
            },
            pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
        };
        
        // PDFを生成して保存
        await html2pdf().set(opt).from(receiptPreview).save();
        
        console.log('PDF保存が完了しました');
    } catch (error) {
        console.error('PDF保存エラー:', error);
        console.error('エラー詳細:', error.stack);
        alert('PDFの保存に失敗しました: ' + error.message + '\nブラウザのコンソールを確認してください。');
    } finally {
        // 元のスタイルに戻す
        receiptPreview.style.display = originalDisplay;
        if (originalPosition) receiptPreview.style.position = originalPosition;
        if (originalLeft) receiptPreview.style.left = originalLeft;
        if (originalTop) receiptPreview.style.top = originalTop;
        if (originalMargin) receiptPreview.style.margin = originalMargin;
        if (originalOverflow) receiptPreview.style.overflow = originalOverflow;
        if (originalMaxWidth) receiptPreview.style.maxWidth = originalMaxWidth;
        if (originalWidth) receiptPreview.style.width = originalWidth;
        if (originalHeight) receiptPreview.style.height = originalHeight;
    }
}

// 履歴に請求書を保存
function saveInvoiceToHistory(invoiceData) {
    try {
        const historyKey = 'invoice_history';
        let history = JSON.parse(localStorage.getItem(historyKey) || '[]');
        
        // 新しい履歴を先頭に追加
        history.unshift(invoiceData);
        
        // 最大100件まで保存（古いものから削除）
        if (history.length > 100) {
            history = history.slice(0, 100);
        }
        
        localStorage.setItem(historyKey, JSON.stringify(history));
        
        // 履歴一覧を更新
        displayInvoiceHistory();
    } catch (error) {
        console.error('履歴の保存に失敗しました:', error);
    }
}

// 履歴一覧を取得
function getInvoiceHistory() {
    try {
        const historyKey = 'invoice_history';
        return JSON.parse(localStorage.getItem(historyKey) || '[]');
    } catch (error) {
        console.error('履歴の読み込みに失敗しました:', error);
        return [];
    }
}

// 履歴一覧を表示
function displayInvoiceHistory() {
    const historyContainer = document.getElementById('invoiceHistoryContainer');
    if (!historyContainer) return;
    
    const history = getInvoiceHistory();
    
    if (history.length === 0) {
        historyContainer.innerHTML = '<p class="no-history">履歴がありません</p>';
        return;
    }
    
    let historyHTML = '<div class="history-list">';
    history.forEach((invoice, index) => {
        const formattedDate = formatDate(invoice.issueDate);
        const createdAt = invoice.createdAt ? new Date(invoice.createdAt).toLocaleString('ja-JP') : '';
        
        historyHTML += `
            <div class="history-item" data-index="${index}">
                <div class="history-item-header">
                    <div class="history-item-title">
                        <strong>${escapeHtml(invoice.invoiceNumber)}</strong>
                        <span class="history-date">${formattedDate}</span>
                    </div>
                    <div class="history-item-actions">
                        <button class="history-btn view-btn" onclick="loadInvoiceFromHistory(${index})">表示</button>
                        <button class="history-btn pdf-btn" onclick="regeneratePDFFromHistory(${index})">PDF</button>
                        <button class="history-btn delete-btn" onclick="deleteInvoiceFromHistory(${index})">削除</button>
                    </div>
                </div>
                <div class="history-item-details">
                    <p><strong>お客様:</strong> ${escapeHtml(invoice.clientName)}</p>
                    <p><strong>請求金額:</strong> ¥${formatNumber(invoice.total)}</p>
                    <p><strong>項目数:</strong> ${invoice.items.length}件</p>
                    ${createdAt ? `<p class="history-created-at">作成日時: ${createdAt}</p>` : ''}
                </div>
            </div>
        `;
    });
    historyHTML += '</div>';
    
    historyContainer.innerHTML = historyHTML;
}

// 履歴から請求書を読み込んで表示
function loadInvoiceFromHistory(index) {
    const history = getInvoiceHistory();
    if (index < 0 || index >= history.length) {
        alert('履歴が見つかりません');
        return;
    }
    
    const invoiceData = history[index];
    
    // フォームにデータを設定
    document.getElementById('issueDate').value = invoiceData.issueDate;
    
    // お客様名を解析して設定
    const clientNameMatch = invoiceData.clientName.match(/^(株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|一般財団法人|医療法人|学校法人|宗教法人|NPO法人|その他)?(.+)$/);
    if (clientNameMatch) {
        const companyType = clientNameMatch[1] || '';
        const name = clientNameMatch[2] || invoiceData.clientName;
        document.getElementById('clientCompanyType').value = companyType;
        document.getElementById('clientName').value = name;
    } else {
        document.getElementById('clientCompanyType').value = '';
        document.getElementById('clientName').value = invoiceData.clientName;
    }
    
    // 項目を設定
    const itemsContainer = document.getElementById('itemsContainer');
    itemsContainer.innerHTML = '';
    
    invoiceData.items.forEach((item, itemIndex) => {
        const itemRow = document.createElement('div');
        itemRow.className = 'item-row';
        itemRow.innerHTML = `
            <input type="date" class="item-date" value="${item.date}" required>
            <input type="text" class="item-description" value="${escapeHtml(item.description)}" placeholder="項目名" required>
            <input type="number" class="item-amount" value="${item.amount}" placeholder="金額" min="0" step="1" required>
            <select class="item-tax-rate">
                <option value="0" ${item.taxRate === 0 ? 'selected' : ''}>0%</option>
                <option value="8" ${item.taxRate === 8 ? 'selected' : ''}>8%</option>
                <option value="10" ${item.taxRate === 10 ? 'selected' : ''}>10%</option>
            </select>
            <select class="item-amount-type">
                <option value="exclusive" ${item.amountType === 'exclusive' ? 'selected' : ''}>税抜</option>
                <option value="inclusive" ${item.amountType === 'inclusive' ? 'selected' : ''}>税込</option>
            </select>
            <button type="button" class="remove-item-btn" onclick="removeItem(this)" ${invoiceData.items.length === 1 ? 'style="display:none;"' : ''}>削除</button>
        `;
        itemsContainer.appendChild(itemRow);
    });
    
    updateRemoveButtons();
    
    // 請求書を再生成
    generateInvoice();
    
    // 履歴セクションを閉じる
    const historySection = document.getElementById('invoiceHistorySection');
    if (historySection) {
        historySection.style.display = 'none';
    }
}

// 履歴からPDFを再生成
async function regeneratePDFFromHistory(index) {
    const history = getInvoiceHistory();
    if (index < 0 || index >= history.length) {
        alert('履歴が見つかりません');
        return;
    }
    
    const invoiceData = history[index];
    
    // フォームにデータを設定
    document.getElementById('issueDate').value = invoiceData.issueDate;
    
    // お客様名を解析して設定
    const clientNameMatch = invoiceData.clientName.match(/^(株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|一般財団法人|医療法人|学校法人|宗教法人|NPO法人|その他)?(.+)$/);
    if (clientNameMatch) {
        const companyType = clientNameMatch[1] || '';
        const name = clientNameMatch[2] || invoiceData.clientName;
        document.getElementById('clientCompanyType').value = companyType;
        document.getElementById('clientName').value = name;
    } else {
        document.getElementById('clientCompanyType').value = '';
        document.getElementById('clientName').value = invoiceData.clientName;
    }
    
    // 項目を設定
    const itemsContainer = document.getElementById('itemsContainer');
    itemsContainer.innerHTML = '';
    
    invoiceData.items.forEach((item) => {
        const itemRow = document.createElement('div');
        itemRow.className = 'item-row';
        itemRow.innerHTML = `
            <input type="date" class="item-date" value="${item.date}" required>
            <input type="text" class="item-description" value="${escapeHtml(item.description)}" placeholder="項目名" required>
            <input type="number" class="item-amount" value="${item.amount}" placeholder="金額" min="0" step="1" required>
            <select class="item-tax-rate">
                <option value="0" ${item.taxRate === 0 ? 'selected' : ''}>0%</option>
                <option value="8" ${item.taxRate === 8 ? 'selected' : ''}>8%</option>
                <option value="10" ${item.taxRate === 10 ? 'selected' : ''}>10%</option>
            </select>
            <select class="item-amount-type">
                <option value="exclusive" ${item.amountType === 'exclusive' ? 'selected' : ''}>税抜</option>
                <option value="inclusive" ${item.amountType === 'inclusive' ? 'selected' : ''}>税込</option>
            </select>
            <button type="button" class="remove-item-btn" onclick="removeItem(this)" ${invoiceData.items.length === 1 ? 'style="display:none;"' : ''}>削除</button>
        `;
        itemsContainer.appendChild(itemRow);
    });
    
    updateRemoveButtons();
    
    // 請求書を再生成して表示
    await generateInvoice();
    
    // 少し待ってからPDF保存を実行
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // PDF保存を実行
    await saveInvoiceAsPDF();
}

// 履歴から請求書を削除
function deleteInvoiceFromHistory(index) {
    if (!confirm('この履歴を削除しますか？')) {
        return;
    }
    
    try {
        const historyKey = 'invoice_history';
        let history = JSON.parse(localStorage.getItem(historyKey) || '[]');
        
        if (index >= 0 && index < history.length) {
            history.splice(index, 1);
            localStorage.setItem(historyKey, JSON.stringify(history));
            displayInvoiceHistory();
        }
    } catch (error) {
        console.error('履歴の削除に失敗しました:', error);
        alert('履歴の削除に失敗しました');
    }
}

// 履歴セクションの表示/非表示を切り替え
function toggleHistorySection() {
    const historySection = document.getElementById('invoiceHistorySection');
    if (historySection) {
        if (historySection.style.display === 'none') {
            historySection.style.display = 'block';
            displayInvoiceHistory();
        } else {
            historySection.style.display = 'none';
        }
    }
}

// ページ読み込み時に今日の日付を設定
window.onload = function() {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('issueDate').value = today;
    // 最初の項目の日付も設定
    const firstItemDate = document.getElementById('firstItemDate');
    if (firstItemDate) {
        firstItemDate.value = today;
    }
    
    // 履歴一覧を表示
    displayInvoiceHistory();
};

