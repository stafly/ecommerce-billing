// --- 初始化默认规则表 ---
const defaultRuleData = [
    { country: 'poland', items: '4060*3', channel: '自提客选', price: 11.0 },
    { country: 'polan', items: '4060*3, 3040*2', channel: '!自提客选', price: 25.0 },
    { country: 'poland', items: '4060*5, 3040*1', channel: '', price: 30.0 }
];

let orderRuleConfig = [];
let pendingFiles = [];
let processedZips = null;

// --- 分页控制参数 ---
let currentPage = 1;
let pageSize = 50; 

// --- 日志系统 ---
const logTerminal = document.getElementById('log-content');
function log(msg, type = 'log-info') {
    const span = document.createElement('span');
    span.className = type;
    span.innerText = msg + '\n';
    logTerminal.appendChild(span);
    logTerminal.parentElement.scrollTop = logTerminal.parentElement.scrollHeight;
}
function clearLog() {
    logTerminal.innerHTML = '';
}

// --- 表格渲染与交互 ---
const tableBody = document.getElementById('price-table-body');
const paginationControls = document.getElementById('pagination-controls');
const prevBtn = document.getElementById('prev-page-btn');
const nextBtn = document.getElementById('next-page-btn');
const pageInfo = document.getElementById('page-info');
const pageSizeSelect = document.getElementById('page-size-select');

function renderTable() {
    tableBody.innerHTML = '';
    
    // ----------- 1. 冲突检测 -----------
    const signatureMap = {};
    const conflictIndexes = new Set();
    
    orderRuleConfig.forEach((row, index) => {
        // 签名: 国家 + 商品 + 渠道 决定了唯一触发条件
        const sig = `${(row.country||'').trim().toLowerCase()}|||${(row.items||'').trim()}|||${(row.channel||'').trim()}`;
        if (!signatureMap[sig]) {
            signatureMap[sig] = [index];
        } else {
            signatureMap[sig].push(index);
        }
    });
    
    for (let sig in signatureMap) {
        if (signatureMap[sig].length > 1) {
            signatureMap[sig].forEach(idx => conflictIndexes.add(idx));
        }
    }
    
    // ----------- 2. 分页计算 -----------
    let totalItems = orderRuleConfig.length;
    let totalPages = pageSize === 'all' ? 1 : Math.ceil(totalItems / pageSize);
    if (currentPage > totalPages && totalPages > 0) currentPage = totalPages;
    
    let startIndex = 0;
    let endIndex = totalItems;
    
    if (pageSize !== 'all') {
        startIndex = (currentPage - 1) * pageSize;
        endIndex = Math.min(startIndex + pageSize, totalItems);
        paginationControls.style.display = 'flex';
        prevBtn.disabled = currentPage === 1;
        nextBtn.disabled = currentPage === totalPages || totalPages === 0;
        pageInfo.innerText = `第 ${totalItems === 0 ? 0 : currentPage} / ${totalPages} 页 (共 ${totalItems} 条)`;
    } else {
        paginationControls.style.display = 'none';
    }

    // ----------- 3. 渲染视图 -----------
    for (let i = startIndex; i < endIndex; i++) {
        const row = orderRuleConfig[i];
        const isConflict = conflictIndexes.has(i);
        const tr = document.createElement('tr');
        if (isConflict) tr.className = 'conflict-row';
        tr.title = isConflict ? "⚠️ 此规则的条件与其它行完全重复，可能导致金额被错误覆盖！" : "";
        
        tr.innerHTML = `
            <td><input type="text" value="${row.country}" data-row="${i}" data-field="country" class="rule-input" style="width:100%"></td>
            <td><input type="text" value="${row.items}" data-row="${i}" data-field="items" class="rule-input" style="width:100%" placeholder="4060*3, 3040*2"></td>
            <td><input type="text" value="${row.channel}" data-row="${i}" data-field="channel" class="rule-input" style="width:100%" placeholder="!排除特定渠道 或 留空"></td>
            <td><input type="number" step="0.01" min="0" value="${row.price}" data-row="${i}" data-field="price" class="rule-input" style="width:80px;"></td>
            <td><button class="btn btn-delete" data-row="${i}">🗑️</button></td>
        `;
        tableBody.appendChild(tr);
    }
    
    document.querySelectorAll('.rule-input').forEach(input => {
        input.addEventListener('change', handleInputChange);
    });
    document.querySelectorAll('.btn-delete').forEach(btn => {
        btn.addEventListener('click', handleDeleteRow);
    });
}

// 分页事件绑定
prevBtn.addEventListener('click', () => {
    if (currentPage > 1) { currentPage--; renderTable(); }
});
nextBtn.addEventListener('click', () => {
    const totalPages = pageSize === 'all' ? 1 : Math.ceil(orderRuleConfig.length / pageSize);
    if (currentPage < totalPages) { currentPage++; renderTable(); }
});
pageSizeSelect.addEventListener('change', (e) => {
    pageSize = e.target.value === 'all' ? 'all' : parseInt(e.target.value);
    currentPage = 1;
    renderTable();
});

function handleInputChange(e) {
    const el = e.target;
    const rowIndex = parseInt(el.getAttribute('data-row'));
    const field = el.getAttribute('data-field');
    const val = el.value;
    
    if (field === 'price') {
        orderRuleConfig[rowIndex][field] = parseSafePrice(val);
    } else {
        orderRuleConfig[rowIndex][field] = val;
    }
    saveConfigToLocal();
}

function handleDeleteRow(e) {
    const el = e.target.closest('button');
    const rowIndex = parseInt(el.getAttribute('data-row'));
    orderRuleConfig.splice(rowIndex, 1);
    saveConfigToLocal();
    renderTable();
}

document.getElementById('add-row-btn').addEventListener('click', () => {
    orderRuleConfig.push({ country: '', items: '', channel: '', price: 0.0 });
    saveConfigToLocal();
    renderTable();
});

document.getElementById('reset-default-btn').addEventListener('click', () => {
    if(confirm("确定要清空并恢复默认规则表吗？当前修改将会全部丢失。")) {
        orderRuleConfig = JSON.parse(JSON.stringify(defaultRuleData));
        saveConfigToLocal();
        renderTable();
    }
});
function safeStr(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim();
}

/**
 * 修复价格录入错误（如多个连续小数点 10..2，或夹杂非数字字符等）
 * @param {string|number} val - 原始价格输入
 * @returns {number} 清洗并解析后的安全浮点数价格
 */
function parseSafePrice(val) {
    if (val === null || val === undefined || val === '') return 0.0;
    let strVal = String(val).trim();
    // 替换多个连续小数点为一个
    strVal = strVal.replace(/\.{2,}/g, '.');
    // 去除除数字和小数点外的其他非法字符
    strVal = strVal.replace(/[^\d.]/g, ''); 
    // 防止形如 10.2.3 的离散多次小数点问题，强制只保留第一个小数点
    const parts = strVal.split('.');
    if (parts.length > 2) {
        strVal = parts[0] + '.' + parts.slice(1).join('');
    }
    return parseFloat(strVal) || 0.0;
}

// --- 规则表 Excel 导入逻辑 ---
const importRulesInput = document.getElementById('import-rules-input');
importRulesInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const wb = XLSX.read(event.target.result, {type: 'array'});
            const sheetName = wb.SheetNames[0];
            const worksheet = wb.Sheets[sheetName];
            const rawJson = XLSX.utils.sheet_to_json(worksheet, {defval: ''});
            
            let newRules = [];
            rawJson.forEach(row => {
                const cleanRow = {};
                for(let k in row) {
                    cleanRow[k.trim()] = row[k];
                }
                
                const country = safeStr(cleanRow['国家'] || cleanRow['Country'] || '');
                const items = safeStr(cleanRow['包含商品及数量'] || cleanRow['包含商品和数量'] || cleanRow['商品及数量'] || cleanRow['包含商品'] || '');
                const channel = safeStr(cleanRow['物流渠道'] || cleanRow['Channel'] || '');
                const price = parseSafePrice(cleanRow['整单总价'] || cleanRow['一口价'] || cleanRow['价格'] || cleanRow['总价']);
                
                if (items || country) {
                    newRules.push({ country, items, channel, price });
                }
            });
            
            if (newRules.length > 0) {
                if(confirm(`成功读取 ${newRules.length} 条规则！\n\n点击【确定】将覆盖当前界面的所有规则。\n点击【取消】将追加到当前列表末尾。`)) {
                    orderRuleConfig = newRules;
                } else {
                    orderRuleConfig = orderRuleConfig.concat(newRules);
                }
                saveConfigToLocal();
                renderTable();
                log(`✅ 已成功导入 ${newRules.length} 条价格规则`, 'log-success');
            } else {
                alert("未在表格中读取到有效规则！请确保Excel至少包含列名：'国家', '包含商品及数量', '物流渠道', '整单总价'。");
            }
        } catch (err) {
            alert("读取 Excel 配置表失败：" + err.message);
        }
        importRulesInput.value = '';
    };
    reader.readAsArrayBuffer(file);
});

function saveConfigToLocal() {
    localStorage.setItem('ecommerce_order_rules', JSON.stringify(orderRuleConfig));
}

// 初始化数据
const savedConfig = localStorage.getItem('ecommerce_order_rules');
if (savedConfig) {
    try {
        orderRuleConfig = JSON.parse(savedConfig);
    } catch (e) {
        orderRuleConfig = JSON.parse(JSON.stringify(defaultRuleData));
    }
} else {
    orderRuleConfig = JSON.parse(JSON.stringify(defaultRuleData));
}
renderTable();

// --- 文件上传逻辑 ---
const fileInput = document.getElementById('file-input');
const dropZone = document.getElementById('drop-zone');
const fileListEl = document.getElementById('file-list');
const fileBadge = document.getElementById('file-count-badge');
const startBtn = document.getElementById('start-process-btn');
const downloadBtn = document.getElementById('download-btn');

function handleFiles(files) {
    pendingFiles = Array.from(files).filter(f => f.name.includes('.xls') && !f.name.startsWith('~$'));
    updateFileList();
}

fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files) handleFiles(e.dataTransfer.files);
});

function updateFileList() {
    fileListEl.innerHTML = '';
    if (pendingFiles.length > 0) {
        fileBadge.innerText = `已选择 ${pendingFiles.length} 个文件`;
        startBtn.disabled = false;
        pendingFiles.forEach(f => {
            const li = document.createElement('li');
            li.innerHTML = `<span>📄 ${f.name}</span> <span style="color:#888;">${(f.size/1024).toFixed(1)} KB</span>`;
            fileListEl.appendChild(li);
        });
    } else {
        fileBadge.innerText = '未选择文件';
        startBtn.disabled = true;
    }
    downloadBtn.style.display = 'none';
}

// --- 核心业务逻辑 (委托给内联 Web Worker 处理解决本地 file:// 跨域限制) ---

function getWorkerCode() {
    return `
importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
importScripts('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');

// --- 通信日志封装 ---
function logMain(msg, type = 'log-info') {
    postMessage({ type: 'log', message: msg, logType: type });
}

self.onmessage = async function(e) {
    const { filesData, orderRuleConfig } = e.data;
    
    try {
        logMain('⚡ [后台线程] Web Worker 多线程处理引擎已启动...', 'log-info');
        const priceDict = orderRuleConfig;
        const finalResults = [];
        
        for (let i = 0; i < filesData.length; i++) {
            const fileObj = filesData[i];
            const wb = XLSX.read(fileObj.buffer, {type: 'array'});
            const data = readExcelDataWithFuzzyHeaders(wb);
            const processed = processDataAndVerify(data, fileObj.name, priceDict);
            processed.fileName = fileObj.name;
            finalResults.push(processed);
        }
        
        const zip = exportExcelAndZip(finalResults);
        const zipBlob = await zip.generateAsync({type:"blob"});
        
        logMain('\\n🎁 处理完毕！正在打包生成 ZIP 文件...', 'log-success');
        postMessage({ type: 'done', zipBlob: zipBlob });
        
    } catch (err) {
        logMain(\`❌ [后台线程] 发生致命错误: \${err.message}\`, 'log-error');
        console.error(err);
        postMessage({ type: 'error', error: err.message });
    }
};

// --- 重构的智能表头读取逻辑 (Fuzzy Headers) ---
function readExcelDataWithFuzzyHeaders(workbook) {
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawJson = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
    if (rawJson.length === 0) return [];
    
    const headers = rawJson[0] || [];
    const headerMap = {}; 
    
    const aliasMap = {
        '交易编号': ['交易编号', '订单号', '订单编号', 'transaction id', 'order id'],
        '客户姓名': ['客户姓名', '买家姓名', '收件人姓名', 'buyer name', 'customer name'],
        '国家': ['国家', '收件人国家', '目标国家', 'country', 'destination country'],
        '平台SKU多属性': ['平台sku多属性', 'sku属性', '商品属性', 'sku'],
        '订单商品名称': ['订单商品名称', '商品名称', '产品名称', 'item name', 'product name'],
        '商品数量': ['商品数量', '购买数量', '数量', 'qty', 'quantity'],
        '物流渠道': ['物流渠道', '发货方式', '指定发货方式', 'shipping method', 'logistics']
    };
    
    headers.forEach((h, idx) => {
        if (!h) return;
        const lowerH = String(h).trim().toLowerCase();
        for (let standardKey in aliasMap) {
            if (aliasMap[standardKey].some(alias => lowerH.includes(alias) || alias === lowerH)) {
                if (!headerMap[standardKey]) {
                    headerMap[standardKey] = h;
                }
            }
        }
    });

    const dataJson = XLSX.utils.sheet_to_json(worksheet, {defval: ''});
    return dataJson.map(row => {
        const cleanRow = {};
        for(let k in row) { cleanRow[k.trim()] = row[k]; }
        
        const mappedRow = {};
        for (let standardKey in aliasMap) {
            const originalCol = headerMap[standardKey];
            mappedRow[standardKey] = originalCol ? cleanRow[originalCol] : '';
        }
        return mappedRow;
    });
}

function parseItemsToMap(itemStr) {
    const map = {};
    if (!itemStr) return map;
    const parts = String(itemStr).split(/,|，/);
    parts.forEach(p => {
        p = p.trim();
        if (!p) return;
        const match = p.match(/(.+?)\\s*(?:\\*|x|X|×)\\s*(\\d+)$/i);
        if (match) {
            const name = match[1].trim();
            const qty = parseInt(match[2]);
            map[name] = (map[name] || 0) + qty;
        } else {
            map[p] = (map[p] || 0) + 1;
        }
    });
    return map;
}

function safeStr(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim();
}

function isItemMapMatch(orderMap, ruleMap) {
    const orderKeys = Object.keys(orderMap);
    const ruleKeys = Object.keys(ruleMap);
    if (orderKeys.length !== ruleKeys.length) return false;
    for (let k of ruleKeys) {
        if (orderMap[k] !== ruleMap[k]) {
            let foundMatch = false;
            for(let ok of orderKeys) {
                if(ok.includes(k) && orderMap[ok] === ruleMap[k]) {
                    foundMatch = true; break;
                }
            }
            if (!foundMatch) return false;
        }
    }
    return true;
}

function preprocessRow(row) {
    let name = safeStr(row['订单商品名称']);
    if (name === '' || name.toLowerCase() === 'nan') {
        const skuAttr = safeStr(row['平台SKU多属性']);
        if (skuAttr.includes(';')) {
            name = skuAttr.substring(0, skuAttr.lastIndexOf(';')).trim();
        } else {
            name = skuAttr;
        }
    }
    row['最终商品名称'] = name;
    
    const skuAttr2 = safeStr(row['平台SKU多属性']);
    if (skuAttr2.includes(';')) {
        row['最终尺码'] = skuAttr2.split(';').pop().trim();
    } else {
        row['最终尺码'] = '默认';
    }
    row['解析数量'] = parseFloat(row['商品数量']) || 0;
}

function processDataAndVerify(data, fileName, orderRuleConfig) {
    let allPassed = true;
    let origRowCount = data.length;
    let origQtySum = 0;
    
    logMain(\`\\n📊 正在后台处理并比对: [\${fileName}]\`);
    
    const orderGroups = {}; 
    data.forEach(row => {
        preprocessRow(row);
        origQtySum += row['解析数量'];
        
        const optKey = \`\${safeStr(row['交易编号'])}|||\${safeStr(row['客户姓名'])}\`;
        if (!orderGroups[optKey]) {
            orderGroups[optKey] = {
                rows: [], country: safeStr(row['国家']), channel: safeStr(row['物流渠道']),
                itemsMap: {}, matchedPrice: 0, isMatched: false
            };
        }
        orderGroups[optKey].rows.push(row);
        const itemName = row['最终商品名称'];
        orderGroups[optKey].itemsMap[itemName] = (orderGroups[optKey].itemsMap[itemName] || 0) + row['解析数量'];
    });
    
    const parsedRules = orderRuleConfig.map(r => ({
        country: safeStr(r.country).toLowerCase(),
        itemsMap: parseItemsToMap(r.items),
        channel: safeStr(r.channel).trim(),
        price: r.price
    }));
    
    let unmatchedOrders = [];
    let detailMoneySum = 0;
    
    for (let key in orderGroups) {
        const order = orderGroups[key];
        const ordCountry = order.country.toLowerCase();
        const ordChannel = order.channel;
        let matchFound = false;
        
        for (let rule of parsedRules) {
            if (rule.country && ordCountry !== rule.country) continue;
            if (rule.channel && rule.channel !== '*' && rule.channel !== '') {
                if (rule.channel.startsWith('!')) {
                    const excludeKeyword = rule.channel.substring(1).trim();
                    if (ordChannel.includes(excludeKeyword)) continue; 
                } else if (!ordChannel.includes(rule.channel)) {
                    continue; 
                }
            }
            if (!isItemMapMatch(order.itemsMap, rule.itemsMap)) continue;
            
            order.isMatched = true; order.matchedPrice = rule.price; matchFound = true; break; 
        }
        
        if (!matchFound) {
            let itemsArr = [];
            // 对商品名称进行升序排序 (例如: 3040 排在 4060 前面)
            let sortedKeys = Object.keys(order.itemsMap).sort();
            for (let itemName of sortedKeys) {
                itemsArr.push(\`\${itemName}*\${order.itemsMap[itemName]}\`);
            }
            // 按照需求模糊匹配渠道：只要包含"自提客选"的保留，其余统统标记为"!自提客选"
            let reportChannel = ordChannel.includes('自提客选') ? '自提客选' : '!自提客选';
            
            unmatchedOrders.push({
                transaction: key.split('|||')[0],
                customer: key.split('|||')[1],
                country: order.country,
                channel: reportChannel,
                items: itemsArr.join(', ')
            });
        }
        
        detailMoneySum += order.matchedPrice;
        let isFirstRow = true;
        order.rows.forEach(r => {
            if(isFirstRow) { r['本单最终匹配总价'] = order.matchedPrice; isFirstRow = false; } 
            else { r['本单最终匹配总价'] = 0; }
            r['整单状态'] = matchFound ? '成功匹配' : '查无此规则';
        });
    }
    
    const orderSummary = [];
    let orderQtySum = 0, orderMoneySum = 0;
    for (let k in orderGroups) {
        const order = orderGroups[k];
        const parts = k.split('|||'); 
        let tQty = 0; for(let item in order.itemsMap) { tQty += order.itemsMap[item]; }
        
        let itemsDesc = []; 
        let sortedKeys = Object.keys(order.itemsMap).sort();
        for(let item of sortedKeys) { itemsDesc.push(\`\${item}*\${order.itemsMap[item]}\`); }
        
        orderSummary.push({
            '交易编号': parts[0], '客户姓名': parts[1], '国家': order.country,
            '物流渠道': order.channel, '包含商品': itemsDesc.join(', '),
            '总商品数量': tQty, '订单最终总价': order.matchedPrice
        });
        orderQtySum += tQty; orderMoneySum += order.matchedPrice;
    }
    
    const skuSummary = [];
    const skuSumMap = {}; 
    data.forEach(row => {
        const skuKey = row['最终商品名称']; 
        if (!skuSumMap[skuKey]) skuSumMap[skuKey] = 0;
        skuSumMap[skuKey] += row['解析数量'];
    });
    for(let k in skuSumMap) { skuSummary.push({'订单商品名称': k, '总销量': skuSumMap[k]}); }
    skuSummary.sort((a,b) => b['总销量'] - a['总销量']);
    
    if (origRowCount !== data.length) { logMain(\`   ❌ 致命警报！数据行丢失！\`, 'log-error'); allPassed = false; }
    if (Math.abs(origQtySum - orderQtySum) > 0.01) { logMain(\`   ❌ 源头警报！商品总件数不匹配！\`, 'log-error'); allPassed = false; }
    
    let detailMoneyRealSum = data.reduce((sum, r) => sum + (r['本单最终匹配总价'] || 0), 0);
    if (Math.abs(detailMoneyRealSum - orderMoneySum) > 0.01) { logMain(\`   ❌ 汇总警报！内部金额求和对不上！\`, 'log-error'); allPassed = false; }
    
    if (unmatchedOrders.length > 0) {
        allPassed = false;
        logMain(\`   ❌ 规则匹配警报！发现 \${unmatchedOrders.length} 个订单未匹配到任何定价规则：\`, 'log-error');
        unmatchedOrders.slice(0, 10).forEach(uo => {
            logMain(\`      - 交易编号: \${uo.transaction} | 客户: \${uo.customer} | 国家: \${uo.country} | 包含: \${uo.items}\`, 'log-error');
        });
        if (unmatchedOrders.length > 10) logMain(\`      ... (等共计 \${unmatchedOrders.length} 条未匹配记录)\`, 'log-error');
    } else {
        logMain(\`   ✅ 源文件无缝对接，全部规则识别成功！流水: ¥\${orderMoneySum.toFixed(2)}\`, 'log-success');
    }
    logMain(\`--------------------------------------------------\`, 'log-info');

    data.forEach(r => { delete r['最终商品名称']; delete r['最终尺码']; delete r['解析数量']; });

    return {
        passed: allPassed,
        unmatchedData: unmatchedOrders, 
        sheets: {
            '添加价格后的明细': data,
            '按订单汇总': orderSummary,
            'SKU总销量统计': skuSummary
        }
    };
}

function exportExcelAndZip(results) {
    const zip = new JSZip();
    let isGlobalSuccess = true;
    let globalUnmatched = [];
    
    results.forEach(res => {
        if(!res.passed) isGlobalSuccess = false;
        if (res.unmatchedData && res.unmatchedData.length > 0) {
            globalUnmatched = globalUnmatched.concat(res.unmatchedData);
        }
        
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['添加价格后的明细']), '添加价格后的明细');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['按订单汇总']), '按订单汇总');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['SKU总销量统计']), 'SKU总销量统计');
        
        const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
        const extIndex = res.fileName.lastIndexOf('.');
        const baseName = extIndex > -1 ? res.fileName.substring(0, extIndex) : res.fileName;
        zip.file(\`\${baseName}已处理.xlsx\`, wbout);
    });
    
    if (globalUnmatched.length > 0) {
        // 利用对象键进行去重，避免生成几十行完全重复的待配置规则
        const uniqueRules = {};
        globalUnmatched.forEach(u => {
            const ruleKey = \`\${u.country}|||\${u.items}|||\${u.channel}\`;
            if (!uniqueRules[ruleKey]) {
                uniqueRules[ruleKey] = {
                    '国家': u.country,
                    '包含商品及数量': u.items,
                    '物流渠道': u.channel,
                    '整单总价': '' // 留空给客户填写
                };
            }
        });
        
        const exportFormat = Object.values(uniqueRules);
        
        const errorWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(errorWb, XLSX.utils.json_to_sheet(exportFormat), '未匹配订单列表');
        const errorOut = XLSX.write(errorWb, {bookType:'xlsx', type:'array'});
        zip.file(\`🚨未匹配异常订单排查清单.xlsx\`, errorOut);
        logMain(\`\\n[优化补丁] 已提取并自动去重 \${exportFormat.length} 种无规则的订单组合，保存到排查清单中！\`, 'log-warning');
    }
    
    if (isGlobalSuccess) {
        logMain(\`\\n🎉 恭喜！全链路对账通过，全部订单精准匹配到规则！\`, 'log-success');
    } else {
        logMain(\`\\n🚨 注意：压缩包内包含【🚨未匹配异常订单排查清单】，请直接复制并完善规则库后重试！\`, 'log-error');
    }
    
    return zip;
}
    `;
}

startBtn.addEventListener('click', async () => {
    startBtn.disabled = true;
    startBtn.innerText = '⚡ 后台处理中...';
    clearLog();
    log('⚡ 初始化多线程引擎，准备发送数据至后台...', 'log-info');
    
    try {
        // 读取所有未决文件为 ArrayBuffer (可转移给 Worker)
        const filesData = await Promise.all(pendingFiles.map(async (file) => {
            const buffer = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = e => resolve(e.target.result);
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
            return { name: file.name, buffer: buffer };
        }));
        
        // 动态生成内联 Web Worker（解决本地 file:// 协议下载的跨域读取限制）
        const workerBlob = new Blob([getWorkerCode()], { type: "application/javascript" });
        const workerUrl = URL.createObjectURL(workerBlob);
        const worker = new Worker(workerUrl);
        
        worker.onmessage = function(e) {
            const msg = e.data;
            if (msg.type === 'log') {
                log(msg.message, msg.logType);
            } else if (msg.type === 'done') {
                processedZips = msg.zipBlob;
                downloadBtn.style.display = 'inline-block';
                startBtn.disabled = false;
                startBtn.innerText = '🚀 重新开始自动化处理';
                URL.revokeObjectURL(workerUrl); // 释放内存
                worker.terminate();
            } else if (msg.type === 'error') {
                log(`❌ 后台线程崩溃: ${msg.error}`, 'log-error');
                startBtn.disabled = false;
                startBtn.innerText = '🚀 重新开始自动化处理';
                URL.revokeObjectURL(workerUrl);
                worker.terminate();
            }
        };
        
        worker.onerror = function(err) {
            log(`❌ Worker 加载失败/发生异常: ${err.message}`, 'log-error');
            startBtn.disabled = false;
            startBtn.innerText = '🚀 重新开始自动化处理';
            URL.revokeObjectURL(workerUrl);
            worker.terminate();
        };

        // 发送给后台处理
        worker.postMessage({
            orderRuleConfig: orderRuleConfig,
            filesData: filesData
        });
        
    } catch (e) {
        log(`❌ 主线程发送错误: ${e.message}`, 'log-error');
        console.error(e);
        startBtn.disabled = false;
        startBtn.innerText = '🚀 重新开始自动化处理';
    }
});

downloadBtn.addEventListener('click', () => {
    if (processedZips) {
        saveAs(processedZips, "优化版_订单对账结果与异常排查表.zip");
    }
});
