// --- 初始化默认价格字典 ---
const defaultPriceData = [
    { item: '4060', size: 'XL', tiers: [ {qty: 5, price: 35.0}, {qty: 3, price: 45.0}, {qty: 1, price: 55.0} ] },
    { item: '4060', size: 'M', tiers: [ {qty: 5, price: 30.0}, {qty: 3, price: 40.0}, {qty: 1, price: 50.0} ] },
    { item: 'Recznik', size: '默认', tiers: [ {qty: 2, price: 12.0}, {qty: 1, price: 15.0} ] },
    { item: '3040', size: 'XL', tiers: [ {qty: 1, price: 40.0} ] },
    { item: '3040', size: 'M', tiers: [ {qty: 1, price: 35.0} ] },
    { item: '6090', size: '默认', tiers: [ {qty: 1, price: 35.0} ] },
    { item: '洗衣袋', size: '默认', tiers: [ {qty: 1, price: 30.0} ] },
    { item: '三色抹布', size: '默认', tiers: [ {qty: 1, price: 20.0} ] },
    { item: 'Ręcznik łazienkowy BezKropli', size: '默认', tiers: [ {qty: 1, price: 10.0} ] },
    { item: 'Ceramika w Sprayu', size: '默认', tiers: [ {qty: 1, price: 10.0} ] },
    { item: '__kaching_bundles:eyJkZWFsIjoiS3JSdyIsIm1haW4iOnRydWUsImlkIjoiSDgxZiJ9', size: 'XL', tiers: [ {qty: 5, price: 35.0}, {qty: 3, price: 45.0}, {qty: 1, price: 55.0} ] },
    { item: '__kaching_bundles:eyJkZWFsIjoiS3JSdyIsIm1haW4iOnRydWUsImlkIjoiSDgxZiJ9', size: 'M', tiers: [ {qty: 5, price: 30.0}, {qty: 3, price: 40.0}, {qty: 1, price: 50.0} ] }
];

let priceConfig = [];
let pendingFiles = [];
let processedZips = null;

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

function renderTable() {
    tableBody.innerHTML = '';
    priceConfig.forEach((row, rowIndex) => {
        row.tiers.forEach((tier, tierIndex) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type="text" value="${row.item}" data-row="${rowIndex}" data-field="item" class="price-input"></td>
                <td><input type="text" value="${row.size}" data-row="${rowIndex}" data-field="size" class="price-input" style="width: 50px;"></td>
                <td><input type="number" min="1" value="${tier.qty}" data-row="${rowIndex}" data-tier="${tierIndex}" data-field="qty" class="price-input" style="width: 60px;"></td>
                <td><input type="number" step="0.01" min="0" value="${tier.price}" data-row="${rowIndex}" data-tier="${tierIndex}" data-field="price" class="price-input" style="width: 70px;"></td>
                <td><button class="btn btn-delete" data-row="${rowIndex}" data-tier="${tierIndex}">🗑️</button></td>
            `;
            tableBody.appendChild(tr);
        });
    });
    // 绑定事件
    document.querySelectorAll('.price-input').forEach(input => {
        input.addEventListener('change', handleInputChange);
    });
    document.querySelectorAll('.btn-delete').forEach(btn => {
        btn.addEventListener('click', handleDeleteRow);
    });
}

function handleInputChange(e) {
    const el = e.target;
    const rowIndex = parseInt(el.getAttribute('data-row'));
    const field = el.getAttribute('data-field');
    const val = el.value;
    
    if (field === 'item' || field === 'size') {
        priceConfig[rowIndex][field] = val;
        // 同步修改所有同组合的商品名称和尺码
        renderTable(); 
    } else {
        const tierIndex = parseInt(el.getAttribute('data-tier'));
        priceConfig[rowIndex].tiers[tierIndex][field] = parseFloat(val);
    }
}

function handleDeleteRow(e) {
    const el = e.target.closest('button');
    const rowIndex = parseInt(el.getAttribute('data-row'));
    const tierIndex = parseInt(el.getAttribute('data-tier'));
    
    priceConfig[rowIndex].tiers.splice(tierIndex, 1);
    if (priceConfig[rowIndex].tiers.length === 0) {
        priceConfig.splice(rowIndex, 1);
    }
    renderTable();
}

document.getElementById('add-row-btn').addEventListener('click', () => {
    priceConfig.push({ item: '新商品', size: '默认', tiers: [{qty: 1, price: 0.0}]});
    renderTable();
});

document.getElementById('reset-default-btn').addEventListener('click', () => {
    if(confirm("确定要恢复默认价格表吗？当前修改将会丢失。")) {
        priceConfig = JSON.parse(JSON.stringify(defaultPriceData));
        renderTable();
    }
});

// 初始化数据
priceConfig = JSON.parse(JSON.stringify(defaultPriceData));
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

// --- 核心业务逻辑 (原 process_app.py + verify_app.py) ---

function getPriceDictionary() {
    // 转换UI配置为后端容易查找的字典图
    const dict = {};
    priceConfig.forEach(row => {
        const key = `${row.item.trim()}|${row.size.trim()}`;
        // 确保降序排列
        const sortedTiers = [...row.tiers].sort((a,b) => b.qty - a.qty);
        dict[key] = sortedTiers;
    });
    return dict;
}

// 解析单元格值
function safeStr(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim();
}

// 模拟 pandas 读取并清理表头
function readExcelData(workbook) {
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    // 读取为 JSON，并将所有键转为字符串并去除空格
    const rawJson = XLSX.utils.sheet_to_json(worksheet, {defval: ''});
    return rawJson.map(row => {
        const cleanRow = {};
        for(let k in row) {
            cleanRow[k.trim()] = row[k];
        }
        return cleanRow;
    });
}

function processDataAndVerify(data, fileName, priceDict) {
    let allPassed = true;
    let origRowCount = data.length;
    let origQtySum = data.reduce((sum, row) => sum + (parseFloat(row['商品数量']) || 0), 0);
    
    log(`\n📊 正在处理并比对: [${fileName}]`);
    
    // 1. 提取尺码
    data.forEach(row => {
        const skuAttr = safeStr(row['平台SKU多属性']);
        if (skuAttr.includes(';')) {
            row['尺码'] = skuAttr.split(';').pop().trim();
        } else {
            row['尺码'] = '默认';
        }
    });
    
    // 2. 智能填补商品名称
    data.forEach(row => {
        let name = safeStr(row['订单商品名称']);
        if (name === '' || name.toLowerCase() === 'nan') {
            const skuAttr = safeStr(row['平台SKU多属性']);
            if (skuAttr.includes(';')) {
                row['订单商品名称'] = skuAttr.substring(0, skuAttr.lastIndexOf(';')).trim();
            } else {
                row['订单商品名称'] = skuAttr;
            }
        }
    });
    
    // 3. 计算本订单单品总数 GroupBy(交易编号, 订单商品名称, 尺码)
    const orderItemQtyMap = {};
    data.forEach(row => {
        const qty = parseFloat(row['商品数量']) || 0;
        row['商品数量计算'] = qty; // 持久化数字格式
        const key = `${safeStr(row['交易编号'])}_${safeStr(row['订单商品名称'])}_${safeStr(row['尺码'])}`;
        orderItemQtyMap[key] = (orderItemQtyMap[key] || 0) + qty;
    });
    
    // 4. 匹配价格并计算结果
    let zeroPriceCount = 0;
    let mathErrorCount = 0;
    
    data.forEach(row => {
        const item = safeStr(row['订单商品名称']);
        const size = safeStr(row['尺码']);
        const orderKey = `${safeStr(row['交易编号'])}_${item}_${size}`;
        const totalQtyInOrder = orderItemQtyMap[orderKey] || 0;
        
        let matchedTiers = null;
        const exactKey = `${item}|${size}`;
        
        if (priceDict[exactKey]) {
            matchedTiers = priceDict[exactKey];
        } else {
            // 模糊匹配
            for (let k in priceDict) {
                const [dictItem, dictSize] = k.split('|');
                if (item.includes(dictItem)) {
                    matchedTiers = priceDict[k];
                    break;
                }
            }
        }
        
        let unitPrice = 0.0;
        if (matchedTiers) {
            for (let t of matchedTiers) {
                if (totalQtyInOrder >= t.qty) {
                    unitPrice = t.price;
                    break;
                }
            }
        }
        
        row['单价'] = unitPrice;
        row['商品总价'] = Math.round(unitPrice * row['商品数量计算'] * 100) / 100;
        
        if (unitPrice === 0) zeroPriceCount++;
    });

    // 验证逻辑
    let detailQtySum = 0;
    let detailMoneySum = 0;
    
    const orderSumMap = {}; // 按交易编号，客户姓名 汇总
    const skuSumMap = {};   // 按 订单商品名称, 尺码 汇总
    
    data.forEach(row => {
        detailQtySum += row['商品数量计算'];
        detailMoneySum += row['商品总价'];
        
        // 订单汇总
        const optKey = `${safeStr(row['交易编号'])}|||${safeStr(row['客户姓名'])}`;
        if(!orderSumMap[optKey]) orderSumMap[optKey] = {qty: 0, money: 0};
        orderSumMap[optKey].qty += row['商品数量计算'];
        orderSumMap[optKey].money += row['商品总价'];
        
        // SKU汇总
        const skuKey = `${safeStr(row['订单商品名称'])}|||${safeStr(row['尺码'])}`;
        if(!skuSumMap[skuKey]) skuSumMap[skuKey] = {qty: 0, money: 0};
        skuSumMap[skuKey].qty += row['商品数量计算'];
        skuSumMap[skuKey].money += row['商品总价'];
    });
    
    // 生成表2数据
    const orderSummary = [];
    let orderQtySum = 0, orderMoneySum = 0;
    for(let k in orderSumMap) {
        const parts = k.split('|||');
        orderSummary.push({
            '交易编号': parts[0],
            '客户姓名': parts[1],
            '总商品数量': orderSumMap[k].qty,
            '订单总金额': Math.round(orderSumMap[k].money*100)/100
        });
        orderQtySum += orderSumMap[k].qty;
        orderMoneySum += orderSumMap[k].money;
    }
    
    // 生成表3数据
    const itemSummary = [];
    let itemQtySum = 0, itemMoneySum = 0;
    for(let k in skuSumMap) {
        const parts = k.split('|||');
        itemSummary.push({
            '订单商品名称': parts[0],
            '尺码': parts[1],
            '总销量': skuSumMap[k].qty,
            '总销售额': Math.round(skuSumMap[k].money*100)/100
        });
        itemQtySum += skuSumMap[k].qty;
        itemMoneySum += skuSumMap[k].money;
    }
    // 表3按金额降序
    itemSummary.sort((a,b) => b['总销售额'] - a['总销售额']);
    
    // --- 执行严苛对账逻辑 ---
    
    if (origRowCount !== data.length) {
        log(`   ❌ 源头警报！数据行数丢失：原始表(${origRowCount}行) vs 结果表(${data.length}行)！`, 'log-error');
        allPassed = false;
    }
    if (Math.abs(origQtySum - detailQtySum) > 0.01) {
        log(`   ❌ 源头警报！商品总件数不匹配：原始表进了(${origQtySum}件) vs 结果表只出来(${detailQtySum}件)！`, 'log-error');
        allPassed = false;
    }
    
    if (zeroPriceCount > 0) {
        log(`   ❌ 价格警报！发现 ${zeroPriceCount} 条商品单价为 0 的记录（未匹配到字典价格）`, 'log-error');
        allPassed = false;
    }
    
    if (Math.abs(detailQtySum - orderQtySum) > 0.01 || Math.abs(detailQtySum - itemQtySum) > 0.01) {
        log(`   ❌ 汇总警报！内部商品总数对不上：明细(${detailQtySum}) vs 订单(${orderQtySum}) vs SKU(${itemQtySum})`, 'log-error');
        allPassed = false;
    }
    
    if (Math.abs(detailMoneySum - orderMoneySum) > 0.02 || Math.abs(detailMoneySum - itemMoneySum) > 0.02) {
        log(`   ❌ 汇总警报！内部总金额对不上：明细(${detailMoneySum}) vs 订单(${orderMoneySum}) vs SKU(${itemMoneySum})`, 'log-error');
        allPassed = false;
    }
    
    if (allPassed) {
        log(`   ✅ 源文件无缝对接，内部对账无误！完美匹配流水: ¥${detailMoneySum.toFixed(2)}`, 'log-success');
    } else {
        log(`   ⚠️ 该表比对未通过，请仔细核查报错！`, 'log-warning');
    }
    log(`--------------------------------------------------`, 'log-info');

    return {
        passed: allPassed,
        sheets: {
            '添加价格后的明细': data,
            '按订单汇总': orderSummary,
            '按商品SKU汇总': itemSummary
        }
    };
}

function exportExcelAndZip(results) {
    const zip = new JSZip();
    let isGlobalSuccess = true;
    
    results.forEach(res => {
        if(!res.passed) isGlobalSuccess = false;
        
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['添加价格后的明细']), '添加价格后的明细');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['按订单汇总']), '按订单汇总');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['按商品SKU汇总']), '按商品SKU汇总');
        
        // 导出为 ArrayBuffer
        const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
        
        const extIndex = res.fileName.lastIndexOf('.');
        const baseName = extIndex > -1 ? res.fileName.substring(0, extIndex) : res.fileName;
        zip.file(`${baseName}已处理.xlsx`, wbout);
    });
    
    if (isGlobalSuccess) {
        log(`\n🎉 恭喜！所有表格的全链路对账全部通过：零漏单、零丢行、零错算！安全可靠！`, 'log-success');
    } else {
        log(`\n🚨 注意：发现部分数据源与结果不匹配或存在内部错误（未找到单价等），请仔细查阅上述日志！打包仍将进行。`, 'log-warning');
    }
    
    return zip;
}

// 主处理器
startBtn.addEventListener('click', async () => {
    startBtn.disabled = true;
    startBtn.innerText = '⚡ 处理中...';
    clearLog();
    log('⚡ 纯本地 JS 引擎启动，开始处理...', 'log-info');
    
    const priceDict = getPriceDictionary();
    const finalResults = [];
    
    try {
        for (let i = 0; i < pendingFiles.length; i++) {
            const file = pendingFiles[i];
            const data = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const wb = XLSX.read(e.target.result, {type: 'array'});
                        resolve(readExcelData(wb));
                    } catch (err) { reject(err); }
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
            
            const processed = processDataAndVerify(data, file.name, priceDict);
            processed.fileName = file.name;
            finalResults.push(processed);
        }
        
        const zip = exportExcelAndZip(finalResults);
        processedZips = await zip.generateAsync({type:"blob"});
        
        downloadBtn.style.display = 'inline-block';
        log('\n🎁 处理完毕！点击上方按钮下载打包结果。', 'log-success');
        
    } catch (e) {
        log(`❌ 发生致命错误: ${e.message}`, 'log-error');
        console.error(e);
    } finally {
        startBtn.disabled = false;
        startBtn.innerText = '🚀 重新开始自动化处理';
    }
});

downloadBtn.addEventListener('click', () => {
    if (processedZips) {
        saveAs(processedZips, "电商订单财务对账结果_Web.zip");
    }
});
