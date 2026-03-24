// --- 初始化默认规则表 ---
const defaultRuleData = [
    { country: 'poland', items: '4060*3', channel: '自提客选', price: 11.0 },
    { country: 'polan', items: '4060*3, 3040*2', channel: '!自提客选', price: 25.0 },
    { country: 'poland', items: '4060*5, 3040*1', channel: '', price: 30.0 }
];

let orderRuleConfig = [];
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
    orderRuleConfig.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" value="${row.country}" data-row="${rowIndex}" data-field="country" class="rule-input" style="width:100%"></td>
            <td><input type="text" value="${row.items}" data-row="${rowIndex}" data-field="items" class="rule-input" style="width:100%" placeholder="4060*3, 3040*2"></td>
            <td><input type="text" value="${row.channel}" data-row="${rowIndex}" data-field="channel" class="rule-input" style="width:100%" placeholder="!排除特定渠道 或 留空"></td>
            <td><input type="number" step="0.01" min="0" value="${row.price}" data-row="${rowIndex}" data-field="price" class="rule-input" style="width:80px;"></td>
            <td><button class="btn btn-delete" data-row="${rowIndex}">🗑️</button></td>
        `;
        tableBody.appendChild(tr);
    });
    
    document.querySelectorAll('.rule-input').forEach(input => {
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
    
    if (field === 'price') {
        orderRuleConfig[rowIndex][field] = parseFloat(val) || 0;
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
    if(confirm("确定要恢复默认规则表吗？当前修改将会丢失。")) {
        orderRuleConfig = JSON.parse(JSON.stringify(defaultRuleData));
        saveConfigToLocal();
        renderTable();
    }
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

// --- 核心业务逻辑 (基于订单整单匹配) ---

// 解析 UI 中的 "4060*3, 3040*2" 为 Map 结构
function parseItemsToMap(itemStr) {
    const map = {};
    if (!itemStr) return map;
    const parts = String(itemStr).split(/,|，/);
    parts.forEach(p => {
        p = p.trim();
        if (!p) return;
        
        // 匹配末尾的 *数量 或 x数量，防止商品名中自带 x 字母被错误截断
        const match = p.match(/(.+?)\s*(?:\*|x|X|×)\s*(\d+)$/i);
        if (match) {
            const name = match[1].trim();
            const qty = parseInt(match[2]);
            map[name] = (map[name] || 0) + qty;
        } else {
            // 如果只有商品名没有写数量，默认计为 1
            map[p] = (map[p] || 0) + 1;
        }
    });
    return map;
}

// 解析单元格值
function safeStr(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim();
}

function readExcelData(workbook) {
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawJson = XLSX.utils.sheet_to_json(worksheet, {defval: ''});
    return rawJson.map(row => {
        const cleanRow = {};
        for(let k in row) {
            cleanRow[k.trim()] = row[k];
        }
        return cleanRow;
    });
}

// Map 深度对比判断 (A 和 B 的 key 和 value 必须一模一样，或者模糊匹配？暂时要求准确匹配汇总出的商品)
function isItemMapMatch(orderMap, ruleMap) {
    const orderKeys = Object.keys(orderMap);
    const ruleKeys = Object.keys(ruleMap);
    
    // 如果订单里面有别的商品，或者缺少商品，直接返回false (严格匹配)
    if (orderKeys.length !== ruleKeys.length) return false;
    
    for (let k of ruleKeys) {
        // 先尝试精确匹配
        if (orderMap[k] !== ruleMap[k]) {
            // 如果精确匹配失败，由于原系统允许名字带有附加字符，我们可以做一个非常鲁棒的包含判定
            let foundMatch = false;
            for(let ok of orderKeys) {
                if(ok.includes(k) && orderMap[ok] === ruleMap[k]) {
                    foundMatch = true;
                    break;
                }
            }
            if (!foundMatch) return false;
        }
    }
    return true;
}

// 获取单行的核心字段处理
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
    
    // 尺码
    const skuAttr2 = safeStr(row['平台SKU多属性']);
    if (skuAttr2.includes(';')) {
        row['最终尺码'] = skuAttr2.split(';').pop().trim();
    } else {
        row['最终尺码'] = '默认';
    }
    
    row['解析数量'] = parseFloat(row['商品数量']) || 0;
}

function processDataAndVerify(data, fileName) {
    let allPassed = true;
    let origRowCount = data.length;
    let origQtySum = 0;
    
    log(`\n📊 正在处理并比对: [${fileName}]`);
    
    // --- 1. 按订单聚合并清洗数据 ---
    const orderGroups = {}; // { '交易编号_客户姓名' : {rows: [], country: '', channel: '', itemsMap: {}} }
    
    data.forEach(row => {
        preprocessRow(row);
        origQtySum += row['解析数量'];
        
        const optKey = `${safeStr(row['交易编号'])}|||${safeStr(row['客户姓名'])}`;
        if (!orderGroups[optKey]) {
            orderGroups[optKey] = {
                rows: [],
                country: safeStr(row['国家']),
                channel: safeStr(row['物流渠道']),
                itemsMap: {},
                matchedPrice: 0,
                isMatched: false
            };
        }
        
        orderGroups[optKey].rows.push(row);
        
        const itemName = row['最终商品名称'];
        orderGroups[optKey].itemsMap[itemName] = (orderGroups[optKey].itemsMap[itemName] || 0) + row['解析数量'];
    });
    
    // --- 2. 基于配置表进行订单级规则匹配 ---
    // 预解析规则
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
            // A. 国家匹配
            if (rule.country && ordCountry !== rule.country) continue;
            
            // B. 物流渠道匹配
            if (rule.channel && rule.channel !== '*' && rule.channel !== '') {
                if (rule.channel.startsWith('!')) {
                    const excluded = rule.channel.substring(1).trim();
                    if (ordChannel === excluded) continue; // 属于排除渠道，本规则不通过
                } else if (rule.channel !== ordChannel) {
                    continue; // 指定渠道不相等
                }
            }
            
            // C. 商品组合精确匹配
            if (!isItemMapMatch(order.itemsMap, rule.itemsMap)) {
                continue;
            }
            
            // 全部条件达成！
            order.isMatched = true;
            order.matchedPrice = rule.price;
            matchFound = true;
            break; 
        }
        
        if (!matchFound) {
            unmatchedOrders.push({
                key: key,
                country: order.country,
                items: JSON.stringify(order.itemsMap)
            });
        }
        
        detailMoneySum += order.matchedPrice;
        
        // --- 3. 将最终结果下卷到 Excel 的每一行 (展示用) ---
        let isFirstRow = true;
        order.rows.forEach(r => {
            // 为了保证明细表的金额汇总不出错，总额只挂在订单的第一行，其余行记录为0，或者可以展示一个总括说明。
            // 我们选择加一个专列 [本单核对手续费]
            if(isFirstRow) {
                r['本单最终匹配总价'] = order.matchedPrice;
                isFirstRow = false;
            } else {
                r['本单最终匹配总价'] = 0; // 其余行为0，避免表格直接求和时重复累计
            }
            r['整单状态'] = matchFound ? '成功匹配' : '查无此规则';
        });
    }
    
    // --- 4. 生成三个维度的工作表 ---
    // 表1: 明细
    // 已在上面赋予 data
    
    // 表2: 按订单汇总
    const orderSummary = [];
    let orderQtySum = 0, orderMoneySum = 0;
    
    for (let k in orderGroups) {
        const order = orderGroups[k];
        const parts = k.split('|||'); // 交易编号 ||| 客户姓名
        
        // 算出这个订单的总商品数量
        let tQty = 0;
        for(let item in order.itemsMap) { tQty += order.itemsMap[item]; }
        
        // 把商品 map 转化为直观的文字描述
        let itemsDesc = [];
        for(let item in order.itemsMap) { itemsDesc.push(`${item}*${order.itemsMap[item]}`); }
        
        orderSummary.push({
            '交易编号': parts[0],
            '客户姓名': parts[1],
            '国家': order.country,
            '物流渠道': order.channel,
            '包含商品': itemsDesc.join(', '),
            '总商品数量': tQty,
            '订单最终总价': order.matchedPrice
        });
        orderQtySum += tQty;
        orderMoneySum += order.matchedPrice;
    }
    
    // 表3: SKU销量汇总 (由于现在是整单计价，无法准确拆分单品销售额，只能汇总销量)
    const skuSummary = [];
    const skuSumMap = {}; 
    data.forEach(row => {
        const skuKey = `${row['最终商品名称']}|||${row['最终尺码']}`;
        if (!skuSumMap[skuKey]) skuSumMap[skuKey] = 0;
        skuSumMap[skuKey] += row['解析数量'];
    });
    
    let itemQtySum = 0;
    for(let k in skuSumMap) {
        const parts = k.split('|||');
        skuSummary.push({
            '订单商品名称': parts[0],
            '尺码': parts[1],
            '总销量': skuSumMap[k]
        });
        itemQtySum += skuSumMap[k];
    }
    skuSummary.sort((a,b) => b['总销量'] - a['总销量']);
    
    
    // --- 5. 严苛对账执行 ---
    if (origRowCount !== data.length) {
        log(`   ❌ 致命警报！数据行丢失：原始(${origRowCount}) vs 处理(${data.length})！`, 'log-error');
        allPassed = false;
    }
    if (Math.abs(origQtySum - orderQtySum) > 0.01) {
        log(`   ❌ 源头警报！商品总件数不匹配：原始表进了(${origQtySum}件) vs 订单结果只出来(${orderQtySum}件)！`, 'log-error');
        allPassed = false;
    }
    
    // 对比表1的金额累计(只加第一行)和表2的金额累计是否相等
    let detailMoneyRealSum = data.reduce((sum, r) => sum + (r['本单最终匹配总价'] || 0), 0);
    if (Math.abs(detailMoneyRealSum - orderMoneySum) > 0.01) {
        log(`   ❌ 汇总警报！内部金额求和对不上：明细表汇总(${detailMoneyRealSum}) vs 订单表汇总(${orderMoneySum})`, 'log-error');
        allPassed = false;
    }
    
    // 报告未匹配订单详情
    if (unmatchedOrders.length > 0) {
        allPassed = false;
        log(`   ❌ 规则匹配警报！发现 ${unmatchedOrders.length} 个订单未匹配到任何定价规则：`, 'log-error');
        // 最多展示 10 条，避免刷屏
        unmatchedOrders.slice(0, 10).forEach(uo => {
            const parts = uo.key.split('|||');
            log(`      - 交易编号: ${parts[0]} | 客户: ${parts[1]} | 国家: ${uo.country} | 包含: ${uo.items}`, 'log-error');
        });
        if (unmatchedOrders.length > 10) {
            log(`      ... (等共计 ${unmatchedOrders.length} 条未匹配记录)`, 'log-error');
        }
    }
    
    if (allPassed) {
        log(`   ✅ 源文件无缝对接，全部规则识别成功！本表完美结账流水: ¥${orderMoneySum.toFixed(2)}`, 'log-success');
    } else {
        log(`   ⚠️ 该表对账未通过，请检查上方【规则匹配警报】提示的客户订单配置！`, 'log-warning');
    }
    log(`--------------------------------------------------`, 'log-info');

    // 格式化输出的表头和多余列清理
    data.forEach(r => {
        delete r['最终商品名称'];
        delete r['最终尺码'];
        delete r['解析数量'];
    });

    return {
        passed: allPassed,
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
    
    results.forEach(res => {
        if(!res.passed) isGlobalSuccess = false;
        
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['添加价格后的明细']), '添加价格后的明细');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['按订单汇总']), '按订单汇总');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res.sheets['SKU总销量统计']), 'SKU总销量统计');
        
        const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
        
        const extIndex = res.fileName.lastIndexOf('.');
        const baseName = extIndex > -1 ? res.fileName.substring(0, extIndex) : res.fileName;
        zip.file(`${baseName}已处理.xlsx`, wbout);
    });
    
    if (isGlobalSuccess) {
        log(`\n🎉 恭喜！所有表格全链路对账通过，全部订单精准匹配到规则！`, 'log-success');
    } else {
        log(`\n🚨 注意：存在未匹配到规则价格的“漏网”订单，请在日志中核对这些【交易编号】！打包仍将进行。`, 'log-warning');
    }
    
    return zip;
}

// 主处理器
startBtn.addEventListener('click', async () => {
    startBtn.disabled = true;
    startBtn.innerText = '⚡ 处理中...';
    clearLog();
    log('⚡ 整单匹配智能引擎启动中...', 'log-info');
    
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
            
            const processed = processDataAndVerify(data, file.name);
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
        saveAs(processedZips, "规则匹配版_订单对账结果.zip");
    }
});
