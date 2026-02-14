let playerData = [];
let gridData = [];

function processData() {
    const fileInput = document.getElementById('excelFile');
    if (!fileInput.files.length) {
        showMessage('请先上传 Excel 文件', 'error');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            // 实际表头在第 2 行，数据从第 3 行开始
            const rawData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            const headerRow = rawData[1];
            const dataRows = rawData.slice(2);

            const jsonData = dataRows.map(row => {
                const obj = {};
                headerRow.forEach((header, index) => {
                    if (index < row.length) {
                        obj[header] = row[index];
                    }
                });
                return obj;
            });

            processPlayerData(jsonData);
            generateGrid();
            showMessage('排位生成成功！', 'success');
        } catch (error) {
            console.error(error);
            showMessage('解析 Excel 文件失败：' + error.message, 'error');
        }
    };

    reader.onerror = function() {
        showMessage('读取文件失败', 'error');
    };

    reader.readAsArrayBuffer(file);
}

function processPlayerData(jsonData) {
    playerData = [];

    jsonData.forEach(row => {
        // 检查是否有游戏ID字段
        if (row['游戏ID'] && String(row['游戏ID']).trim()) {
            const player = {
                name: String(row['游戏ID']).trim(),
                stats: parseFloat(row['四维和'] || 0),
                defense: parseFloat(row['步维(坦度)'] || 0),
                attack: parseFloat(row['弓维(输出)'] || 0)
            };

            if (!isNaN(player.stats)) {
                playerData.push(player);
            }
        }
    });

    playerData.sort((a, b) => b.stats - a.stats);

    updateStats();
}

function generateGrid() {
    const lubuX = parseInt(document.getElementById('lubuX').value);
    const lubuY = parseInt(document.getElementById('lubuY').value);
    const ringCount = parseInt(document.getElementById('ringCount').value);

    const gridContainer = document.getElementById('grid-container');
    const heights = {
        1: 600,
        2: 800,
        3: 1000,
        4: 1200
    };
    gridContainer.style.minHeight = `${heights[ringCount]}px`;

    const positions = calculatePositions(lubuX, lubuY, ringCount);

    const grid = document.getElementById('grid');
    grid.innerHTML = '';

    gridData = [];

    let minX = Infinity, maxX = -Infinity;
    let minY = Infinity, maxY = -Infinity;

    positions.forEach(pos => {
        minX = Math.min(minX, pos.x);
        maxX = Math.max(maxX, pos.x);
        minY = Math.min(minY, pos.y);
        maxY = Math.max(maxY, pos.y);
    });

    // 布局坐标可以为负，不需要限制最小值
    minX = Math.min(minX, lubuX);
    maxX = Math.max(maxX, lubuX + 1);
    minY = Math.min(minY, lubuY);
    maxY = Math.max(maxY, lubuY + 1);

    const cols = maxX - minX + 1;
    const rows = maxY - minY + 1;

    grid.style.setProperty('--cols', cols);
    grid.style.setProperty('--rows', rows);

    // 调试信息：显示网格生成参数
    const DEBUG = true;
    if (DEBUG) {
        console.log("=== 网格生成 ===");
        console.log(`minY: ${minY}, maxY: ${maxY}`);
        console.log(`minX: ${minX}, maxX: ${maxX}`);
    }

    // 由于网格使用 transform: rotate(-45deg) 旋转，需要反向遍历 Y 轴
    // 以确保视觉上 Y 坐标值由下至上增大，符合常规坐标系习惯
    for (let y = maxY; y >= minY; y--) {
        if (DEBUG) console.log(`\n--- Y坐标: ${y} ---`);
        for (let x = minX; x <= maxX; x++) {
            // 只有当显示坐标（游戏中的坐标）为负数时才不绘制单元格
            if (x < 0 || y < 0) {
                continue;
            }

            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.dataset.x = x;
            cell.dataset.y = y;

            if (x >= lubuX && x <= lubuX + 1 && y >= lubuY && y <= lubuY + 1) {
                cell.className += ' lubu';
                if (x === lubuX && y === lubuY) {
                    cell.innerHTML = `吕布<br>校场<br><small>(${x},${y})</small>`;
                } else {
                    cell.innerHTML = `<small>(${x},${y})</small>`;
                }
            } else {
                const playerPos = positions.find(pos => pos.x === x && pos.y === y);
                if (playerPos) {
                    const player = playerData[playerPos.index];
                    if (player) {
                        cell.className += ` player ring-${playerPos.ring}`;
                        cell.innerHTML = `
                            <div class="rank">${playerPos.index + 1}</div>
                            <div class="name">${player.name}</div>
                            <div class="coords">${x},${y}</div>
                        `;

                        gridData.push({
                            rank: playerPos.index + 1,
                            name: player.name,
                            stats: player.stats,
                            defense: player.defense,
                            attack: player.attack,
                            x: x,
                            y: y,
                            ring: playerPos.ring,
                            positionType: playerPos.positionType,
                            distance: playerPos.distance.toFixed(2) // 保留两位小数
                        });
                    } else {
                        // 当找到位置但没有玩家数据时，显示坐标
                        cell.innerHTML = `<small>(${x},${y})</small>`;
                    }
                } else {
                    // 当找不到位置时，显示坐标
                    cell.innerHTML = `<small>(${x},${y})</small>`;
                }
            }

            grid.appendChild(cell);
        }
    }

    updateResultsTable();
    updateStats();
}

function calculatePositions(lubuX, lubuY, ringCount) {
    const positions = [];
    const baseX = lubuX;
    const baseY = lubuY;

    // 计算吕布校场4个cell的中心点 (lubuX + 0.5, lubuY + 0.5)
    const centerX = lubuX + 0.5;
    const centerY = lubuY + 0.5;

    // 直接实现用户公式：环数n，总位置数=(2n+2)^2 - 4 (吕布校场)
    // 为了确保坐标不小于0，我们需要调整起始坐标
    for (let ring = 1; ring <= ringCount; ring++) {
        const width = 2 * ring + 2;
        const height = 2 * ring + 2;

        // 确保起始坐标不小于0
        const startX = Math.max(0, lubuX - ring);
        const startY = Math.max(0, lubuY - ring);

        // 调整结束坐标，确保总宽度和高度
        const endX = startX + width;
        const endY = startY + height;

        for (let x = startX; x < endX; x++) {
            for (let y = startY; y < endY; y++) {
                if (isLubuField(x, y, lubuX, lubuY)) continue;

                // 检查位置是否已经存在
                const exists = positions.some(pos => pos.x === x && pos.y === y);
                if (!exists) {
                    // 计算当前位置到吕布校场中心点的距离
                    const distance = Math.sqrt(Math.pow(x - centerX, 2) + Math.pow(y - centerY, 2));

                    positions.push({
                        x,
                        y,
                        ring,
                        positionType: getPositionType(x, y, lubuX, lubuY, ring),
                        index: positions.length,
                        distance: distance
                    });
                }
            }
        }
    }

    // 按照距离由近到远排序
    positions.sort((a, b) => a.distance - b.distance);

    // 更新每个位置的index，确保排序后index是连续的
    positions.forEach((pos, index) => {
        pos.index = index;
    });

    return positions;
}

function isLubuField(x, y, baseX, baseY) {
    return x >= baseX && x <= baseX + 1 && y >= baseY && y <= baseY + 1;
}

function isOnRing(x, y, baseX, baseY, ring) {
    const distLeft = x - baseX;
    const distRight = (baseX + 1) - x;
    const distTop = y - baseY;
    const distBottom = (baseY + 1) - y;

    const minDist = Math.min(Math.abs(distLeft), Math.abs(distRight),
                           Math.abs(distTop), Math.abs(distBottom));

    return minDist === ring;
}

function getPositionType(x, y, baseX, baseY, ring) {
    const distLeft = Math.abs(x - baseX);
    const distRight = Math.abs(x - (baseX + 1));
    const distTop = Math.abs(y - baseY);
    const distBottom = Math.abs(y - (baseY + 1));

    const distX = Math.min(distLeft, distRight);
    const distY = Math.min(distTop, distBottom);

    if (distX === ring && distY === ring) {
        return '角';
    } else {
        return '边';
    }
}

function updateResultsTable() {
    const tbody = document.getElementById('resultsTableBody');
    tbody.innerHTML = '';

    // 按照排名升序排序
    const sortedData = [...gridData].sort((a, b) => a.rank - b.rank);

    sortedData.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.rank}</td>
            <td>${row.name}</td>
            <td>${row.stats.toFixed(0)}</td>
            <td>${row.defense.toFixed(0)}</td>
            <td>${row.attack.toFixed(0)}</td>
            <td>${row.x}</td>
            <td>${row.y}</td>
            <td>${row.ring}</td>
            <td>${row.positionType}</td>
            <td>${row.distance}</td>
        `;
        tbody.appendChild(tr);
    });
}

function updateStats() {
    const totalPlayers = playerData.length;
    const ringCount = parseInt(document.getElementById('ringCount').value);
    const lubuX = parseInt(document.getElementById('lubuX').value);
    const lubuY = parseInt(document.getElementById('lubuY').value);

    // 直接使用用户公式计算总位置数：(2n+2)^2 -4
    const totalPositions = calculatePositions(lubuX, lubuY, ringCount).length;

    const filledPositions = Math.min(totalPlayers, totalPositions);
    const emptyPositions = Math.max(0, totalPositions - totalPlayers);

    document.getElementById('totalPlayers').textContent = totalPlayers;
    document.getElementById('totalPositions').textContent = totalPositions;
    document.getElementById('filledPositions').textContent = filledPositions;
    document.getElementById('emptyPositions').textContent = emptyPositions;
}

function downloadResults() {
    if (gridData.length === 0) {
        showMessage('没有可导出的数据', 'error');
        return;
    }

    const data = gridData.map(row => ({
        '排名': row.rank,
        '玩家': row.name,
        '四维和': row.stats,
        '坦度': row.defense,
        '输出': row.attack,
        'X坐标': row.x,
        'Y坐标': row.y,
        '环数': row.ring,
        '位置类型': row.positionType,
        '距离': row.distance
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '排位结果');

    const fileName = `吕布校场排位结果_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
}

function showMessage(text, type = 'success') {
    const messageDiv = document.getElementById('message');
    messageDiv.className = type;
    messageDiv.textContent = text;
    messageDiv.style.display = 'block';

    if (type === 'success') {
        setTimeout(() => {
            messageDiv.style.display = 'none';
        }, 3000);
    }
}

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                // 实际表头在第2行（索引为1），数据从第3行开始
                // 使用 header: 1 来获取原始数组格式，然后重新格式化
                const rawData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

                // 跳过前2行（第1行是填写说明，第2行是表头）
                const headerRow = rawData[1];
                const dataRows = rawData.slice(2);

                // 转换为对象格式
                const jsonData = dataRows.map(row => {
                    const obj = {};
                    headerRow.forEach((header, index) => {
                        if (index < row.length) {
                            obj[header] = row[index];
                        }
                    });
                    return obj;
                });

                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = function() {
            reject(new Error('读取文件失败'));
        };

        reader.readAsArrayBuffer(file);
    });
}

document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        readExcelFile(file)
            .then(jsonData => {
                processPlayerData(jsonData);
                generateGrid(); // 自动触发排位逻辑
                showMessage('Excel 文件读取成功！', 'success');
            })
            .catch(error => {
                showMessage('读取文件失败：' + error.message, 'error');
            });
    }
});

document.addEventListener('DOMContentLoaded', function() {
    updateStats();
    generateGrid();

    // 为可选参数添加事件监听器，当参数变化时自动触发排位更新
    const lubuXInput = document.getElementById('lubuX');
    const lubuYInput = document.getElementById('lubuY');
    const ringCountInput = document.getElementById('ringCount');

    // 监听输入框变化事件
    lubuXInput.addEventListener('change', processData);
    lubuYInput.addEventListener('change', processData);
    ringCountInput.addEventListener('change', processData);

    // 监听输入框输入事件（实时更新）
    lubuXInput.addEventListener('input', processData);
    lubuYInput.addEventListener('input', processData);
    ringCountInput.addEventListener('input', processData);

    // 检查对角线上的单元格内容的脚本
    setTimeout(() => {
        console.log('=== 页面加载完成 ===');

        // 获取所有单元格
        const cells = document.querySelectorAll('.cell');
        console.log(`总单元格数: ${cells.length}`);

        // 找到对角线上的单元格
        const diagonalCells = [];
        cells.forEach(cell => {
            const x = parseInt(cell.dataset.x);
            const y = parseInt(cell.dataset.y);

            // 对角线单元格：x == y 或 x == 10 - y
            if (x === y || x === 10 - y) {
                diagonalCells.push(cell);
            }
        });

        console.log(`对角线上的单元格数: ${diagonalCells.length}`);

        // 打印对角线上的单元格信息
        diagonalCells.forEach(cell => {
            const x = cell.dataset.x;
            const y = cell.dataset.y;
            console.log(`\n--- 单元格(${x},${y}) ---`);
            console.log('innerHTML:', cell.innerHTML);
            console.log('textContent:', cell.textContent);
            console.log('className:', cell.className);

            // 检查是否包含坐标值
            const hasCoords = cell.innerHTML.includes(x + ',' + y);
            console.log('包含坐标值:', hasCoords);

            // 检查是否有 coords 元素
            const coordsElement = cell.querySelector('.coords');
            if (coordsElement) {
                console.log('coords元素内容:', coordsElement.textContent);
            } else {
                console.log('未找到 coords 元素');
            }
        });

        // 打印玩家单元格信息
        const playerCells = document.querySelectorAll('.player');
        console.log(`\n=== 玩家单元格(${playerCells.length}个) ===`);
        playerCells.forEach(cell => {
            const x = cell.dataset.x;
            const y = cell.dataset.y;
            console.log(`\n玩家单元格(${x},${y}):`);
            console.log('innerHTML:', cell.innerHTML);
            console.log('textContent:', cell.textContent);
            console.log('className:', cell.className);
        });
    }, 1000);
});
