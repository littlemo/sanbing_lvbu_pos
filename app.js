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
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);

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
        if (row['角色'] && row['角色'].toString().trim()) {
            const player = {
                name: String(row['角色']).trim(),
                stats: parseFloat(row['四维和'] || 0),
                power: parseFloat(row['战力'] || 0)
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

    const positions = calculatePositions(lubuX, lubuY, ringCount);

    const grid = document.getElementById('grid');
    grid.innerHTML = '';

    gridData = [];

    let maxX = 0, maxY = 0;
    positions.forEach(pos => {
        maxX = Math.max(maxX, pos.x);
        maxY = Math.max(maxY, pos.y);
    });

    maxX = Math.max(maxX, lubuX + 1);
    maxY = Math.max(maxY, lubuY + 1);

    const maxPlayers = 100;
    const maxGridSize = 20;
    if (maxX > maxGridSize) maxX = maxGridSize;
    if (maxY > maxGridSize) maxY = maxGridSize;

    grid.style.setProperty('--cols', maxX + 1);
    grid.style.setProperty('--rows', maxY + 1);

    for (let y = 0; y <= maxY; y++) {
        for (let x = 0; x <= maxX; x++) {
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.dataset.x = x;
            cell.dataset.y = y;

            if (x >= lubuX && x <= lubuX + 1 && y >= lubuY && y <= lubuY + 1) {
                cell.className += ' lubu';
                if (x === lubuX + 1 && y === lubuY + 1) {
                    cell.innerHTML = '吕布<br>校场';
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
                            power: player.power,
                            x: x,
                            y: y,
                            ring: playerPos.ring,
                            positionType: playerPos.positionType
                        });
                    }
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

    for (let ring = 1; ring <= ringCount; ring++) {
        const edgePositions = getEdgePositions(lubuX, lubuY, ring);
        const cornerPositions = getCornerPositions(lubuX, lubuY, ring);

        edgePositions.forEach(pos => {
            pos.index = positions.length;
            positions.push(pos);
        });
        cornerPositions.forEach(pos => {
            pos.index = positions.length;
            positions.push(pos);
        });
    }

    return positions;
}

function getEdgePositions(lubuX, lubuY, ring) {
    const positions = [];

    const baseX = lubuX + 1;
    const baseY = lubuY + 1;

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX + i,
            y: baseY,
            ring: ring,
            positionType: '边'
        });
    }

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX - i,
            y: baseY,
            ring: ring,
            positionType: '边'
        });
    }

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX,
            y: baseY + i,
            ring: ring,
            positionType: '边'
        });
    }

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX,
            y: baseY - i,
            ring: ring,
            positionType: '边'
        });
    }

    return positions;
}

function getCornerPositions(lubuX, lubuY, ring) {
    const positions = [];

    const baseX = lubuX + 1;
    const baseY = lubuY + 1;

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX + i,
            y: baseY + (ring - i + 1),
            ring: ring,
            positionType: '角'
        });
    }

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX - i,
            y: baseY + (ring - i + 1),
            ring: ring,
            positionType: '角'
        });
    }

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX + i,
            y: baseY - (ring - i + 1),
            ring: ring,
            positionType: '角'
        });
    }

    for (let i = 1; i <= ring; i++) {
        positions.push({
            x: baseX - i,
            y: baseY - (ring - i + 1),
            ring: ring,
            positionType: '角'
        });
    }

    return positions;
}

function updateResultsTable() {
    const tbody = document.getElementById('resultsTableBody');
    tbody.innerHTML = '';

    gridData.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.rank}</td>
            <td>${row.name}</td>
            <td>${row.stats.toFixed(0)}</td>
            <td>${row.power.toFixed(0)}</td>
            <td>${row.x}</td>
            <td>${row.y}</td>
            <td>${row.ring}</td>
            <td>${row.positionType}</td>
        `;
        tbody.appendChild(tr);
    });
}

function updateStats() {
    const totalPlayers = playerData.length;
    const ringCount = parseInt(document.getElementById('ringCount').value);
    const totalPositions = ringCount * (ringCount + 1) * 5;
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
        '战力': row.power,
        'X坐标': row.x,
        'Y坐标': row.y,
        '环数': row.ring,
        '位置类型': row.positionType
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
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);

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
                showMessage('Excel 文件读取成功！', 'success');
            })
            .catch(error => {
                showMessage('读取文件失败：' + error.message, 'error');
            });
    }
});

document.addEventListener('DOMContentLoaded', function() {
    updateStats();
});
