// 官渡之战安排工具主逻辑

let playerData = [];
let mapData = [];

// 建筑数据
const buildings = [
    { id: 'wuchao', name: '乌巢', position: { x: 0, y: 0 }, openTime: '9:15', pointsPerMinute: '海量积分', effect: '每7分钟刷辎重车（共4波）', priority: 1, type: 'special', phase: 2 },
    { id: 'bingqifang', name: '兵器坊', position: { x: 0, y: 1 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '占领时间+30秒', priority: 2, type: 'production', phase: 1 },
    { id: 'dalangcang-left', name: '大粮仓(左)', position: { x: -1, y: 2 }, openTime: '9:25', pointsPerMinute: '8400/分钟', effect: '后期核心积分点', priority: 3, type: 'production', phase: 3 },
    { id: 'dalangcang-right', name: '大粮仓(右)', position: { x: 1, y: 2 }, openTime: '9:25', pointsPerMinute: '8400/分钟', effect: '后期核心积分点', priority: 3, type: 'production', phase: 3 },
    { id: 'guandu', name: '官渡', position: { x: 0, y: 2 }, openTime: '9:25', pointsPerMinute: '全场最高', effect: '胜负核心', priority: 1, type: 'special', phase: 3 },
    { id: 'piliuche', name: '霹雳车', position: { x: 0, y: 3 }, openTime: '9:25', pointsPerMinute: '8400/分钟', effect: '每30秒打官渡守军10%兵力', priority: 2, type: 'defense', phase: 3 },
    { id: 'gongjiangfang', name: '工匠坊', position: { x: 0, y: 4 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '部队破坏+生命值+20%', priority: 2, type: 'defense', phase: 1 },
    { id: 'xiaoliangcang-left-1', name: '小粮仓(左1)', position: { x: -2, y: 1 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '开局稳占积分', priority: 4, type: 'production', phase: 1 },
    { id: 'xiaoliangcang-left-2', name: '小粮仓(左2)', position: { x: -2, y: 2 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '开局稳占积分', priority: 4, type: 'production', phase: 1 },
    { id: 'xiaoliangcang-left-3', name: '小粮仓(左3)', position: { x: -2, y: 3 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '开局稳占积分', priority: 4, type: 'production', phase: 1 },
    { id: 'xiaoliangcang-right-1', name: '小粮仓(右1)', position: { x: 2, y: 1 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '开局稳占积分', priority: 4, type: 'production', phase: 1 },
    { id: 'xiaoliangcang-right-2', name: '小粮仓(右2)', position: { x: 2, y: 2 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '开局稳占积分', priority: 4, type: 'production', phase: 1 },
    { id: 'xiaoliangcang-right-3', name: '小粮仓(右3)', position: { x: 2, y: 3 }, openTime: '9:00', pointsPerMinute: '4200/分钟', effect: '开局稳占积分', priority: 4, type: 'production', phase: 1 }
];

// 初始化函数
function init() {
    updateStats();
    renderMap();

    // 为控件添加事件监听器
    const excelFileInput = document.getElementById('excelFile');
    const phaseSelect = document.getElementById('phase');
    const sortBySelect = document.getElementById('sortBy');
    const legionSelect = document.getElementById('legion');

    excelFileInput.addEventListener('change', handleFileUpload);
    phaseSelect.addEventListener('change', processData);
    sortBySelect.addEventListener('change', processData);
    legionSelect.addEventListener('change', processData);
}

// 文件上传处理
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (file) {
        readExcelFile(file)
            .then(jsonData => {
                processPlayerData(jsonData);
                processData();
                showMessage('Excel 文件读取成功！', 'success');
            })
            .catch(error => {
                showMessage('读取文件失败：' + error.message, 'error');
            });
    }
}

// 读取Excel文件
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
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

// 处理玩家数据
function processPlayerData(jsonData) {
    playerData = [];

    const ignoredPlayers = JSON.parse(localStorage.getItem('ignoredPlayersGuandu') || '{}');

    jsonData.forEach(row => {
        if (row['游戏ID'] && String(row['游戏ID']).trim()) {
            const playerName = String(row['游戏ID']).trim();
            const player = {
                name: playerName,
                stats: parseFloat(row['四维和'] || 0),
                sixStats: parseFloat(row['六维和'] || 0),
                cavalryArchery: parseFloat(row['骑弓维(输出)'] || 0),
                attack: parseFloat(row['弓维(输出)'] || 0),
                defense: parseFloat(row['步维(坦度)'] || 0),
                role: '中战', // 默认分工
                legion: 1, // 默认军团
                ignore: ignoredPlayers[playerName] || false
            };

            if (!isNaN(player.stats)) {
                playerData.push(player);
            }
        }
    });
}

// 智能分配算法
function allocatePlayers(phase = 1, sortBy = 'stats', legion = 1) {
    console.log('global.playerData:', global.playerData);
    // 如果没有传递参数，则从DOM中获取
    if (phase === null) {
        phase = parseInt(document.getElementById('phase').value);
        sortBy = document.getElementById('sortBy').value;
        legion = parseInt(document.getElementById('legion').value);
    }

    // 筛选当前军团的玩家
    let filteredPlayers = global.playerData.filter(player => player.legion === legion && !player.ignore);
    console.log('filteredPlayers:', filteredPlayers);


    // 排序
    if (sortBy === 'stats') {
        filteredPlayers.sort((a, b) => b.stats - a.stats);
    } else if (sortBy === 'sixStats') {
        filteredPlayers.sort((a, b) => b.sixStats - a.sixStats);
    } else if (sortBy === 'cavalryArchery') {
        filteredPlayers.sort((a, b) => b.cavalryArchery - a.cavalryArchery);
    } else if (sortBy === 'attack') {
        filteredPlayers.sort((a, b) => b.attack - a.attack);
    }

    // 按分工分组
    const roleGroups = {
        '大车头': filteredPlayers.slice(0, 3), // 前3名为大车头
        '二车头': filteredPlayers.slice(3, 10), // 4-10名为二车头
        '中战': filteredPlayers.slice(10, 25), // 11-25名为中战
        '低战/辅助': filteredPlayers.slice(25) // 26名以后为低战/辅助
    };

    // 重置mapData
    global.mapData = [];

    // 分配玩家到建筑
    const availableBuildings = buildings.filter(building => building.phase <= phase);

    // 按建筑优先级排序
    availableBuildings.sort((a, b) => a.priority - b.priority);

    availableBuildings.forEach(building => {
        let playersToAssign = [];

        // 根据建筑类型和阶段分配不同分工的玩家
        if (phase === 1) {
            playersToAssign = assignPlayersToBuildingPhase1(building, roleGroups);
        } else if (phase === 2) {
            playersToAssign = assignPlayersToBuildingPhase2(building, roleGroups);
        } else if (phase === 3) {
            playersToAssign = assignPlayersToBuildingPhase3(building, roleGroups);
        }

        // 将分配的玩家添加到mapData
        playersToAssign.forEach(player => {
            if (!global.mapData.find(row => row.name === player.name)) {
                global.mapData.push({
                    ...player,
                    buildingId: building.id,
                    buildingName: building.name,
                    rank: global.mapData.length + 1
                });
            }
        });
    });

    // 更新结果表格
    updateResultsTable();
    updateStats();
    console.log('mapData:', global.mapData);
}

// 第一阶段（9:00-9:15）建筑分配
function assignPlayersToBuildingPhase1(building, roleGroups) {
    switch (building.id) {
        case 'bingqifang':
            return roleGroups['大车头']; // 兵器坊由大车头占领
        case 'gongjiangfang':
            return roleGroups['二车头'].slice(0, 3); // 工匠坊由二车头占领
        case 'xiaoliangcang-left-1':
        case 'xiaoliangcang-left-2':
        case 'xiaoliangcang-left-3':
        case 'xiaoliangcang-right-1':
        case 'xiaoliangcang-right-2':
        case 'xiaoliangcang-right-3':
            return roleGroups['中战'].slice(0, 2); // 小粮仓由中战占领
        default:
            return [];
    }
}

// 第二阶段（9:15-9:25）建筑分配
function assignPlayersToBuildingPhase2(building, roleGroups) {
    if (building.phase <= 1) {
        return assignPlayersToBuildingPhase1(building, roleGroups);
    }

    switch (building.id) {
        case 'wuchao':
            return [...roleGroups['大车头'], ...roleGroups['二车头'].slice(0, 2)]; // 乌巢由大车头+部分二车头占领
        default:
            return [];
    }
}

// 第三阶段（9:25-9:40）建筑分配
function assignPlayersToBuildingPhase3(building, roleGroups) {
    if (building.phase <= 2) {
        return assignPlayersToBuildingPhase2(building, roleGroups);
    }

    switch (building.id) {
        case 'guandu':
            return [...roleGroups['大车头'], ...roleGroups['二车头']]; // 官渡由大车头和二车头占领
        case 'piliuche':
            return roleGroups['二车头'].slice(3, 7); // 霹雳车由部分二车头占领
        case 'dalangcang-left':
        case 'dalangcang-right':
            return roleGroups['中战'].slice(10, 14); // 大粮仓由中战占领
        default:
            return [];
    }
}

// 处理数据和分配玩家
function processData() {
    allocatePlayers();
    renderMap();
}

// 渲染地图
function renderMap() {
    const mapContainer = document.getElementById('map');
    mapContainer.innerHTML = '';
    mapContainer.className = 'map-container';

    // 地图网格布局
    for (let y = 0; y < 5; y++) {
        for (let x = -1; x <= 1; x++) {
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.dataset.x = x;
            cell.dataset.y = y;

            const cellContent = document.createElement('div');
            cellContent.className = 'cell-content';

            // 查找是否有建筑
            const building = buildings.find(b => b.position.x === x && b.position.y === y);
            if (building) {
                cell.className += ` building ${building.id}`;
                cellContent.innerHTML = `
                    <div class="building-name">${building.name}</div>
                    <div class="building-info">
                        <div class="open-time">${building.openTime}</div>
                        <div class="points">${building.pointsPerMinute}</div>
                    </div>
                `;
            }

            cell.appendChild(cellContent);
            mapContainer.appendChild(cell);
        }
    }

    // 渲染时间线
    renderTimeLine();
}

// 时间线渲染
function renderTimeLine() {
    const timeLineContainer = document.createElement('div');
    timeLineContainer.className = 'time-line';

    const timePoints = [
        { time: '9:00', phase: 1, description: '开局' },
        { time: '9:15', phase: 2, description: '乌巢开放' },
        { time: '9:25', phase: 3, description: '官渡开放' },
        { time: '9:40', phase: 3, description: '结束' }
    ];

    timePoints.forEach(point => {
        const timePoint = document.createElement('div');
        timePoint.className = `time-point ${parseInt(document.getElementById('phase').value) === point.phase ? 'active' : ''}`;
        timePoint.textContent = `${point.time} - ${point.description}`;
        timeLineContainer.appendChild(timePoint);
    });

    const mapContainer = document.getElementById('map-container');
    mapContainer.appendChild(timeLineContainer);
}

// 更新结果表格
function updateResultsTable() {
    const tbody = document.getElementById('resultsTableBody');
    tbody.innerHTML = '';

    playerData.forEach(player => {
        const mapRow = global.mapData.find(row => row.name === player.name);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${mapRow ? mapRow.rank : '-'}</td>
            <td>${player.name}</td>
            <td>${player.stats.toFixed(0)}</td>
            <td>${player.sixStats.toFixed(0)}</td>
            <td>${player.cavalryArchery.toFixed(0)}</td>
            <td>${player.attack.toFixed(0)}</td>
            <td>${player.defense.toFixed(0)}</td>
            <td>${mapRow ? mapRow.role : '-'}</td>
            <td>${mapRow ? mapRow.buildingName : '-'}</td>
            <td>${player.legion}</td>
            <td><input type="checkbox" ${player.ignore ? 'checked' : ''} onchange="toggleIgnore('${player.name}')"></td>
        `;
        tbody.appendChild(tr);
    });
}

// 切换忽略状态
function toggleIgnore(playerName) {
    const player = playerData.find(p => p.name === playerName);
    if (player) {
        player.ignore = !player.ignore;
        const ignoredPlayers = JSON.parse(localStorage.getItem('ignoredPlayersGuandu') || '{}');
        ignoredPlayers[playerName] = player.ignore;
        localStorage.setItem('ignoredPlayersGuandu', JSON.stringify(ignoredPlayers));
        processData();
        showMessage(`玩家${playerName}已${player.ignore ? '忽略' : '恢复'}`, 'success');
    }
}

// 更新统计信息
function updateStats() {
    const totalPlayers = playerData.length;
    const assignedPlayers = global.mapData.length;
    const unassignedPlayers = playerData.filter(p => !p.ignore && !global.mapData.find(m => m.name === p.name)).length;
    const ignoredPlayers = playerData.filter(p => p.ignore).length;

    document.getElementById('totalPlayers').textContent = totalPlayers;
    document.getElementById('assignedPlayers').textContent = assignedPlayers;
    document.getElementById('unassignedPlayers').textContent = unassignedPlayers;
    document.getElementById('ignoredPlayers').textContent = ignoredPlayers;
}

// 导出结果
function downloadResults() {
    if (global.mapData.length === 0) {
        showMessage('没有可导出的数据', 'error');
        return;
    }

    const data = global.mapData.map(row => ({
        '排名': row.rank,
        '玩家姓名': row.name,
        '四维和': row.stats.toFixed(0),
        '六维和': row.sixStats.toFixed(0),
        '骑弓维(输出)': row.cavalryArchery.toFixed(0),
        '输出': row.attack.toFixed(0),
        '坦度': row.defense.toFixed(0),
        '分工': row.role,
        '分配建筑': row.buildingName,
        '军团': row.legion
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '分配结果');

    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    const fileName = `官渡之战分配结果_${year}${month}${day}${hours}${minutes}${seconds}.xlsx`;
    XLSX.writeFile(workbook, fileName);
}

// 显示消息
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

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', init);

// 导出函数供测试使用
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    allocatePlayers,
    processPlayerData,
    mapData,
    playerData
  };
}
