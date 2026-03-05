# 官渡之战安排工具实现计划

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 为《三国·冰河时代》游戏设计和实现官渡之战玩家安排工具，帮助联盟指挥快速合理地分配玩家到各个建筑点位。

**Architecture:** 基于现有吕布校场排位工具的基础架构，使用HTML5 Canvas进行地图渲染，JavaScript处理玩家分配算法，SheetJS库处理Excel导入导出。

**Tech Stack:** 纯前端技术栈，包括HTML5、CSS3、JavaScript（ES6+）、Canvas、SheetJS库。

---

## 任务1: 创建项目基础架构

**Files:**
- 创建: `index-guandu.html` - 官渡之战工具的主页面
- 创建: `app-guandu.js` - 官渡之战工具的主逻辑文件
- 复制: `package.json` - 从现有项目复制依赖配置

**Step 1: 创建主页面文件**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>官渡之战安排工具 - 三国·冰河时代</title>
    <style>
        /* 基础样式，可从现有项目复制 */
        :root {
            --primary-color: #d4af37;
            --secondary-color: #3a3a3a;
            --accent-color: #c11b17;
            --bg-color: #1a1a1a;
            --text-color: #e0e0e0;
            --grid-bg: #2a2a2a;
            --cell-size: 80px;
            --border-radius: 4px;
        }

        /* 全局样式 */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        body {
            background: linear-gradient(135deg, #1a1a1a 0%, #2a2a2a 100%);
            color: var(--text-color);
            min-height: 100vh;
            line-height: 1.6;
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 12px;
            backdrop-filter: blur(10px);
        }

        /* 其他样式从现有项目复制... */
    </style>
</head>
<body>
    <div class="container">
        <h1>官渡之战安排工具</h1>
        <div style="text-align: center; color: var(--text-color); margin-bottom: 20px;">作者：#620 雪色梦貘</div>

        <div class="controls">
            <div class="control-group">
                <label for="excelFile">上传 Excel 文件</label>
                <input type="file" id="excelFile" accept=".xlsx, .xls">
            </div>

            <div class="control-group">
                <label for="phase">战役阶段</label>
                <select id="phase">
                    <option value="1">第一阶段（9:00-9:15）</option>
                    <option value="2">第二阶段（9:15-9:25）</option>
                    <option value="3">第三阶段（9:25-9:40）</option>
                </select>
            </div>

            <div class="control-group">
                <label for="sortBy">排位方式</label>
                <select id="sortBy">
                    <option value="stats">四维和</option>
                    <option value="sixStats">六维和</option>
                    <option value="cavalryArchery">骑弓维</option>
                    <option value="attack">输出</option>
                </select>
            </div>

            <div class="control-group">
                <label for="legion">军团</label>
                <select id="legion">
                    <option value="1">军团1</option>
                    <option value="2">军团2</option>
                </select>
            </div>

            <div style="display: flex; flex-direction: column; justify-content: flex-end; gap: 10px;">
                <button class="btn" onclick="downloadResults()">导出结果</button>
            </div>
        </div>

        <div class="stats">
            <div class="stat-item">
                <div class="stat-value" id="totalPlayers">0</div>
                <div class="stat-label">总玩家数</div>
            </div>
            <div class="stat-item">
                <div class="stat-value" id="assignedPlayers">0</div>
                <div class="stat-label">已分配</div>
            </div>
            <div class="stat-item">
                <div class="stat-value" id="unassignedPlayers">0</div>
                <div class="stat-label">未分配</div>
            </div>
        </div>

        <div id="map-container">
            <div id="map"></div>
        </div>

        <div id="results">
            <h2 style="text-align: center; color: var(--primary-color); margin-bottom: 15px;">分配结果</h2>
            <table>
                <thead>
                    <tr>
                        <th>排名</th>
                        <th>玩家姓名</th>
                        <th>四维和</th>
                        <th>六维和</th>
                        <th>骑弓维</th>
                        <th>输出</th>
                        <th>坦度</th>
                        <th>分工</th>
                        <th>分配建筑</th>
                        <th>军团</th>
                        <th>忽略</th>
                    </tr>
                </thead>
                <tbody id="resultsTableBody">
                </tbody>
            </table>
        </div>
    </div>

    <div id="message"></div>

    <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
    <script src="app-guandu.js"></script>
</body>
</html>
```

**Step 2: 创建主逻辑文件**

```javascript
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

// 处理数据和分配玩家
function processData() {
    const phase = parseInt(document.getElementById('phase').value);
    const sortBy = document.getElementById('sortBy').value;
    const legion = parseInt(document.getElementById('legion').value);

    // 筛选当前军团的玩家
    let filteredPlayers = playerData.filter(player => player.legion === legion && !player.ignore);

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

    // 分配玩家到建筑
    mapData = [];
    const availableBuildings = buildings.filter(building => building.phase <= phase);

    // 按建筑优先级排序
    availableBuildings.sort((a, b) => a.priority - b.priority);

    // 分配逻辑
    availableBuildings.forEach(building => {
        // 为每个建筑分配玩家
        const neededPlayers = getPlayersNeeded(building);
        if (neededPlayers > 0 && filteredPlayers.length > 0) {
            const assignedPlayers = filteredPlayers.slice(0, neededPlayers);
            filteredPlayers = filteredPlayers.slice(neededPlayers);

            assignedPlayers.forEach(player => {
                mapData.push({
                    ...player,
                    buildingId: building.id,
                    buildingName: building.name,
                    rank: mapData.length + 1
                });
            });
        }
    });

    // 渲染地图和表格
    renderMap();
    updateResultsTable();
    updateStats();
}

// 渲染地图
function renderMap() {
    const mapContainer = document.getElementById('map');
    mapContainer.innerHTML = '';

    // 地图网格布局
    const rows = 5; // 从上到下5行
    const cols = 3; // 左右对称3列

    for (let y = 0; y < rows; y++) {
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
                cellContent.innerHTML = `<strong>${building.name}</strong><br><small>(${building.openTime})</small>`;
            }

            cell.appendChild(cellContent);
            mapContainer.appendChild(cell);
        }
    }
}

// 获取建筑需要的玩家数量
function getPlayersNeeded(building) {
    switch (building.id) {
        case 'wuchao':
            return 5; // 乌巢需要5人
        case 'bingqifang':
            return 3; // 兵器坊需要3人
        case 'guandu':
            return 8; // 官渡需要8人
        case 'piliuche':
            return 4; // 霹雳车需要4人
        case 'gongjiangfang':
            return 3; // 工匠坊需要3人
        case 'dalangcang-left':
        case 'dalangcang-right':
            return 4; // 大粮仓需要4人
        case 'xiaoliangcang-left-1':
        case 'xiaoliangcang-left-2':
        case 'xiaoliangcang-left-3':
        case 'xiaoliangcang-right-1':
        case 'xiaoliangcang-right-2':
        case 'xiaoliangcang-right-3':
            return 2; // 小粮仓需要2人
        default:
            return 1;
    }
}

// 更新结果表格
function updateResultsTable() {
    const tbody = document.getElementById('resultsTableBody');
    tbody.innerHTML = '';

    playerData.forEach(player => {
        const mapRow = mapData.find(row => row.name === player.name);

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
    }
}

// 更新统计信息
function updateStats() {
    const totalPlayers = playerData.length;
    const assignedPlayers = mapData.length;
    const unassignedPlayers = totalPlayers - assignedPlayers;

    document.getElementById('totalPlayers').textContent = totalPlayers;
    document.getElementById('assignedPlayers').textContent = assignedPlayers;
    document.getElementById('unassignedPlayers').textContent = unassignedPlayers;
}

// 导出结果
function downloadResults() {
    if (mapData.length === 0) {
        showMessage('没有可导出的数据', 'error');
        return;
    }

    const data = mapData.map(row => ({
        '排名': row.rank,
        '玩家姓名': row.name,
        '四维和': row.stats.toFixed(0),
        '六维和': row.sixStats.toFixed(0),
        '骑弓维': row.cavalryArchery.toFixed(0),
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
```

**Step 3: 复制package.json文件**

从现有项目复制package.json文件到官渡之战工具目录。

**Step 4: 测试页面是否能正常加载**

在浏览器中打开index-guandu.html，检查页面是否能正常加载，是否有JavaScript错误。

**Step 5: 提交代码**

```bash
git add index-guandu.html app-guandu.js
git commit -m "feat: 创建官渡之战安排工具基础架构"
```

---

## 任务2: 实现地图可视化和建筑渲染

**Files:**
- Modify: `app-guandu.js:123-150` - 优化地图渲染逻辑
- Modify: `index-guandu.html:50-80` - 优化地图容器样式
- Create: `styles-guandu.css` - 地图和建筑样式

**Step 1: 创建地图样式文件**

```css
/* 地图和建筑样式 */
#map-container {
    overflow-x: auto;
    margin-bottom: 30px;
    padding: 20px;
    background: rgba(0, 0, 0, 0.3);
    border-radius: 8px;
    display: flex;
    justify-content: center;
    min-height: 600px;
    align-items: center;
}

#map {
    display: grid;
    gap: 10px;
    background: #3a3a3a;
    padding: 10px;
    border-radius: var(--border-radius);
    grid-template-rows: repeat(5, 100px);
    grid-template-columns: repeat(3, 150px);
}

.cell {
    width: 150px;
    height: 100px;
    background: #2a2a2a;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    font-size: 0.8rem;
    text-align: center;
    padding: 4px;
    border-radius: 4px;
    position: relative;
    cursor: pointer;
    transition: all 0.3s ease;
    color: var(--text-color);
}

.cell-content {
    width: 100%;
    height: 100%;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    transition: all 0.3s ease;
}

.cell:hover {
    background: #3a3a3a;
    transform: scale(1.05);
    z-index: 10;
}

/* 建筑特殊样式 */
.building {
    background: linear-gradient(135deg, var(--primary-color), #8b6914);
    font-weight: bold;
    color: white;
}

.building.wuchao {
    background: linear-gradient(135deg, #ff6b6b, #ee5a24);
}

.building.bingqifang {
    background: linear-gradient(135deg, #4ecdc4, #26a69a);
}

.building.guandu {
    background: linear-gradient(135deg, #ffd93d, #ffb347);
    font-size: 1.1rem;
}

.building.piliuche {
    background: linear-gradient(135deg, #a8e6cf, #7fcdbb);
}

.building.gongjiangfang {
    background: linear-gradient(135deg, #d4a5a5, #c06c84);
}

.building.dalangcang-left,
.building.dalangcang-right {
    background: linear-gradient(135deg, #feca57, #ff9ff3);
}

.building.xiaoliangcang-left-1,
.building.xiaoliangcang-left-2,
.building.xiaoliangcang-left-3,
.building.xiaoliangcang-right-1,
.building.xiaoliangcang-right-2,
.building.xiaoliangcang-right-3 {
    background: linear-gradient(135deg, #54a0ff, #0984e3);
}

/* 时间线样式 */
.time-line {
    margin: 20px 0;
    padding: 15px;
    background: rgba(0, 0, 0, 0.3);
    border-radius: 8px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
}

.time-point {
    padding: 5px 10px;
    background: var(--primary-color);
    border-radius: 4px;
    font-weight: bold;
    color: white;
    font-size: 0.9rem;
}

.time-point.active {
    background: var(--accent-color);
    transform: scale(1.1);
}

/* 响应式设计 */
@media (max-width: 768px) {
    #map {
        grid-template-columns: repeat(3, 120px);
        grid-template-rows: repeat(5, 80px);
    }

    .cell {
        width: 120px;
        height: 80px;
        font-size: 0.7rem;
    }
}
```

**Step 2: 优化地图渲染逻辑**

```javascript
// 地图渲染优化
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
```

**Step 3: 优化HTML文件**

```html
<!-- 在index-guandu.html中添加地图容器样式 -->
<div id="map-container">
    <div id="map"></div>
</div>
```

**Step 4: 测试地图渲染**

在浏览器中打开index-guandu.html，检查地图是否能正常渲染，建筑是否显示正确。

**Step 5: 提交代码**

```bash
git add styles-guandu.css app-guandu.js index-guandu.html
git commit -m "feat: 实现地图可视化和建筑渲染"
```

---

## 任务3: 实现玩家分配算法

**Files:**
- Modify: `app-guandu.js:80-120` - 实现智能分配算法
- Modify: `app-guandu.js:150-180` - 优化建筑需求计算
- Create: `tests/player-allocation.test.js` - 分配算法测试

**Step 1: 实现智能分配算法**

```javascript
// 智能分配算法
function allocatePlayers() {
    const phase = parseInt(document.getElementById('phase').value);
    const sortBy = document.getElementById('sortBy').value;
    const legion = parseInt(document.getElementById('legion').value);

    // 筛选当前军团的玩家
    let filteredPlayers = playerData.filter(player => player.legion === legion && !player.ignore);

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
    mapData = [];

    // 分配建筑
    const availableBuildings = buildings.filter(b => b.phase <= phase);

    // 按优先级排序建筑
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
            if (!mapData.find(row => row.name === player.name)) {
                mapData.push({
                    ...player,
                    buildingId: building.id,
                    buildingName: building.name,
                    rank: mapData.length + 1
                });
            }
        });
    });

    // 更新结果表格
    updateResultsTable();
    updateStats();
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
```

**Step 2: 优化分配函数**

```javascript
// 替换原有的processData函数
function processData() {
    allocatePlayers();
    renderMap();
}
```

**Step 3: 创建测试文件**

```javascript
// 分配算法测试
describe('Player Allocation', () => {
    describe('Role Groups', () => {
        it('should group players correctly by role', () => {
            // 创建模拟数据
            const mockPlayers = [];
            for (let i = 0; i < 30; i++) {
                mockPlayers.push({
                    name: `玩家${i+1}`,
                    stats: 10000 - i * 100,
                    sixStats: 15000 - i * 100,
                    cavalryArchery: 8000 - i * 100,
                    attack: 5000 - i * 100,
                    defense: 3000 - i * 100,
                    role: '中战',
                    legion: 1,
                    ignore: false
                });
            }

            playerData = mockPlayers;

            const sortBy = 'stats';
            const phase = 3;
            const legion = 1;

            let filteredPlayers = playerData.filter(player => player.legion === legion && !player.ignore);
            filteredPlayers.sort((a, b) => b.stats - a.stats);

            const roleGroups = {
                '大车头': filteredPlayers.slice(0, 3),
                '二车头': filteredPlayers.slice(3, 10),
                '中战': filteredPlayers.slice(10, 25),
                '低战/辅助': filteredPlayers.slice(25)
            };

            expect(roleGroups['大车头'].length).toBe(3);
            expect(roleGroups['二车头'].length).toBe(7);
            expect(roleGroups['中战'].length).toBe(15);
            expect(roleGroups['低战/辅助'].length).toBe(5);
        });
    });

    describe('Building Allocation', () => {
        it('should allocate correct number of players to buildings', () => {
            // 创建模拟数据
            const mockPlayers = [];
            for (let i = 0; i < 30; i++) {
                mockPlayers.push({
                    name: `玩家${i+1}`,
                    stats: 10000 - i * 100,
                    sixStats: 15000 - i * 100,
                    cavalryArchery: 8000 - i * 100,
                    attack: 5000 - i * 100,
                    defense: 3000 - i * 100,
                    role: '中战',
                    legion: 1,
                    ignore: false
                });
            }

            playerData = mockPlayers;
            document.getElementById('phase').value = 3;
            document.getElementById('sortBy').value = 'stats';
            document.getElementById('legion').value = 1;

            allocatePlayers();

            // 检查关键建筑分配
            const guanduPlayers = mapData.filter(row => row.buildingId === 'guandu');
            const bingqifangPlayers = mapData.filter(row => row.buildingId === 'bingqifang');
            const wuchaoPlayers = mapData.filter(row => row.buildingId === 'wuchao');

            expect(guanduPlayers.length).toBeGreaterThanOrEqual(8);
            expect(bingqifangPlayers.length).toBeGreaterThanOrEqual(3);
            expect(wuchaoPlayers.length).toBeGreaterThanOrEqual(5);
        });
    });
});
```

**Step 4: 运行测试**

```bash
npm install --save-dev jest
npm run test
```

**Step 5: 提交代码**

```bash
git add app-guandu.js tests/player-allocation.test.js
git commit -m "feat: 实现智能玩家分配算法"
```

---

## 任务4: 优化用户体验和功能完整性

**Files:**
- Modify: `app-guandu.js:180-220` - 优化Excel导入导出功能
- Modify: `app-guandu.js:220-250` - 优化忽略功能和统计
- Modify: `index-guandu.html:80-110` - 优化表格样式
- Create: `utils-guandu.js` - 工具函数库

**Step 1: 创建工具函数库**

```javascript
// 工具函数库
function formatTime(date) {
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
}

function getCurrentPhase() {
    const now = new Date();
    const hour = now.getHours();
    const minute = now.getMinutes();

    if (hour < 9 || (hour === 9 && minute < 15)) {
        return 1;
    } else if (hour === 9 && minute < 25) {
        return 2;
    } else if (hour === 9 && minute < 40) {
        return 3;
    } else {
        return 1; // 默认第一阶段
    }
}

function calculateDistance(x1, y1, x2, y2) {
    return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
}

function validateExcelData(data) {
    // 验证Excel数据格式
    const requiredFields = ['游戏ID', '四维和', '步维(坦度)', '弓维(输出)', '六维和', '骑弓维(输出)'];
    const headerRow = data[1];

    const missingFields = requiredFields.filter(field => !headerRow.includes(field));
    if (missingFields.length > 0) {
        throw new Error(`Excel文件缺少以下列：${missingFields.join(', ')}`);
    }

    return true;
}
```

**Step 2: 优化Excel导入功能**

```javascript
// 优化Excel导入
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (file) {
        readExcelFile(file)
            .then(jsonData => {
                try {
                    validateExcelData(jsonData);
                    processPlayerData(jsonData);
                    processData();
                    showMessage('Excel 文件读取成功！', 'success');
                } catch (error) {
                    showMessage('数据格式错误：' + error.message, 'error');
                }
            })
            .catch(error => {
                showMessage('读取文件失败：' + error.message, 'error');
            });
    }
}
```

**Step 3: 优化导出功能**

```javascript
// 优化导出功能
function downloadResults() {
    if (mapData.length === 0) {
        showMessage('没有可导出的数据', 'error');
        return;
    }

    const data = mapData.map(row => ({
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
    const fileName = `官渡之战分配结果_${formatTime(now)}.xlsx`;
    XLSX.writeFile(workbook, fileName);
}
```

**Step 4: 优化忽略功能**

```javascript
// 优化忽略功能
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
```

**Step 5: 优化统计信息**

```javascript
// 优化统计信息
function updateStats() {
    const totalPlayers = playerData.length;
    const assignedPlayers = mapData.length;
    const unassignedPlayers = playerData.filter(p => !p.ignore && !mapData.find(m => m.name === p.name)).length;
    const ignoredPlayers = playerData.filter(p => p.ignore).length;

    document.getElementById('totalPlayers').textContent = totalPlayers;
    document.getElementById('assignedPlayers').textContent = assignedPlayers;
    document.getElementById('unassignedPlayers').textContent = unassignedPlayers;
    document.getElementById('ignoredPlayers').textContent = ignoredPlayers;
}
```

**Step 6: 优化表格样式**

```html
<!-- 在index-guandu.html中添加统计信息 -->
<div class="stats">
    <div class="stat-item">
        <div class="stat-value" id="totalPlayers">0</div>
        <div class="stat-label">总玩家数</div>
    </div>
    <div class="stat-item">
        <div class="stat-value" id="assignedPlayers">0</div>
        <div class="stat-label">已分配</div>
    </div>
    <div class="stat-item">
        <div class="stat-value" id="unassignedPlayers">0</div>
        <div class="stat-label">未分配</div>
    </div>
    <div class="stat-item">
        <div class="stat-value" id="ignoredPlayers">0</div>
        <div class="stat-label">已忽略</div>
    </div>
</div>
```

**Step 7: 提交代码**

```bash
git add utils-guandu.js app-guandu.js index-guandu.html
git commit -m "feat: 优化用户体验和功能完整性"
```

---

## 任务5: 测试和部署

**Files:**
- Create: `tests/integration.test.js` - 集成测试
- Create: `deploy.md` - 部署说明
- Modify: `package.json:10-15` - 添加脚本命令

**Step 1: 创建集成测试**

```javascript
// 集成测试
describe('Integration Tests', () => {
    describe('Complete Flow', () => {
        it('should complete the entire allocation process', async () => {
            // 模拟用户交互
            const fileInput = document.getElementById('excelFile');
            const phaseSelect = document.getElementById('phase');
            const sortBySelect = document.getElementById('sortBy');
            const legionSelect = document.getElementById('legion');
            const downloadButton = document.querySelector('.btn');

            // 测试文件上传（使用模拟文件）
            const mockFile = new File([''], 'test-data.xlsx');
            fileInput.files = [mockFile];
            fileInput.dispatchEvent(new Event('change'));

            // 等待文件加载
            await new Promise(resolve => setTimeout(resolve, 1000));

            // 测试阶段切换
            phaseSelect.value = '3';
            phaseSelect.dispatchEvent(new Event('change'));

            // 测试排序切换
            sortBySelect.value = 'sixStats';
            sortBySelect.dispatchEvent(new Event('change'));

            // 测试军团切换
            legionSelect.value = '2';
            legionSelect.dispatchEvent(new Event('change'));

            // 检查地图是否渲染
            const mapContainer = document.getElementById('map');
            expect(mapContainer.children.length).toBeGreaterThan(0);

            // 检查表格是否显示
            const tableBody = document.getElementById('resultsTableBody');
            expect(tableBody.children.length).toBeGreaterThan(0);

            // 检查统计信息
            const totalPlayers = parseInt(document.getElementById('totalPlayers').textContent);
            expect(totalPlayers).toBeGreaterThan(0);
        });
    });
});
```

**Step 2: 创建部署说明**

```markdown
# 官渡之战安排工具部署说明

## 本地运行

1. 安装依赖（无额外依赖）
2. 在项目根目录下启动HTTP服务器
   ```bash
   python3 -m http.server 8000
   ```
3. 在浏览器中访问 `http://localhost:8000/index-guandu.html`

## 生产部署

1. 确保项目文件完整
2. 构建项目（如果需要）
3. 部署到任何静态文件服务器
4. 或直接部署到GitHub Pages

## 依赖

- SheetJS (xlsx) 库：从CDN加载，无需额外安装

## 浏览器支持

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+
```

**Step 3: 优化package.json**

```json
{
  "name": "guandu-battle-tool",
  "version": "1.0.0",
  "description": "官渡之战安排工具",
  "scripts": {
    "dev": "python3 -m http.server 8000",
    "build": "mkdir -p dist && cp index-guandu.html dist/ && cp app-guandu.js dist/ && cp styles-guandu.css dist/ && cp utils-guandu.js dist/",
    "test": "jest"
  },
  "devDependencies": {
    "jest": "^29.5.0"
  }
}
```

**Step 4: 运行测试**

```bash
npm run test
```

**Step 5: 测试构建**

```bash
npm run build
```

**Step 6: 提交代码**

```bash
git add tests/integration.test.js deploy.md package.json
git commit -m "feat: 完成测试和部署说明"
```

---

## 任务6: 项目整合和最终检查

**Files:**
- Modify: `README.md:10-30` - 添加官渡之战工具说明
- Create: `docs/guandu-battle-tool.md` - 工具使用说明
- Create: `examples/test-data.xlsx` - 测试数据文件

**Step 1: 更新README.md**

```markdown
# 三国·冰河时代游戏工具集

包含两个核心工具：

## 1. 吕布校场排位工具

**文件：**
- `index.html` - 主页面
- `app.js` - 主逻辑
- `styles.css` - 样式

**功能：**
- Excel文件解析
- 智能排序（四维和/输出/六维和/骑弓维）
- 网格布局可视化
- 自定义配置
- 导出功能

---

## 2. 官渡之战安排工具

**文件：**
- `index-guandu.html` - 主页面
- `app-guandu.js` - 主逻辑
- `styles-guandu.css` - 样式
- `utils-guandu.js` - 工具函数

**功能：**
- Excel文件解析
- 智能分配算法
- 地图可视化
- 阶段管理
- 时间线显示
- 导出功能

---

## 使用方法

### 1. 吕布校场排位工具

1. 准备包含玩家数据的Excel文件
2. 打开`index.html`
3. 上传文件并配置参数
4. 查看排位结果

### 2. 官渡之战安排工具

1. 准备包含玩家数据的Excel文件
2. 打开`index-guandu.html`
3. 上传文件并配置参数
4. 查看分配结果

---

## 部署说明

### 本地运行

```bash
npm run dev
```

### 生产部署

```bash
npm run build
```

部署`dist/`目录到任何静态文件服务器。
```

**Step 2: 创建使用说明**

```markdown
# 官渡之战安排工具使用说明

## 功能介绍

官渡之战安排工具是专为《三国·冰河时代》游戏设计的玩家分配辅助工具，帮助联盟指挥快速合理地分配玩家到各个建筑点位，最大化积分获取效率，提升获胜概率。

## 基本使用

### 1. 准备数据

准备包含以下列的Excel文件（xlsx格式）：
- **游戏ID**：玩家姓名
- **四维和**：四维属性总和（步维+弓维+骑维+智维）
- **步维(坦度)**：步维属性值
- **弓维(输出)**：弓维属性值
- **六维和**：六维属性总和（步维+弓维+骑维+智维+政维+统维）
- **骑弓维(输出)**：骑维+弓维属性值

### 2. 上传文件

点击"上传 Excel 文件"按钮，选择准备好的数据文件。

### 3. 配置参数

- **战役阶段**：选择当前战役阶段（1-3阶段）
- **排位方式**：选择玩家排序方式（四维和/六维和/骑弓维/输出）
- **军团**：选择要分配的军团（军团1或军团2）

### 4. 查看结果

- **地图视图**：显示战役地图和建筑位置，点击建筑可查看详细信息
- **表格视图**：显示玩家分配详情，包括排名、姓名、属性、分工和分配建筑
- **统计信息**：显示总人数、已分配、未分配和已忽略玩家数
- **时间线**：显示战役关键时间节点

### 5. 导出结果

点击"导出结果"按钮，将分配结果导出为Excel文件。

## 高级功能

### 忽略玩家

在表格视图中，勾选玩家姓名旁的"忽略"复选框，该玩家将不计入分配。

### 分工说明

- **大车头**：战力前3名，负责关键建筑（乌巢、官渡）的占领
- **二车头**：战力4-10名，负责重要建筑（兵器坊、霹雳车）的占领
- **中战**：战力11-25名，负责小粮仓和大粮仓的占领
- **低战/辅助**：战力26名以后，负责驻防和补位

### 建筑优先级

**第一阶段（9:00-9:15）**：兵器坊 → 工匠坊 → 小粮仓
**第二阶段（9:15-9:25）**：乌巢 → 兵器坊 → 小粮仓
**第三阶段（9:25-9:40）**：官渡 → 大粮仓 → 霹雳车

## 注意事项

1. 确保Excel文件格式正确，包含所需列
2. 建议使用最新版本的浏览器以获得最佳体验
3. 导出功能需要浏览器支持FileSaver API
4. 玩家数据会保存在浏览器本地存储中，刷新页面不会丢失

## 技术支持

如遇到问题，请联系作者：#620 雪色梦貘
```

**Step 3: 创建测试数据**

创建一个包含10个玩家数据的测试Excel文件`examples/test-data.xlsx`。

**Step 4: 最终检查**

- 检查所有文件是否存在
- 检查代码是否有语法错误
- 检查页面是否能正常加载
- 检查功能是否正常运行

**Step 5: 提交代码**

```bash
git add README.md docs/guandu-battle-tool.md examples/test-data.xlsx
git commit -m "feat: 完成项目整合和最终检查"
```

---

## 任务7: 运行完整流程测试

**Files:**
- Run: `npm run test` - 运行所有测试
- Run: `npm run build` - 构建项目
- Test: 浏览器测试

**Step 1: 运行所有测试**

```bash
npm run test
```

**Step 2: 构建项目**

```bash
npm run build
```

**Step 3: 浏览器测试**

1. 打开`dist/index-guandu.html`
2. 上传测试数据文件
3. 测试所有功能
4. 检查是否有JavaScript错误

**Step 4: 检查构建产物**

```bash
ls -la dist/
```

**Step 5: 提交最终代码**

```bash
git add dist/
git commit -m "feat: 完成项目构建"
```

---

## 项目完成

### 完成内容

- 实现了官渡之战安排工具的所有核心功能
- 还原了地图布局和建筑位置
- 实现了智能分配算法
- 提供了完整的用户界面
- 添加了详细的文档和测试

### 功能特性

1. **地图可视化**：还原了官渡之战的垂直分布地图
2. **建筑显示**：标注了所有建筑位置、开放时间和作用
3. **阶段管理**：支持三个阶段的自动/手动切换
4. **智能分配**：根据战力和分工分配玩家到不同建筑
5. **Excel导入导出**：支持Excel文件的导入和导出
6. **忽略功能**：支持忽略特定玩家
7. **响应式设计**：适配不同屏幕尺寸
8. **统计信息**：显示完整的统计信息

### 使用说明

1. 准备包含玩家数据的Excel文件
2. 打开`index-guandu.html`
3. 上传文件并配置参数
4. 查看分配结果
5. 导出结果（可选）

---

## 后续优化

### 功能扩展

- 实时协作功能
- AI分配建议
- 历史数据分析
- 移动端适配

### 界面优化

- 3D地图效果
- 动画效果
- 主题定制

### 性能优化

- 数据缓存
- 虚拟滚动
- 图片压缩
