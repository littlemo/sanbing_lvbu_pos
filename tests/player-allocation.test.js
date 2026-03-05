// 官渡之战安排工具测试
// 测试智能分配算法的功能

// 导入必要的模块
const { allocatePlayers } = require('../app-guandu');
const fs = require('fs');
const path = require('path');
let mapData = [];
let playerData = [];

describe('官渡之战安排工具 - 智能分配算法', () => {
  // 初始化模拟数据
  let mapData = [];
  const mockPlayers = [
    { name: '玩家1', stats: 12000, sixStats: 18000, cavalryArchery: 9000, attack: 8500, defense: 6000, role: '大车头', legion: 1, ignore: false },
    { name: '玩家2', stats: 11500, sixStats: 17500, cavalryArchery: 8800, attack: 8200, defense: 5800, role: '大车头', legion: 1, ignore: false },
    { name: '玩家3', stats: 11000, sixStats: 17000, cavalryArchery: 8600, attack: 8000, defense: 5600, role: '大车头', legion: 1, ignore: false },
    { name: '玩家4', stats: 10500, sixStats: 16500, cavalryArchery: 8400, attack: 7800, defense: 5400, role: '二车头', legion: 1, ignore: false },
    { name: '玩家5', stats: 10000, sixStats: 16000, cavalryArchery: 8200, attack: 7600, defense: 5200, role: '二车头', legion: 1, ignore: false },
    { name: '玩家6', stats: 9500, sixStats: 15500, cavalryArchery: 8000, attack: 7400, defense: 5000, role: '二车头', legion: 1, ignore: false },
    { name: '玩家7', stats: 9000, sixStats: 15000, cavalryArchery: 7800, attack: 7200, defense: 4800, role: '二车头', legion: 1, ignore: false },
    { name: '玩家8', stats: 8500, sixStats: 14500, cavalryArchery: 7600, attack: 7000, defense: 4600, role: '二车头', legion: 1, ignore: false },
    { name: '玩家9', stats: 8000, sixStats: 14000, cavalryArchery: 7400, attack: 6800, defense: 4400, role: '二车头', legion: 1, ignore: false },
    { name: '玩家10', stats: 7500, sixStats: 13500, cavalryArchery: 7200, attack: 6600, defense: 4200, role: '二车头', legion: 1, ignore: false },
    { name: '玩家11', stats: 7000, sixStats: 13000, cavalryArchery: 7000, attack: 6400, defense: 4000, role: '中战', legion: 1, ignore: false },
    { name: '玩家12', stats: 6500, sixStats: 12500, cavalryArchery: 6800, attack: 6200, defense: 3800, role: '中战', legion: 1, ignore: false },
    // 更多模拟玩家...
  ];

  // 模拟DOM元素
  beforeAll(() => {
    // 创建模拟的DOM元素
    const phaseSelect = document.createElement('select');
    phaseSelect.id = 'phase';
    phaseSelect.innerHTML = '<option value="1">第一阶段</option><option value="2">第二阶段</option><option value="3">第三阶段</option>';
    document.body.appendChild(phaseSelect);

    const legionSelect = document.createElement('select');
    legionSelect.id = 'legion';
    legionSelect.innerHTML = '<option value="1">军团1</option><option value="2">军团2</option>';
    document.body.appendChild(legionSelect);

    const sortBySelect = document.createElement('select');
    sortBySelect.id = 'sortBy';
    sortBySelect.innerHTML = '<option value="stats">四维和</option><option value="sixStats">六维和</option><option value="cavalryArchery">骑弓维</option><option value="attack">输出</option>';
    document.body.appendChild(sortBySelect);

    // 模拟结果表格
    const resultsTable = document.createElement('table');
    const tbody = document.createElement('tbody');
    tbody.id = 'resultsTableBody';
    resultsTable.appendChild(tbody);
    document.body.appendChild(resultsTable);

    // 模拟统计元素
    const totalPlayersDiv = document.createElement('div');
    totalPlayersDiv.id = 'totalPlayers';
    document.body.appendChild(totalPlayersDiv);

    const assignedPlayersDiv = document.createElement('div');
    assignedPlayersDiv.id = 'assignedPlayers';
    document.body.appendChild(assignedPlayersDiv);

    const unassignedPlayersDiv = document.createElement('div');
    unassignedPlayersDiv.id = 'unassignedPlayers';
    document.body.appendChild(unassignedPlayersDiv);

    const ignoredPlayersDiv = document.createElement('div');
    ignoredPlayersDiv.id = 'ignoredPlayers';
    document.body.appendChild(ignoredPlayersDiv);

    // 设置模拟数据
    global.playerData = [...mockPlayers];
    global.mapData = [];
    mapData = global.mapData;
    playerData = global.playerData;
  });

  describe('智能分配算法', () => {
    describe('第一阶段分配', () => {
      it('应该分配玩家到第一阶段的建筑', () => {
        // 执行分配
        allocatePlayers(1, 'stats', 1);

        // 同步mapData变量
        mapData = global.mapData;

        // 检查是否分配了玩家
        expect(mapData.length).toBeGreaterThan(0);
      });

      it('应该为兵器坊分配大车头', () => {
        allocatePlayers(1, 'stats', 1);
        mapData = global.mapData;

        // 检查兵器坊是否分配了大车头
        const weaponWorkshopPlayers = mapData.filter(player => player.buildingId === 'bingqifang');
        expect(weaponWorkshopPlayers.length).toBeGreaterThan(0);
        // 检查是否有大车头
        const hasBigHead = weaponWorkshopPlayers.some(player => mockPlayers.slice(0, 3).includes(player));
        expect(hasBigHead).toBe(true);
      });
    });

    describe('第二阶段分配', () => {
      it('应该分配玩家到第二阶段的建筑', () => {
        allocatePlayers(2, 'stats', 1);
        mapData = global.mapData;

        expect(mapData.length).toBeGreaterThan(0);
      });

      it('应该为乌巢分配大车头和二车头', () => {
        allocatePlayers(2, 'stats', 1);
        mapData = global.mapData;

        const wuchaoPlayers = mapData.filter(player => player.buildingId === 'wuchao');
        expect(wuchaoPlayers.length).toBeGreaterThan(0);

        const bigHeadPlayers = mockPlayers.slice(0, 3);
        const secondHeadPlayers = mockPlayers.slice(3, 5);

        const hasBigHead = wuchaoPlayers.some(player => bigHeadPlayers.includes(player));
        const hasSecondHead = wuchaoPlayers.some(player => secondHeadPlayers.includes(player));

        expect(hasBigHead).toBe(true);
        expect(hasSecondHead).toBe(true);
      });
    });

    describe('第三阶段分配', () => {
      it('应该分配玩家到第三阶段的建筑', () => {
        allocatePlayers(3, 'stats', 1);
        mapData = global.mapData;

        expect(mapData.length).toBeGreaterThan(0);
      });

      it('应该为官渡分配大车头和二车头', () => {
        allocatePlayers(3, 'stats', 1);
        mapData = global.mapData;

        const guanduPlayers = mapData.filter(player => player.buildingId === 'guandu');
        expect(guanduPlayers.length).toBeGreaterThan(0);

        const bigHeadPlayers = mockPlayers.slice(0, 3);
        const secondHeadPlayers = mockPlayers.slice(3, 10);

        const hasBigHead = guanduPlayers.some(player => bigHeadPlayers.includes(player));
        const hasSecondHead = guanduPlayers.some(player => secondHeadPlayers.includes(player));

        expect(hasBigHead).toBe(true);
        expect(hasSecondHead).toBe(true);
      });

      it('应该为霹雳车分配二车头', () => {
        allocatePlayers(3, 'stats', 1);
        mapData = global.mapData;

        const piliuchePlayers = mapData.filter(player => player.buildingId === 'piliuche');
        expect(piliuchePlayers.length).toBeGreaterThan(0);

        const secondHeadPlayers = mockPlayers.slice(3, 10);

        const hasSecondHead = piliuchePlayers.some(player => secondHeadPlayers.includes(player));
        expect(hasSecondHead).toBe(true);
      });

      it('应该为大粮仓分配中战', () => {
        allocatePlayers(3, 'stats', 1);
        mapData = global.mapData;

        const granaryPlayers = mapData.filter(player => ['dalangcang-left', 'dalangcang-right'].includes(player.buildingId));
        expect(granaryPlayers.length).toBeGreaterThan(0);

        const midPlayers = mockPlayers.slice(10, 25);

        const hasMidPlayers = granaryPlayers.some(player => midPlayers.includes(player));
        expect(hasMidPlayers).toBe(true);
      });
    });

    describe('排序功能', () => {
      it('应该按四维和降序排序', () => {
        allocatePlayers(1, 'stats', 1);
        mapData = global.mapData;

        // 检查是否按四维和降序排序
        for (let i = 0; i < mapData.length - 1; i++) {
          expect(mapData[i].stats).toBeGreaterThanOrEqual(mapData[i + 1].stats);
        }
      });

      it('应该按六维和降序排序', () => {
        allocatePlayers(1, 'sixStats', 1);
        mapData = global.mapData;

        for (let i = 0; i < mapData.length - 1; i++) {
          expect(mapData[i].sixStats).toBeGreaterThanOrEqual(mapData[i + 1].sixStats);
        }
      });

      it('应该按骑弓维降序排序', () => {
        allocatePlayers(1, 'cavalryArchery', 1);
        mapData = global.mapData;

        for (let i = 0; i < mapData.length - 1; i++) {
          expect(mapData[i].cavalryArchery).toBeGreaterThanOrEqual(mapData[i + 1].cavalryArchery);
        }
      });

      it('应该按输出降序排序', () => {
        allocatePlayers(1, 'attack', 1);
        mapData = global.mapData;

        for (let i = 0; i < mapData.length - 1; i++) {
          expect(mapData[i].attack).toBeGreaterThanOrEqual(mapData[i + 1].attack);
        }
      });
    });
  });
});
