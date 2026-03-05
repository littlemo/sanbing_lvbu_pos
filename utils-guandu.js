// 工具函数库

// 格式化时间
function formatTime(date) {
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
}

// 获取当前阶段
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

// 计算两点之间的距离
function calculateDistance(x1, y1, x2, y2) {
    return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
}

// 验证Excel数据格式
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

// 生成随机ID
function generateId() {
    return Math.random().toString(36).substring(2, 10);
}

// 防抖函数
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// 节流函数
function throttle(func, limit) {
    let inThrottle;
    return function() {
        const args = arguments;
        const context = this;
        if (!inThrottle) {
            func.apply(context, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    }
}

// 格式化数字
function formatNumber(num) {
    if (num >= 10000) {
        return (num / 10000).toFixed(1) + '万';
    }
    return num.toString();
}

// 深拷贝对象
function deepClone(obj) {
    if (obj === null || typeof obj !== 'object') {
        return obj;
    }

    if (obj instanceof Date) {
        return new Date(obj.getTime());
    }

    if (obj instanceof Array) {
        return obj.map(item => deepClone(item));
    }

    if (typeof obj === 'object') {
        const clonedObj = {};
        for (let key in obj) {
            if (obj.hasOwnProperty(key)) {
                clonedObj[key] = deepClone(obj[key]);
            }
        }
        return clonedObj;
    }
}

// 数组去重
function removeDuplicates(arr, key) {
    if (!key) {
        return [...new Set(arr)];
    }

    const seen = new Map();
    return arr.filter(item => {
        const itemKey = item[key];
        if (!seen.has(itemKey)) {
            seen.set(itemKey, true);
            return true;
        }
        return false;
    });
}

// 随机颜色生成
function getRandomColor() {
    const letters = '0123456789ABCDEF';
    let color = '#';
    for (let i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
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
