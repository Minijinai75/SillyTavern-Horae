/** Horae - 時間工具函數 */

/** 中文周幾映射 */
const WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'];

/** 季節名稱 */
const SEASONS = ['冬季', '冬季', '春季', '春季', '春季', '夏季', '夏季', '夏季', '秋季', '秋季', '秋季', '冬季'];

/** 中文數字映射 */
const CHINESE_NUMS = {
    '零': 0, '〇': 0,
    '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
    '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
    '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15,
    '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20,
    '廿': 20, '廿一': 21, '廿二': 22, '廿三': 23, '廿四': 24, '廿五': 25,
    '廿六': 26, '廿七': 27, '廿八': 28, '廿九': 29, '三十': 30,
    '三十一': 31, '卅': 30, '卅一': 31
};

/** 從日期字元串中提取日數 */
function extractDayNumber(dateStr) {
    if (!dateStr) return null;
    
    const arabicMatch = dateStr.match(/(?:第|Day\s*|day\s*)(\d+)(?:日)?/i) ||
                       dateStr.match(/(\d+)(?:日|号)/);
    if (arabicMatch) return parseInt(arabicMatch[1]);
    
    // 中文數字配對
    const sortedEntries = Object.entries(CHINESE_NUMS).sort((a, b) => b[0].length - a[0].length);
    
    for (const [cn, num] of sortedEntries) {
        const patterns = [
            new RegExp(`第${cn}日`),
            new RegExp(`第${cn}(?![\u4e00-\u9fa5])`),  // 第X 後面不跟漢字
            new RegExp(`[月]${cn}日`),
            new RegExp(`${cn}日`)
        ];
        
        for (const pattern of patterns) {
            if (pattern.test(dateStr)) {
                return num;
            }
        }
    }
    
    const anyNumMatch = dateStr.match(/(\d+)/);
    if (anyNumMatch) return parseInt(anyNumMatch[1]);
    
    return null;
}

/** 從日期字元串中提取月份標識 */
function extractMonthIdentifier(dateStr) {
    if (!dateStr) return null;
    
    // 配對"X月"格式
    const monthMatch = dateStr.match(/([^\s\d]+月)/);
    if (monthMatch) return monthMatch[1];
    
    const numMatch = dateStr.match(/(?:\d{4}[\/\-])?(\d{1,2})[\/\-]\d{1,2}/);
    if (numMatch) return numMatch[1] + '月';
    
    return null;
}

/** 解析劇情日期字元串 */
export function parseStoryDate(dateStr) {
    if (!dateStr) return null;
    
    // 清理AI寫的周幾標註
    let cleanStr = dateStr.trim();
    
    const aiWeekdayMatch = cleanStr.match(/\(([日一二三四五六])\)/);
    cleanStr = cleanStr.replace(/\s*\([日一二三四五六]\)\s*/g, ' ').trim();
    
    // 無效日期按奇幻日曆處理
    if (/[xX]{2}|[?？]{2}/.test(cleanStr)) {
        return { 
            type: 'fantasy',
            raw: dateStr.trim(),
            aiWeekday: aiWeekdayMatch ? aiWeekdayMatch[1] : undefined
        };
    }
    
    // 標準數字格式
    const fullMatch = cleanStr.match(/^(\d{4,})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (fullMatch) {
        const year = parseInt(fullMatch[1]);
        const month = parseInt(fullMatch[2]);
        const day = parseInt(fullMatch[3]);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return { year, month, day, type: 'standard' };
        }
    }
    
    const shortMatch = cleanStr.match(/^(\d{1,2})[\/\-](\d{1,2})(?:\s|$)/);
    if (shortMatch) {
        const month = parseInt(shortMatch[1]);
        const day = parseInt(shortMatch[2]);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return { month, day, type: 'standard' };
        }
    }
    
    // X年M月D日格式
    // 這個必須在純 X月X日 之前，否則會遺失年份
    const yearCnMatch = cleanStr.match(/(\d+)年\s*(\d{1,2})月(\d{1,2})日?/);
    if (yearCnMatch) {
        const year = parseInt(yearCnMatch[1]);
        const month = parseInt(yearCnMatch[2]);
        const day = parseInt(yearCnMatch[3]);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            // 提取曆法前綴
            const fullMatchStr = yearCnMatch[0];
            const prefixEnd = cleanStr.indexOf(fullMatchStr);
            const calendarPrefix = cleanStr.substring(0, prefixEnd).trim() || undefined;
            return { year, month, day, type: 'standard', calendarPrefix };
        }
    }
    
    // X月X日格式
    const cnMatch = cleanStr.match(/(\d{1,2})月(\d{1,2})日?/);
    if (cnMatch) {
        const month = parseInt(cnMatch[1]);
        const day = parseInt(cnMatch[2]);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return { month, day, type: 'standard' };
        }
    }
    
    // 奇幻日曆格式
    const monthId = extractMonthIdentifier(cleanStr);
    const dayNum = extractDayNumber(cleanStr);
    
    if (monthId || dayNum !== null) {
        return { 
            monthId: monthId,
            day: dayNum,
            type: 'fantasy',
            raw: dateStr.trim(),
            aiWeekday: aiWeekdayMatch ? aiWeekdayMatch[1] : undefined
        };
    }
    
    return null;
}

/** 計算兩個日期之間的天數差 */
export function calculateRelativeTime(fromDate, toDate) {
    if (!fromDate || !toDate) return null;
    
    // 去掉尾部時間部分（如 "15:00" / "下午" / "酉時"），保留完整日期進行比較
    const stripTime = (s) => s.trim()
        .replace(/\s+\d{1,2}[:：]\d{2}.*$/, '')
        .replace(/\s+(凌晨|早上|上午|中午|下午|傍晚|晚上|深夜|子时|丑时|寅时|卯时|辰时|巳时|午时|未时|申时|酉时|戌时|亥时).*$/i, '')
        .trim();
    const fromDateOnly = stripTime(fromDate);
    const toDateOnly = stripTime(toDate);
    
    if (fromDateOnly === toDateOnly) {
        return 0;
    }
    
    const from = parseStoryDate(fromDate);
    const to = parseStoryDate(toDate);
    
    if (!from || !to) return null;
    
    // 標準格式精確計算
    if (from.type === 'standard' && to.type === 'standard') {
        const defaultYear = 2024;
        const fromYear = from.year || to.year || defaultYear;
        const toYear = to.year || from.year || defaultYear;
        
        const fromObj = new Date(0);
        fromObj.setFullYear(fromYear, from.month - 1, from.day);
        const toObj = new Date(0);
        toObj.setFullYear(toYear, to.month - 1, to.day);
        
        const diffTime = toObj.getTime() - fromObj.getTime();
        return Math.round(diffTime / (1000 * 60 * 60 * 24));
    }
    
    if (from.type === 'fantasy' || to.type === 'fantasy') {
        const fromDay = from.day;
        const toDay = to.day;
        const fromMonth = from.monthId || from.month;
        const toMonth = to.monthId || to.month;
        
        // 同月精確計算
        if (fromMonth && toMonth && fromMonth === toMonth && 
            fromDay !== null && toDay !== null) {
            return toDay - fromDay;
        }
        
        // 跨月：舊邏輯用「日」大小猜先後，在西幻/架空月名日曆上極易誤判（如 霜月3日 vs 火月25日）
        if (fromDay !== null && toDay !== null) {
            if (fromMonth && toMonth && fromMonth !== toMonth) {
                return null;
            }
            return toDay - fromDay;
        }
        
        return -999;
    }
    
    return null;
}

/** 格式化相對時間描述 */
export function formatRelativeTime(days, options = {}) {
    if (days === null || days === undefined) return '未知';
    
    if (days === -999) return '較早';
    if (days === -998) return '之後';
    if (days === -997) return '之前';
    
    // 近幾天
    if (days === 0) return '今天';
    if (days === 1) return '昨天';
    if (days === 2) return '前天';
    if (days === 3) return '大前天';
    if (days === -1) return '明天';
    if (days === -2) return '後天';
    if (days === -3) return '大後天';
    
    const { fromDate, toDate } = options;
    
    if (days > 0) {
        if (days < 7) return `${days}天前`;
        
        // 上週幾
        if (days >= 4 && days <= 13 && fromDate) {
            const weekday = fromDate.getDay();
            return `上週${WEEKDAY_NAMES[weekday]}`;
        }
        
        // 上個月
        if (days >= 20 && days < 60 && fromDate && toDate) {
            const fromMonth = fromDate.getMonth();
            const toMonth = toDate.getMonth();
            if (fromMonth !== toMonth) {
                return `上個月${fromDate.getDate()}號`;
            }
        }
        
        if (days >= 300 && fromDate && toDate) {
            const fromYear = fromDate.getFullYear();
            const toYear = toDate.getFullYear();
            if (fromYear < toYear) {
                const fromMonth = fromDate.getMonth() + 1;
                const fromDay = fromDate.getDate();
                if (days < 730) {
                    return `去年${fromMonth}月${fromDay}日`;
                }
            }
        }
        
        if (days < 14) return `${Math.ceil(days / 7)}周前`;
        if (days < 60) return `${Math.round(days / 30)}個月前`;
        if (days < 365) return `${Math.round(days / 30)}個月前`;
        const years = Math.floor(days / 365);
        const remainMonths = Math.round((days % 365) / 30);
        if (remainMonths > 0 && years < 5) return `${years}年${remainMonths}個月前`;
        return `${years}年前`;
    } else {
        const absDays = Math.abs(days);
        if (absDays < 7) return `${absDays}天后`;
        
        if (absDays >= 4 && absDays <= 13 && fromDate) {
            const weekday = fromDate.getDay();
            return `下週${WEEKDAY_NAMES[weekday]}`;
        }
        
        if (absDays >= 20 && absDays < 60 && fromDate && toDate) {
            const fromMonth = fromDate.getMonth();
            const toMonth = toDate.getMonth();
            if (fromMonth !== toMonth) {
                return `下個月${fromDate.getDate()}號`;
            }
        }
        
        if (absDays < 14) return `${Math.ceil(absDays / 7)}周後`;
        if (absDays < 60) return `${Math.round(absDays / 30)}個月後`;
        if (absDays < 365) return `${Math.round(absDays / 30)}個月後`;
        const years = Math.floor(absDays / 365);
        const remainMonths = Math.round((absDays % 365) / 30);
        if (remainMonths > 0 && years < 5) return `${years}年${remainMonths}個月後`;
        return `${years}年後`;
    }
}

/** 格式化劇情日期為標準格式 */
export function formatStoryDate(dateObj, includeWeekday = false) {
    if (!dateObj) return '';
    // 奇幻日曆保留原始字元串
    if (dateObj.raw && !dateObj.month) {
        let result = dateObj.raw;
        if (includeWeekday && dateObj.aiWeekday && !result.includes(`(${dateObj.aiWeekday})`)) {
            result += ` (${dateObj.aiWeekday})`;
        }
        return result;
    }
    
    let dateStr = '';
    const prefix = dateObj.calendarPrefix || '';
    
    if (dateObj.year) {
        if (prefix) {
            // 保留曆法前綴
            dateStr = `${prefix}${dateObj.year}年${dateObj.month}月${dateObj.day}日`;
        } else {
            dateStr = `${dateObj.year}/${dateObj.month}/${dateObj.day}`;
        }
    } else if (dateObj.month && dateObj.day) {
        dateStr = `${dateObj.month}/${dateObj.day}`;
    }
    
    if (includeWeekday && dateObj.month && dateObj.day) {
        const refYear = dateObj.year || new Date().getFullYear();
        // setFullYear 避免年份自動偏移
        const date = new Date(0);
        date.setFullYear(refYear, dateObj.month - 1, dateObj.day);
        const weekday = WEEKDAY_NAMES[date.getDay()];
        dateStr += ` (${weekday})`;
    }
    
    return dateStr;
}

/** 格式化完整的劇情日期時間 */
export function formatFullDateTime(dateStr, timeStr) {
    const parsed = parseStoryDate(dateStr);
    if (!parsed) return dateStr + (timeStr ? ' ' + timeStr : '');
    
    const dateWithWeekday = formatStoryDate(parsed, true);
    return dateWithWeekday + (timeStr ? ' ' + timeStr : '');
}

/** 獲取目前系統時間 */
export function getCurrentSystemTime() {
    const now = new Date();
    return {
        date: `${now.getMonth() + 1}/${now.getDate()}`,
        time: `${now.getHours()}:${String(now.getMinutes()).padStart(2, '0')}`
    };
}

/** 生成時間參考資訊 */
export function generateTimeReference(currentDate) {
    const current = parseStoryDate(currentDate);
    if (!current) return null;
    
    if (current.type === 'fantasy') {
        return {
            current: currentDate,
            type: 'fantasy',
            note: '奇幻日曆模式，相對日期由插件自動計算'
        };
    }
    
    const refYear = current.year || new Date().getFullYear();
    const baseDate = new Date(0);
    baseDate.setFullYear(refYear, current.month - 1, current.day);
    
    const getDateString = (daysOffset) => {
        const d = new Date(baseDate.getTime());
        d.setDate(d.getDate() + daysOffset);
        const weekday = WEEKDAY_NAMES[d.getDay()];
        return `${d.getMonth() + 1}/${d.getDate()} (${weekday})`;
    };
    
    return {
        current: currentDate,
        type: 'standard',
        yesterday: getDateString(-1),
        dayBefore: getDateString(-2),
        threeDaysAgo: getDateString(-3),
        tomorrow: getDateString(1)
    };
}

/** 計算兩個日期之間的詳細差異 */
export function calculateDetailedRelativeTime(fromDateStr, toDateStr) {
    const days = calculateRelativeTime(fromDateStr, toDateStr);
    if (days === null) return { days: null, relative: '未知' };
    
    const from = parseStoryDate(fromDateStr);
    const to = parseStoryDate(toDateStr);
    
    let fromDate = null;
    let toDate = null;
    
    if (from?.type === 'standard' && to?.type === 'standard') {
        const defaultYear = new Date().getFullYear();
        const fromYear = from.year || to.year || defaultYear;
        const toYear = to.year || from.year || defaultYear;
        fromDate = new Date(0);
        fromDate.setFullYear(fromYear, from.month - 1, from.day);
        toDate = new Date(0);
        toDate.setFullYear(toYear, to.month - 1, to.day);
    }
    
    const relative = formatRelativeTime(days, { fromDate, toDate });
    
    return { days, fromDate, toDate, relative };
}

/** 從目前日期減去指定天數 */
export function subtractDays(dateStr, days) {
    const parsed = parseStoryDate(dateStr);
    if (!parsed || parsed.type === 'fantasy') return dateStr;
    
    const refYear = parsed.year || 2024;
    const date = new Date(0);
    date.setFullYear(refYear, parsed.month - 1, parsed.day);
    date.setDate(date.getDate() - days);
    
    if (parsed.year) {
        return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
    }
    return `${date.getMonth() + 1}/${date.getDate()}`;
}

/** 十二地支 → 起始小時（初=首小時，正=次小時） */
const EARTHLY_BRANCH_HOURS = {
    '子': 23, '醜': 1, '寅': 3, '卯': 5,
    '辰': 7, '巳': 9, '午': 11, '未': 13,
    '申': 15, '酉': 17, '戌': 19, '亥': 21
};

/** 獲取時間段描述 */
export function getTimeOfDay(timeStr) {
    if (!timeStr) return '';
    
    let hour = null;
    
    const match24 = timeStr.match(/(\d{1,2})[:：]/);
    if (match24) {
        hour = parseInt(match24[1]);
    }
    
    const matchCN = timeStr.match(/(凌晨|早上|上午|中午|下午|傍晚|晚上|深夜)/);
    if (matchCN) {
        return matchCN[1];
    }
    
    // 十二地支時辰兜底（子醜寅卯辰巳午未申酉戌亥 + 可選"時"/"初"/"正"）
    if (hour === null) {
        const branchMatch = timeStr.match(/([子丑寅卯辰巳午未申酉戌亥])时?(?:初|正)?/);
        if (branchMatch) {
            const base = EARTHLY_BRANCH_HOURS[branchMatch[0].charAt(0)];
            if (base !== undefined) {
                hour = /正/.test(branchMatch[0]) ? (base + 1) % 24 : base;
            }
        }
    }
    
    if (hour !== null) {
        if (hour >= 0 && hour < 5) return '凌晨';
        if (hour >= 5 && hour < 8) return '早上';
        if (hour >= 8 && hour < 11) return '上午';
        if (hour >= 11 && hour < 13) return '中午';
        if (hour >= 13 && hour < 17) return '下午';
        if (hour >= 17 && hour < 19) return '傍晚';
        if (hour >= 19 && hour < 23) return '晚上';
        return '深夜';
    }
    
    return '';
}
