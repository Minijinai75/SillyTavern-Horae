/**
 * Horae - 核心管理器
 * 負責元數據的存儲、解析、聚合
 */

import { parseStoryDate, calculateRelativeTime, calculateDetailedRelativeTime, generateTimeReference, formatRelativeTime, formatFullDateTime } from '../utils/timeUtils.js';

/**
 * @typedef {Object} HoraeTimestamp
 * @property {string} story_date - 劇情日期，如 "10/1"
 * @property {string} story_time - 劇情時間，如 "15:00" 或 "下午"
 * @property {string} absolute - ISO格式的實際時間戳
 */

/**
 * @typedef {Object} HoraeScene
 * @property {string} location - 場景地點
 * @property {string[]} characters_present - 在場角色列表
 * @property {string} atmosphere - 場景氛圍
 */

/**
 * @typedef {Object} HoraeEvent
 * @property {boolean} is_important - 是否重要事件
 * @property {string} level - 事件層級：一般/重要/關鍵
 * @property {string} summary - 事件摘要
 */

/**
 * @typedef {Object} HoraeItemInfo
 * @property {string|null} icon - emoji圖示
 * @property {string|null} holder - 持有者
 * @property {string} location - 位置描述
 */

/**
 * @typedef {Object} HoraeMeta
 * @property {HoraeTimestamp} timestamp
 * @property {HoraeScene} scene
 * @property {Object.<string, string>} costumes - 角色服裝 {角色名: 服裝描述}
 * @property {Object.<string, HoraeItemInfo>} items - 物品追蹤
 * @property {HoraeEvent|null} event
 * @property {Object.<string, string|number>} affection - 好感度
 * @property {Object.<string, {description: string, first_seen: string}>} npcs - 臨時NPC
 */

/** 建立空的元數據對象 */
export function createEmptyMeta() {
    return {
        timestamp: {
            story_date: '',
            story_time: '',
            absolute: ''
        },
        scene: {
            location: '',
            characters_present: [],
            atmosphere: ''
        },
        costumes: {},
        items: {},
        deletedItems: [],
        events: [],
        affection: {},
        npcs: {},
        agenda: [],
        mood: {},
        relationships: [],
    };
}

/**
 * 提取物品的基本名稱（去掉末尾的數量括號）
 * "新鮮牛大骨(5斤)" → "新鮮牛大骨"
 * "清水(9L)" → "清水"
 * "簡易急救包" → "簡易急救包"（無數量，不變）
 * "簡易急救包(已開封)" → 不變（非數字開頭的括號不去掉）
 */
// 個體量詞：1個 = 就一個，可省略。純量詞(個)(把)也無意義
const COUNTING_CLASSIFIERS = '個把條塊張根口份枚只顆支件套雙對碗杯盤盆串束扎';
// 容器/批量組織：1箱 = 一箱(裡面有很多)，不可省略
// 度量組織(斤/L/kg等)：有實際計量意義，不可省略

// 物品ID：3位數字左補零，如 001, 002, ...
function padItemId(id) { return String(id).padStart(3, '0'); }

export function getItemBaseName(name) {
    return name
        .replace(/[\(（][\d][\d\.\/]*[a-zA-Z\u4e00-\u9fff]*[\)）]$/, '')  // 數字+任意組織
        .replace(new RegExp(`[\\(（][${COUNTING_CLASSIFIERS}][\\)）]$`), '')  // 純個體量詞（AI錯誤格式）
        .trim();
}

/** 按基本名搜尋已有物品 */
function findExistingItemByBaseName(stateItems, newName) {
    const newBase = getItemBaseName(newName);
    if (stateItems[newName]) return newName;
    for (const existingName of Object.keys(stateItems)) {
        if (getItemBaseName(existingName) === newBase) {
            return existingName;
        }
    }
    return null;
}

/** Horae 管理器 */
class HoraeManager {
    constructor() {
        this.context = null;
        this.settings = null;
    }

    /** 初始化管理器 */
    init(context, settings) {
        this.context = context;
        this.settings = settings;
    }

    /** 獲取目前聊天記錄 */
    getChat() {
        return this.context?.chat || [];
    }

    /** 獲取訊息元數據 */
    getMessageMeta(messageIndex) {
        const chat = this.getChat();
        if (messageIndex < 0 || messageIndex >= chat.length) return null;
        return chat[messageIndex].horae_meta || null;
    }

    /** 設定訊息元數據 */
    setMessageMeta(messageIndex, meta) {
        const chat = this.getChat();
        if (messageIndex < 0 || messageIndex >= chat.length) return;
        chat[messageIndex].horae_meta = meta;
    }

    /** 聚合所有訊息元數據，獲取最新狀態 */
    getLatestState(skipLast = 0) {
        const chat = this.getChat();
        const state = createEmptyMeta();
        state._previousLocation = '';
        const end = Math.max(0, chat.length - skipLast);
        
        for (let i = 0; i < end; i++) {
            const meta = chat[i].horae_meta;
            if (!meta) continue;
            if (meta._skipHorae) continue;
            
            if (meta.timestamp?.story_date) {
                state.timestamp.story_date = meta.timestamp.story_date;
            }
            if (meta.timestamp?.story_time) {
                state.timestamp.story_time = meta.timestamp.story_time;
            }
            
            if (meta.scene?.location) {
                state._previousLocation = state.scene.location;
                state.scene.location = meta.scene.location;
            }
            if (meta.scene?.atmosphere) {
                state.scene.atmosphere = meta.scene.atmosphere;
            }
            if (meta.scene?.characters_present?.length > 0) {
                state.scene.characters_present = [...meta.scene.characters_present];
            }
            
            if (meta.costumes) {
                Object.assign(state.costumes, meta.costumes);
            }
            
            // 物品：合併更新
            if (meta.items) {
                for (let [name, newInfo] of Object.entries(meta.items)) {
                    // 去掉無意義的數量標記
                    // (1) 裸數字1 → 去掉
                    name = name.replace(/[\(（]1[\)）]$/, '').trim();
                    // 個體量詞+數字1 → 去掉
                    name = name.replace(new RegExp(`[\\(（]1[${COUNTING_CLASSIFIERS}][\\)）]$`), '').trim();
                    // 純個體量詞 → 去掉
                    name = name.replace(new RegExp(`[\\(（][${COUNTING_CLASSIFIERS}][\\)）]$`), '').trim();
                    // 度量/容器組織保留
                    
                    // 數量為0視為消耗，自動刪除
                    const zeroMatch = name.match(/[\(（]0[a-zA-Z\u4e00-\u9fff]*[\)）]$/);
                    if (zeroMatch) {
                        const baseName = getItemBaseName(name);
                        for (const itemName of Object.keys(state.items)) {
                            if (getItemBaseName(itemName).toLowerCase() === baseName.toLowerCase()) {
                                delete state.items[itemName];
                                console.log(`[Horae] 物品數量歸零自動刪除: ${itemName}`);
                            }
                        }
                        continue;
                    }
                    
                    // 檢測消耗狀態標記，視為刪除
                    const consumedPatterns = /[\(（](已消耗|已用完|已销毁|消耗殆尽|消耗|用尽)[\)）]/;
                    const holderConsumed = /^(消耗|已消耗|已用完|消耗殆尽|用尽|无)$/;
                    if (consumedPatterns.test(name) || holderConsumed.test(newInfo.holder || '')) {
                        const cleanName = name.replace(consumedPatterns, '').trim();
                        const baseName = getItemBaseName(cleanName || name);
                        for (const itemName of Object.keys(state.items)) {
                            if (getItemBaseName(itemName).toLowerCase() === baseName.toLowerCase()) {
                                delete state.items[itemName];
                                console.log(`[Horae] 物品已消耗自動刪除: ${itemName}`);
                            }
                        }
                        continue;
                    }
                    
                    // 基本名配對已有物品
                    const existingKey = findExistingItemByBaseName(state.items, name);
                    
                    if (existingKey) {
                        const existingItem = state.items[existingKey];
                        const mergedItem = { ...existingItem };
                        const locked = !!existingItem._locked;
                        if (!locked && newInfo.icon) mergedItem.icon = newInfo.icon;
                        if (!locked) {
                            const _impRank = { '': 0, '!': 1, '!!': 2 };
                            const _newR = _impRank[newInfo.importance] ?? 0;
                            const _oldR = _impRank[existingItem.importance] ?? 0;
                            mergedItem.importance = _newR >= _oldR ? (newInfo.importance || '') : (existingItem.importance || '');
                        }
                        if (newInfo.holder !== undefined) mergedItem.holder = newInfo.holder;
                        if (newInfo.location !== undefined) mergedItem.location = newInfo.location;
                        if (!locked && newInfo.description !== undefined && newInfo.description.trim()) {
                            mergedItem.description = newInfo.description;
                        }
                        if (!mergedItem.description) mergedItem.description = existingItem.description || '';
                        
                        if (existingKey !== name) {
                            delete state.items[existingKey];
                        }
                        state.items[name] = mergedItem;
                    } else {
                        state.items[name] = newInfo;
                    }
                }
            }
            
            // 處理已刪除物品
            if (meta.deletedItems && meta.deletedItems.length > 0) {
                for (const deletedItem of meta.deletedItems) {
                    const deleteBase = getItemBaseName(deletedItem).toLowerCase();
                    for (const itemName of Object.keys(state.items)) {
                        const itemBase = getItemBaseName(itemName).toLowerCase();
                        if (itemName.toLowerCase() === deletedItem.toLowerCase() ||
                            itemBase === deleteBase) {
                            delete state.items[itemName];
                        }
                    }
                }
            }
            
            // 好感度：支援絕對值和相對值
            if (meta.affection) {
                for (const [key, value] of Object.entries(meta.affection)) {
                    if (typeof value === 'object' && value !== null) {
                        // 新格式：{type: 'absolute'|'relative', value: number|string}
                        if (value.type === 'absolute') {
                            state.affection[key] = value.value;
                        } else if (value.type === 'relative') {
                            const delta = parseFloat(value.value) || 0;
                            state.affection[key] = (state.affection[key] || 0) + delta;
                        }
                    } else {
                        // 舊格式相容
                        const numValue = typeof value === 'number' ? value : parseFloat(value) || 0;
                        state.affection[key] = (state.affection[key] || 0) + numValue;
                    }
                }
            }
            
            // NPC：逐資料欄合併，保留_id
            if (meta.npcs) {
                // 可更新資料欄 vs 受保護資料欄
                const updatableFields = ['appearance', 'personality', 'relationship', 'age', 'job', 'note'];
                const protectedFields = ['gender', 'race', 'birthday'];
                for (const [name, newNpc] of Object.entries(meta.npcs)) {
                    const existing = state.npcs[name];
                    if (existing) {
                        for (const field of updatableFields) {
                            if (newNpc[field] !== undefined) existing[field] = newNpc[field];
                        }
                        // age變更時記錄劇情日期作為基準
                        if (newNpc.age !== undefined && newNpc.age !== '') {
                            if (!existing._ageRefDate) {
                                existing._ageRefDate = state.timestamp.story_date || '';
                            }
                            const oldAgeNum = parseInt(existing.age);
                            const newAgeNum = parseInt(newNpc.age);
                            if (!isNaN(oldAgeNum) && !isNaN(newAgeNum) && oldAgeNum !== newAgeNum) {
                                existing._ageRefDate = state.timestamp.story_date || '';
                            }
                        }
                        // 受保護資料欄：僅在未設定時才填入
                        for (const field of protectedFields) {
                            if (newNpc[field] !== undefined && !existing[field]) {
                                existing[field] = newNpc[field];
                            }
                        }
                        if (newNpc.last_seen) existing.last_seen = newNpc.last_seen;
                    } else {
                        state.npcs[name] = {
                            appearance: newNpc.appearance || '',
                            personality: newNpc.personality || '',
                            relationship: newNpc.relationship || '',
                            gender: newNpc.gender || '',
                            age: newNpc.age || '',
                            race: newNpc.race || '',
                            job: newNpc.job || '',
                            birthday: newNpc.birthday || '',
                            note: newNpc.note || '',
                            _ageRefDate: newNpc.age ? (state.timestamp.story_date || '') : '',
                            first_seen: newNpc.first_seen || new Date().toISOString(),
                            last_seen: newNpc.last_seen || new Date().toISOString()
                        };
                    }
                }
            }
            // 情緒狀態（覆蓋式）
            if (meta.mood) {
                for (const [charName, emotion] of Object.entries(meta.mood)) {
                    state.mood[charName] = emotion;
                }
            }
        }
        
        // 過濾用戶已刪除的NPC（防還原）
        const deletedNpcs = chat[0]?.horae_meta?._deletedNpcs;
        if (deletedNpcs?.length) {
            for (const name of deletedNpcs) {
                delete state.npcs[name];
                delete state.affection[name];
                delete state.costumes[name];
                delete state.mood[name];
                if (state.scene.characters_present) {
                    state.scene.characters_present = state.scene.characters_present.filter(c => c !== name);
                }
            }
        }
        
        // 為無ID物品分配ID
        let maxId = 0;
        for (const info of Object.values(state.items)) {
            if (info._id) {
                const num = parseInt(info._id, 10);
                if (num > maxId) maxId = num;
            }
        }
        for (const info of Object.values(state.items)) {
            if (!info._id) {
                maxId++;
                info._id = padItemId(maxId);
            }
        }
        
        // 為無ID的NPC分配ID
        let maxNpcId = 0;
        for (const info of Object.values(state.npcs)) {
            if (info._id) {
                const num = parseInt(info._id, 10);
                if (num > maxNpcId) maxNpcId = num;
            }
        }
        for (const info of Object.values(state.npcs)) {
            if (!info._id) {
                maxNpcId++;
                info._id = padItemId(maxNpcId);
            }
        }
        
        return state;
    }

    /** 解析生日字元串，支援 yyyy-mm-dd / yyyy/mm/dd / mm-dd / mm/dd */
    _parseBirthday(str) {
        if (!str) return null;
        let m = str.match(/(\d{2,4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
        if (m) return { year: parseInt(m[1]), month: parseInt(m[2]), day: parseInt(m[3]) };
        m = str.match(/^(\d{1,2})[\/\-.](\d{1,2})$/);
        if (m) return { year: null, month: parseInt(m[1]), day: parseInt(m[2]) };
        return null;
    }

    /** 根據劇情時間推移計算NPC目前年齡（優先使用生日精確計算） */
    calcCurrentAge(npcInfo, currentStoryDate) {
        const original = npcInfo.age || '';
        if (!original || !currentStoryDate) {
            return { display: original, original, changed: false };
        }

        const ageNum = parseInt(original);
        if (isNaN(ageNum)) {
            return { display: original, original, changed: false };
        }

        const curParsed = parseStoryDate(currentStoryDate);
        if (!curParsed || curParsed.type !== 'standard' || !curParsed.year) {
            return { display: original, original, changed: false };
        }

        const bdParsed = this._parseBirthday(npcInfo.birthday);

        // ── 有完整生日(含年份)：精確計算 ──
        if (bdParsed?.year) {
            let age = curParsed.year - bdParsed.year;
            if (bdParsed.month && curParsed.month) {
                if (curParsed.month < bdParsed.month ||
                    (curParsed.month === bdParsed.month && (curParsed.day || 1) < (bdParsed.day || 1))) {
                    age -= 1;
                }
            }
            age = Math.max(0, age);
            return { display: String(age), original, changed: age !== ageNum };
        }

        // 以下兩種情況都需要 _ageRefDate
        const refDate = npcInfo._ageRefDate || '';
        if (!refDate) return { display: original, original, changed: false };

        const refParsed = parseStoryDate(refDate);
        if (!refParsed || refParsed.type !== 'standard' || !refParsed.year) {
            return { display: original, original, changed: false };
        }

        // ── 僅有月日生日：用 refDate+age 推算出生年，再精確計算 ──
        if (bdParsed?.month) {
            let birthYear = refParsed.year - ageNum;
            if (refParsed.month) {
                const refBeforeBd = refParsed.month < bdParsed.month ||
                    (refParsed.month === bdParsed.month && (refParsed.day || 1) < (bdParsed.day || 1));
                if (refBeforeBd) birthYear -= 1;
            }
            let currentAge = curParsed.year - birthYear;
            if (curParsed.month) {
                const curBeforeBd = curParsed.month < bdParsed.month ||
                    (curParsed.month === bdParsed.month && (curParsed.day || 1) < (bdParsed.day || 1));
                if (curBeforeBd) currentAge -= 1;
            }
            if (currentAge <= ageNum) return { display: original, original, changed: false };
            return { display: String(currentAge), original, changed: true };
        }

        // ── 無生日：退回舊邏輯 ──
        let yearDiff = curParsed.year - refParsed.year;
        if (refParsed.month && curParsed.month) {
            if (curParsed.month < refParsed.month ||
                (curParsed.month === refParsed.month && (curParsed.day || 1) < (refParsed.day || 1))) {
                yearDiff -= 1;
            }
        }
        if (yearDiff <= 0) return { display: original, original, changed: false };
        return { display: String(ageNum + yearDiff), original, changed: true };
    }

    /** 透過ID搜尋物品 */
    findItemById(items, id) {
        const normalizedId = id.replace(/^#/, '').trim();
        for (const [name, info] of Object.entries(items)) {
            if (info._id === normalizedId || info._id === padItemId(parseInt(normalizedId, 10))) {
                return [name, info];
            }
        }
        return null;
    }

    /** 獲取事件列表（limit=0表示不限制數量） */
    getEvents(limit = 0, filterLevel = 'all', skipLast = 0) {
        const chat = this.getChat();
        const end = Math.max(0, chat.length - skipLast);
        const events = [];
        
        for (let i = 0; i < end; i++) {
            const meta = chat[i].horae_meta;
            if (meta?._skipHorae) continue;
            
            const metaEvents = meta?.events || (meta?.event ? [meta.event] : []);
            
            for (let j = 0; j < metaEvents.length; j++) {
                const evt = metaEvents[j];
                if (!evt?.summary) continue;
                
                if (filterLevel !== 'all' && evt.level !== filterLevel) {
                    continue;
                }
                
                events.push({
                    messageIndex: i,
                    eventIndex: j,
                    timestamp: meta.timestamp,
                    event: evt
                });
                
                if (limit > 0 && events.length >= limit) break;
            }
            if (limit > 0 && events.length >= limit) break;
        }
        
        return events;
    }

    /** 獲取重要事件列表（相容舊呼叫） */
    getImportantEvents(limit = 0) {
        return this.getEvents(limit, 'all');
    }

    /** 生成緊湊的上下文注入內容（skipLast: swipe時跳過末尾N條訊息） */
    generateCompactPrompt(skipLast = 0) {
        const state = this.getLatestState(skipLast);
        const lines = [];
        
        // 狀態快照頭
        lines.push('[目前狀態快照——對比本回合劇情，僅在<horae>中輸出發生實質變化的資料欄]');
        
        const sendTimeline = this.settings?.sendTimeline !== false;
        const sendCharacters = this.settings?.sendCharacters !== false;
        const sendItems = this.settings?.sendItems !== false;
        
        // 時間
        if (state.timestamp.story_date) {
            const fullDateTime = formatFullDateTime(state.timestamp.story_date, state.timestamp.story_time);
            lines.push(`[時間|${fullDateTime}]`);
            
            // 時間參考
            if (sendTimeline) {
                const timeRef = generateTimeReference(state.timestamp.story_date);
                if (timeRef && timeRef.type === 'standard') {
                    // 標準日曆
                    lines.push(`[時間參考|昨天=${timeRef.yesterday}|前天=${timeRef.dayBefore}|3天前=${timeRef.threeDaysAgo}]`);
                } else if (timeRef && timeRef.type === 'fantasy') {
                    // 奇幻日曆
                    lines.push(`[時間參考|奇幻日曆模式，參見劇情軌跡中的相對時間標記]`);
                }
            }
        }
        
        // 場景
        if (state.scene.location) {
            let sceneStr = `[場景|${state.scene.location}`;
            if (state.scene.atmosphere) {
                sceneStr += `|${state.scene.atmosphere}`;
            }
            sceneStr += ']';
            lines.push(sceneStr);

            if (this.settings?.sendLocationMemory) {
                const locMem = this.getLocationMemory();
                const loc = state.scene.location;
                const entry = this._findLocationMemory(loc, locMem, state._previousLocation);
                if (entry?.desc) {
                    lines.push(`[場景記憶|${entry.desc}]`);
                }
                // 附帶父級地點描述（如「酒館·大廳」→ 同時發送「酒館」的描述）
                const sepMatch = loc.match(/[·・\-\/\|]/);
                if (sepMatch) {
                    const parent = loc.substring(0, sepMatch.index).trim();
                    if (parent && locMem[parent] && locMem[parent].desc && parent !== entry?._matchedName) {
                        lines.push(`[場景記憶:${parent}|${locMem[parent].desc}]`);
                    }
                }
            }
        }
        
        // 在場角色和服裝
        if (sendCharacters) {
            const presentChars = state.scene.characters_present || [];
            
            if (presentChars.length > 0) {
                const charStrs = [];
                for (const char of presentChars) {
                    // 模糊配對服裝
                    const costumeKey = Object.keys(state.costumes || {}).find(
                        k => k === char || k.includes(char) || char.includes(k)
                    );
                    if (costumeKey && state.costumes[costumeKey]) {
                        charStrs.push(`${char}(${state.costumes[costumeKey]})`);
                    } else {
                        charStrs.push(char);
                    }
                }
                lines.push(`[在場|${charStrs.join('|')}]`);
            }
            
            // 情緒狀態（僅在場角色，變化驅動）
            if (this.settings?.sendMood) {
                const moodEntries = [];
                for (const char of presentChars) {
                    if (state.mood[char]) {
                        moodEntries.push(`${char}:${state.mood[char]}`);
                    }
                }
                if (moodEntries.length > 0) {
                    lines.push(`[情緒|${moodEntries.join('|')}]`);
                }
            }
            
            // 關係網路（僅在場角色相關的關係，從 chat[0] 讀取，零AI輸出token）
            if (this.settings?.sendRelationships) {
                const rels = this.getRelationshipsForCharacters(presentChars);
                if (rels.length > 0) {
                    lines.push('\n[關係網路]');
                    for (const r of rels) {
                        const noteStr = r.note ? `(${r.note})` : '';
                        lines.push(`${r.from}→${r.to}: ${r.type}${noteStr}`);
                    }
                }
            }
        }
        
        // 物品（已裝備的物品不在此處顯示，避免重複）
        if (sendItems) {
            const items = Object.entries(state.items);
            // 收集已裝備物品名集合
            const equippedNames = new Set();
            if (this.settings?.rpgMode && !!this.settings.sendRpgEquipment) {
                const rpgData = this.getRpgStateAt(skipLast);
                for (const [, slots] of Object.entries(rpgData.equipment || {})) {
                    for (const [, eqItems] of Object.entries(slots)) {
                        for (const eq of eqItems) equippedNames.add(eq.name);
                    }
                }
            }
            const unequipped = items.filter(([name]) => !equippedNames.has(name));
            if (unequipped.length > 0) {
                lines.push('\n[物品清單]');
                for (const [name, info] of unequipped) {
                    const id = info._id || '???';
                    const icon = info.icon || '';
                    const imp = info.importance === '!!' ? '關鍵' : info.importance === '!' ? '重要' : '';
                    const desc = info.description ? ` | ${info.description}` : '';
                    const holder = info.holder || '';
                    const loc = info.location ? `@${info.location}` : '';
                    const impTag = imp ? `[${imp}]` : '';
                    lines.push(`#${id} ${icon}${name}${impTag}${desc} = ${holder}${loc}`);
                }
            } else {
                lines.push('\n[物品清單] (空)');
            }
        }
        
        // 好感度
        if (sendCharacters) {
            const affections = Object.entries(state.affection).filter(([_, v]) => v !== 0);
            if (affections.length > 0) {
                const affStr = affections.map(([k, v]) => `${k}:${v > 0 ? '+' : ''}${v}`).join('|');
                lines.push(`[好感|${affStr}]`);
            }
            
            // NPC資訊
            const npcs = Object.entries(state.npcs);
            if (npcs.length > 0) {
                lines.push('\n[已知NPC]');
                for (const [name, info] of npcs) {
                    const id = info._id || '?';
                    const app = info.appearance || '';
                    const per = info.personality || '';
                    const rel = info.relationship || '';
                    // 主體：N編號 名｜外貌=個性@關係
                    let npcStr = `N${id} ${name}`;
                    if (app || per || rel) {
                        npcStr += `｜${app}=${per}@${rel}`;
                    }
                    // 擴展資料欄
                    const extras = [];
                    if (info._aliases?.length) extras.push(`曾用名:${info._aliases.join('/')}`);
                    if (info.gender) extras.push(`性別:${info.gender}`);
                    if (info.age) {
                        const ageResult = this.calcCurrentAge(info, state.timestamp.story_date);
                        extras.push(`年齡:${ageResult.display}`);
                    }
                    if (info.race) extras.push(`種族:${info.race}`);
                    if (info.job) extras.push(`職業:${info.job}`);
                    if (info.birthday) extras.push(`生日:${info.birthday}`);
                    if (info.note) extras.push(`補充:${info.note}`);
                    if (extras.length > 0) npcStr += `~${extras.join('~')}`;
                    lines.push(npcStr);
                }
            }
        }
        
        // 待辦事項
        const chatForAgenda = this.getChat();
        const allAgendaItems = [];
        const seenTexts = new Set();
        const deletedTexts = new Set(chatForAgenda?.[0]?.horae_meta?._deletedAgendaTexts || []);
        const userAgenda = chatForAgenda?.[0]?.horae_meta?.agenda || [];
        for (const item of userAgenda) {
            if (item._deleted || deletedTexts.has(item.text)) continue;
            if (!seenTexts.has(item.text)) {
                allAgendaItems.push(item);
                seenTexts.add(item.text);
            }
        }
        // AI寫入的（swipe時跳過末尾訊息）
        const agendaEnd = Math.max(0, (chatForAgenda?.length || 0) - skipLast);
        if (chatForAgenda) {
            for (let i = 1; i < agendaEnd; i++) {
                const msgAgenda = chatForAgenda[i].horae_meta?.agenda;
                if (msgAgenda?.length > 0) {
                    for (const item of msgAgenda) {
                        if (item._deleted || deletedTexts.has(item.text)) continue;
                        if (!seenTexts.has(item.text)) {
                            allAgendaItems.push(item);
                            seenTexts.add(item.text);
                        }
                    }
                }
            }
        }
        const activeAgenda = allAgendaItems.filter(a => !a.done);
        if (activeAgenda.length > 0) {
            lines.push('\n[待辦事項]');
            for (const item of activeAgenda) {
                const datePrefix = item.date ? `${item.date} ` : '';
                lines.push(`· ${datePrefix}${item.text}`);
            }
        }
        
        // RPG 狀態（僅啟用時注入，按在場角色過濾）
        if (this.settings?.rpgMode) {
            const rpg = this.getRpgStateAt(skipLast);
            const sendBars = this.settings?.sendRpgBars !== false;
            const sendSkills = this.settings?.sendRpgSkills !== false;

            // 屬性條名稱映射
            const _barCfg = this.settings?.rpgBarConfig || [];
            const _barNames = {};
            for (const b of _barCfg) _barNames[b.key] = b.name;

            // 按在場角色過濾 RPG 數據（無場景數據時發送全部）
            const presentChars = state.scene.characters_present || [];
            const userName = this.context?.name1 || '';
            const _cUoB = !!this.settings?.rpgBarsUserOnly;
            const _cUoS = !!this.settings?.rpgSkillsUserOnly;
            const _cUoA = !!this.settings?.rpgAttrsUserOnly;
            const _cUoE = !!this.settings?.rpgEquipmentUserOnly;
            const _cUoR = !!this.settings?.rpgReputationUserOnly;
            const _cUoL = !!this.settings?.rpgLevelUserOnly;
            const _cUoC = !!this.settings?.rpgCurrencyUserOnly;
            const allRpgNames = new Set([
                ...Object.keys(rpg.bars), ...Object.keys(rpg.status || {}),
                ...Object.keys(rpg.skills), ...Object.keys(rpg.attributes || {}),
                ...Object.keys(rpg.reputation || {}), ...Object.keys(rpg.equipment || {}),
                ...Object.keys(rpg.levels || {}), ...Object.keys(rpg.xp || {}),
                ...Object.keys(rpg.currency || {}),
            ]);
            const rpgAllowed = new Set();
            if (presentChars.length > 0) {
                for (const p of presentChars) {
                    const n = p.trim();
                    if (!n) continue;
                    if (allRpgNames.has(n)) { rpgAllowed.add(n); continue; }
                    if (n === userName && allRpgNames.has(userName)) { rpgAllowed.add(userName); continue; }
                    for (const rn of allRpgNames) {
                        if (rn.includes(n) || n.includes(rn)) { rpgAllowed.add(rn); break; }
                    }
                }
            }
            const filterRpg = rpgAllowed.size > 0;
            // userOnly時構建行不帶角色名前綴
            const _ctxPre = (name, isUo) => {
                if (isUo) return '';
                const npc = state.npcs[name];
                return npc?._id ? `N${npc._id} ${name}: ` : `${name}: `;
            };

            if (sendBars && Object.keys(rpg.bars).length > 0) {
                lines.push('\n[RPG狀態]');
                for (const [name, bars] of Object.entries(rpg.bars)) {
                    if (_cUoB && name !== userName) continue;
                    if (filterRpg && !rpgAllowed.has(name)) continue;
                    const parts = [];
                    for (const [type, val] of Object.entries(bars)) {
                        const label = val[2] || _barNames[type] || type.toUpperCase();
                        parts.push(`${label} ${val[0]}/${val[1]}`);
                    }
                    const sts = rpg.status?.[name];
                    if (sts?.length > 0) parts.push(`狀態:${sts.join('/')}`);
                    if (parts.length > 0) lines.push(`${_ctxPre(name, _cUoB)}${parts.join(' | ')}`);
                }
                for (const [name, effects] of Object.entries(rpg.status || {})) {
                    if (rpg.bars[name] || effects.length === 0) continue;
                    if (_cUoB && name !== userName) continue;
                    if (filterRpg && !rpgAllowed.has(name)) continue;
                    lines.push(`${_ctxPre(name, _cUoB)}狀態:${effects.join('/')}`);
                }
            }

            if (sendSkills && Object.keys(rpg.skills).length > 0) {
                const hasAny = Object.entries(rpg.skills).some(([n, arr]) =>
                    arr?.length > 0 && (!_cUoS || n === userName) && (!filterRpg || rpgAllowed.has(n)));
                if (hasAny) {
                    lines.push('\n[技能列表]');
                    for (const [name, skills] of Object.entries(rpg.skills)) {
                        if (!skills?.length) continue;
                        if (_cUoS && name !== userName) continue;
                        if (filterRpg && !rpgAllowed.has(name)) continue;
                        if (!_cUoS) {
                            const npc = state.npcs[name];
                            const pre = npc?._id ? `N${npc._id} ` : '';
                            lines.push(`${pre}${name}:`);
                        }
                        for (const sk of skills) {
                            const lv = sk.level ? ` ${sk.level}` : '';
                            const desc = sk.desc ? ` | ${sk.desc}` : '';
                            lines.push(`  ${sk.name}${lv}${desc}`);
                        }
                    }
                }
            }

            const sendAttrs = this.settings?.sendRpgAttributes !== false;
            const attrCfg = this.settings?.rpgAttributeConfig || [];
            if (sendAttrs && attrCfg.length > 0 && Object.keys(rpg.attributes || {}).length > 0) {
                lines.push('\n[多維屬性]');
                for (const [name, vals] of Object.entries(rpg.attributes)) {
                    if (_cUoA && name !== userName) continue;
                    if (filterRpg && !rpgAllowed.has(name)) continue;
                    const parts = attrCfg.map(a => `${a.name}${vals[a.key] ?? '?'}`);
                    lines.push(`${_ctxPre(name, _cUoA)}${parts.join(' | ')}`);
                }
            }

            // 裝備（按角色獨立格位，包含完整物品描述以節省 token）
            const sendEq = !!this.settings?.sendRpgEquipment;
            const eqPerChar = (rpg.equipmentConfig?.perChar) || {};
            const storedEq = this.getChat()?.[0]?.horae_meta?.rpg?.equipment || {};
            if (sendEq && Object.keys(rpg.equipment || {}).length > 0) {
                let hasEqData = false;
                for (const [name, slots] of Object.entries(rpg.equipment)) {
                    if (_cUoE && name !== userName) continue;
                    if (filterRpg && !rpgAllowed.has(name)) continue;
                    const ownerCfg = eqPerChar[name];
                    const validEqSlots = (ownerCfg && Array.isArray(ownerCfg.slots))
                        ? new Set(ownerCfg.slots.map(s => s.name)) : null;
                    const deletedEqSlots = ownerCfg ? new Set(ownerCfg._deletedSlots || []) : new Set();
                    const parts = [];
                    for (const [slotName, items] of Object.entries(slots)) {
                        if (deletedEqSlots.has(slotName)) continue;
                        if (validEqSlots && validEqSlots.size > 0 && !validEqSlots.has(slotName)) continue;
                        for (const item of items) {
                            const attrStr = Object.entries(item.attrs || {}).map(([k, v]) => `${k}${v >= 0 ? '+' : ''}${v}`).join(',');
                            const stored = storedEq[name]?.[slotName]?.find(e => e.name === item.name);
                            const desc = stored?._itemMeta?.description || '';
                            const descPart = desc ? ` "${desc}"` : '';
                            parts.push(`[${slotName}]${item.name}${attrStr ? `{${attrStr}}` : ''}${descPart}`);
                        }
                    }
                    if (parts.length > 0) {
                        if (!hasEqData) { lines.push('\n[裝備]'); hasEqData = true; }
                        lines.push(`${_ctxPre(name, _cUoE)}${parts.join(' | ')}`);
                    }
                }
            }

            // 聲望（需開關開啟）
            const sendRep = !!this.settings?.sendRpgReputation;
            const repConfig = rpg.reputationConfig || { categories: [] };
            if (sendRep && repConfig.categories.length > 0 && Object.keys(rpg.reputation || {}).length > 0) {
                const validRepNames = new Set(repConfig.categories.map(c => c.name));
                const deletedRepNames = new Set(repConfig._deletedCategories || []);
                let hasRepData = false;
                for (const [name, cats] of Object.entries(rpg.reputation)) {
                    if (_cUoR && name !== userName) continue;
                    if (filterRpg && !rpgAllowed.has(name)) continue;
                    const parts = [];
                    for (const [catName, data] of Object.entries(cats)) {
                        if (!validRepNames.has(catName) || deletedRepNames.has(catName)) continue;
                        parts.push(`${catName}:${data.value}`);
                    }
                    if (parts.length > 0) {
                        if (!hasRepData) { lines.push('\n[聲望]'); hasRepData = true; }
                        lines.push(`${_ctxPre(name, _cUoR)}${parts.join(' | ')}`);
                    }
                }
            }

            // 等級
            const sendLvl = !!this.settings?.sendRpgLevel;
            if (sendLvl && (Object.keys(rpg.levels || {}).length > 0 || Object.keys(rpg.xp || {}).length > 0)) {
                const allLvlNames = new Set([...Object.keys(rpg.levels || {}), ...Object.keys(rpg.xp || {})]);
                let hasLvlData = false;
                for (const name of allLvlNames) {
                    if (_cUoL && name !== userName) continue;
                    if (filterRpg && !rpgAllowed.has(name)) continue;
                    const lv = rpg.levels?.[name];
                    const xp = rpg.xp?.[name];
                    if (lv == null && !xp) continue;
                    if (!hasLvlData) { lines.push('\n[等級]'); hasLvlData = true; }
                    let lvStr = lv != null ? `Lv.${lv}` : '';
                    if (xp) lvStr += ` (經驗: ${xp[0]}/${xp[1]})`;
                    lines.push(`${_ctxPre(name, _cUoL)}${lvStr.trim()}`);
                }
            }

            // 貨幣
            const sendCur = !!this.settings?.sendRpgCurrency;
            const curConfig = rpg.currencyConfig || { denominations: [] };
            if (sendCur && curConfig.denominations.length > 0 && Object.keys(rpg.currency || {}).length > 0) {
                let hasCurData = false;
                for (const [name, coins] of Object.entries(rpg.currency)) {
                    if (_cUoC && name !== userName) continue;
                    if (filterRpg && !rpgAllowed.has(name)) continue;
                    const parts = [];
                    for (const d of curConfig.denominations) {
                        const val = coins[d.name];
                        if (val != null) parts.push(`${d.name}×${val}`);
                    }
                    if (parts.length > 0) {
                        if (!hasCurData) { lines.push('\n[貨幣]'); hasCurData = true; }
                        lines.push(`${_ctxPre(name, _cUoC)}${parts.join(', ')}`);
                    }
                }
            }

            // 據點
            if (!!this.settings?.sendRpgStronghold) {
                const shNodes = rpg.strongholds || [];
                if (shNodes.length > 0) {
                    lines.push('\n[據點]');
                    function _shTreeStr(nodes, parentId, indent) {
                        const children = nodes.filter(n => (n.parent || null) === parentId);
                        let str = '';
                        for (const c of children) {
                            const lvStr = c.level != null ? ` Lv.${c.level}` : '';
                            str += `${'  '.repeat(indent)}${c.name}${lvStr}`;
                            if (c.desc) str += ` — ${c.desc}`;
                            str += '\n';
                            str += _shTreeStr(nodes, c.id, indent + 1);
                        }
                        return str;
                    }
                    lines.push(_shTreeStr(shNodes, null, 0).trimEnd());
                }
            }
        }

        // 劇情軌跡
        if (sendTimeline) {
            const allEvents = this.getEvents(0, 'all', skipLast);
            // 過濾掉被活躍摘要覆蓋的原始事件（_compressedBy 且摘要為 active）
            const timelineChat = this.getChat();
            const autoSums = timelineChat?.[0]?.horae_meta?.autoSummaries || [];
            const activeSumIds = new Set(autoSums.filter(s => s.active).map(s => s.id));
            // 被活躍摘要壓縮的事件不發送；摘要為 inactive 時其 _summaryId 事件不發送
            const events = allEvents.filter(e => {
                if (e.event?._compressedBy && activeSumIds.has(e.event._compressedBy)) return false;
                if (e.event?._summaryId && !activeSumIds.has(e.event._summaryId)) return false;
                return true;
            });
            if (events.length > 0) {
                lines.push('\n[劇情軌跡]');
                
                const currentDate = state.timestamp?.story_date || '';
                
                const getLevelMark = (level) => {
                    if (level === '關鍵') return '★';
                    if (level === '重要') return '●';
                    return '○';
                };
                
                const getRelativeDesc = (eventDate) => {
                    if (!eventDate || !currentDate) return '';
                    const result = calculateDetailedRelativeTime(eventDate, currentDate);
                    if (result.days === null || result.days === undefined) return '';
                    
                    const { days, fromDate, toDate } = result;
                    
                    if (days === 0) return '(今天)';
                    if (days === 1) return '(昨天)';
                    if (days === 2) return '(前天)';
                    if (days === 3) return '(大前天)';
                    if (days === -1) return '(明天)';
                    if (days === -2) return '(後天)';
                    if (days === -3) return '(大後天)';
                    
                    if (days >= 4 && days <= 13 && fromDate) {
                        const WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'];
                        const weekday = fromDate.getDay();
                        return `(上週${WEEKDAY_NAMES[weekday]})`;
                    }
                    
                    if (days >= 20 && days < 60 && fromDate && toDate) {
                        const fromMonth = fromDate.getMonth();
                        const toMonth = toDate.getMonth();
                        if (fromMonth !== toMonth) {
                            return `(上個月${fromDate.getDate()}號)`;
                        }
                    }
                    
                    if (days >= 300 && fromDate && toDate) {
                        const fromYear = fromDate.getFullYear();
                        const toYear = toDate.getFullYear();
                        if (fromYear < toYear) {
                            const fromMonth = fromDate.getMonth() + 1;
                            return `(去年${fromMonth}月)`;
                        }
                    }
                    
                    if (days > 0 && days < 30) return `(${days}天前)`;
                    if (days > 0) return `(${Math.round(days / 30)}個月前)`;
                    if (days === -999 || days === -998 || days === -997) return '';
                    return '';
                };
                
                const sortedEvents = [...events].sort((a, b) => {
                    return (a.messageIndex || 0) - (b.messageIndex || 0);
                });
                
                const criticalAndImportant = sortedEvents.filter(e => 
                    e.event?.level === '關鍵' || e.event?.level === '重要' || e.event?.level === '摘要' || e.event?.isSummary
                );
                const contextDepth = this.settings?.contextDepth ?? 15;
                const normalAll = sortedEvents.filter(e => 
                    (e.event?.level === '一般' || !e.event?.level) && !e.event?.isSummary
                );
                const normalEvents = contextDepth === 0 ? [] : normalAll.slice(-contextDepth);
                
                const allToShow = [...criticalAndImportant, ...normalEvents]
                    .sort((a, b) => (a.messageIndex || 0) - (b.messageIndex || 0));
                
                // 預構建 summaryId→日期範圍 映射，讓摘要事件帶上時間跨度
                const _sumDateRanges = {};
                for (const s of autoSums) {
                    if (!s.active || !s.originalEvents?.length) continue;
                    const dates = s.originalEvents.map(oe => oe.timestamp?.story_date).filter(Boolean);
                    if (dates.length > 0) {
                        const first = dates[0], last = dates[dates.length - 1];
                        _sumDateRanges[s.id] = first === last ? first : `${first}~${last}`;
                    }
                }

                for (const e of allToShow) {
                    const isSummary = e.event?.isSummary || e.event?.level === '摘要';
                    if (isSummary) {
                        const dateRange = e.event?._summaryId ? _sumDateRanges[e.event._summaryId] : '';
                        const dateTag = dateRange ? `·${dateRange}` : '';
                        const relTag = dateRange ? getRelativeDesc(dateRange.split('~')[0]) : '';
                        lines.push(`📋 [摘要${dateTag}]${relTag}: ${e.event.summary}`);
                    } else {
                        const mark = getLevelMark(e.event?.level);
                        const date = e.timestamp?.story_date || '?';
                        const time = e.timestamp?.story_time || '';
                        const timeStr = time ? `${date} ${time}` : date;
                        const relativeDesc = getRelativeDesc(e.timestamp?.story_date);
                        const msgNum = e.messageIndex !== undefined ? `#${e.messageIndex}` : '';
                        lines.push(`${mark} ${msgNum} ${timeStr}${relativeDesc}: ${e.event.summary}`);
                    }
                }
            }
        }
        
        // 客製化表格數據（合併全域和本地）
        const chat = this.getChat();
        const firstMsg = chat?.[0];
        const localTables = firstMsg?.horae_meta?.customTables || [];
        const resolvedGlobal = this._getResolvedGlobalTables();
        const allTables = [...resolvedGlobal, ...localTables];
        for (const table of allTables) {
            const rows = table.rows || 2;
            const cols = table.cols || 2;
            const data = table.data || {};
            
            // 有內容或有填表說明才輸出
            const hasContent = Object.values(data).some(v => v && v.trim());
            const hasPrompt = table.prompt && table.prompt.trim();
            if (!hasContent && !hasPrompt) continue;
            
            const tableName = table.name || '客製化表格';
            lines.push(`\n[${tableName}](${rows - 1}行×${cols - 1}列)`);
            
            if (table.prompt && table.prompt.trim()) {
                lines.push(`(填寫要求: ${table.prompt.trim()})`);
            }
            
            // 檢測最後有內容的行（含行標題列）
            let lastDataRow = 0;
            for (let r = rows - 1; r >= 1; r--) {
                for (let c = 0; c < cols; c++) {
                    if (data[`${r}-${c}`] && data[`${r}-${c}`].trim()) {
                        lastDataRow = r;
                        break;
                    }
                }
                if (lastDataRow > 0) break;
            }
            if (lastDataRow === 0) lastDataRow = 1;
            
            const lockedRows = new Set(table.lockedRows || []);
            const lockedCols = new Set(table.lockedCols || []);
            const lockedCells = new Set(table.lockedCells || []);

            // 輸出表頭行（帶座標標註）
            const headerRow = [];
            for (let c = 0; c < cols; c++) {
                const label = data[`0-${c}`] || (c === 0 ? '表頭' : `列${c}`);
                const coord = `[0,${c}]`;
                headerRow.push(lockedCols.has(c) ? `${coord}${label}🔒` : `${coord}${label}`);
            }
            lines.push(headerRow.join(' | '));

            // 輸出數據行（帶座標標註）
            for (let r = 1; r <= lastDataRow; r++) {
                const rowData = [];
                for (let c = 0; c < cols; c++) {
                    const coord = `[${r},${c}]`;
                    if (c === 0) {
                        const label = data[`${r}-0`] || `${r}`;
                        rowData.push(lockedRows.has(r) ? `${coord}${label}🔒` : `${coord}${label}`);
                    } else {
                        const val = data[`${r}-${c}`] || '';
                        rowData.push(lockedCells.has(`${r}-${c}`) ? `${coord}${val}🔒` : `${coord}${val}`);
                    }
                }
                lines.push(rowData.join(' | '));
            }
            
            // 標註被省略的尾部空行
            if (lastDataRow < rows - 1) {
                lines.push(`(共${rows - 1}行，第${lastDataRow + 1}-${rows - 1}行暫無數據)`);
            }

            // 提示完全空的數據列
            const emptyCols = [];
            for (let c = 1; c < cols; c++) {
                let colHasData = false;
                for (let r = 1; r < rows; r++) {
                    if (data[`${r}-${c}`] && data[`${r}-${c}`].trim()) { colHasData = true; break; }
                }
                if (!colHasData) emptyCols.push(c);
            }
            if (emptyCols.length > 0) {
                const emptyColNames = emptyCols.map(c => data[`0-${c}`] || `列${c}`);
                lines.push(`(${emptyColNames.join('、')}：暫無數據，如劇情中已有相關資訊請填寫)`);
            }
        }
        
        return lines.join('\n');
    }

    /** 獲取好感度等級描述 */
    getAffectionLevel(value) {
        if (value >= 80) return '摯愛';
        if (value >= 60) return '親密';
        if (value >= 40) return '好感';
        if (value >= 20) return '友好';
        if (value >= 0) return '中立';
        if (value >= -20) return '冷淡';
        if (value >= -40) return '厭惡';
        if (value >= -60) return '敵視';
        return '仇恨';
    }

    /**
     * 根據用戶配置的標籤列表（逗號分隔），
     * 整段移除對應標籤及其內容（含可選屬性），
     * 防止小劇場等客製化區塊內的 horae 標籤汙染正文解析。
     */
    _stripCustomTags(text, tagList) {
        if (!text || !tagList) return text;
        const tags = tagList.split(/[,，\s]+/).map(t => t.trim()).filter(Boolean);
        for (const tag of tags) {
            const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            text = text.replace(new RegExp(`<${escaped}(?:\\s[^>]*)?>[\\s\\S]*?</${escaped}>`, 'gi'), '');
        }
        return text;
    }

    /** 解析AI回覆中的horae標籤 */
    parseHoraeTag(message) {
        if (!message) return null;
        
        // 提取所有 <horae> 塊並選擇包含有效資料欄的塊（防止其他插件生成的同名標籤干擾）
        let match = null;
        const allHoraeMatches = [...message.matchAll(/<horae>([\s\S]*?)<\/horae>/gi)];
        const horaeFieldPattern = /^(time|timestamp|location|atmosphere|scene_desc|characters|costume|item[!]*|item-|event|affection|npc|agenda|agenda-|rel|mood):/m;
        if (allHoraeMatches.length > 1) {
            match = allHoraeMatches.find(m => horaeFieldPattern.test(m[1])) || allHoraeMatches[0];
        } else if (allHoraeMatches.length === 1) {
            match = allHoraeMatches[0];
        }
        if (!match) {
            match = message.match(/<!--horae([\s\S]*?)-->/i);
        }
        
        const allEventMatches = [...message.matchAll(/<horaeevent>([\s\S]*?)<\/horaeevent>/gi)];
        const eventMatch = allEventMatches.length > 1
            ? (allEventMatches.find(m => /^event:/m.test(m[1])) || allEventMatches[0])
            : allEventMatches[0] || null;
        const tableMatches = [...message.matchAll(/<horaetable[:：]\s*(.+?)>([\s\S]*?)<\/horaetable>/gi)];
        const rpgMatches = [...message.matchAll(/<horaerpg>([\s\S]*?)<\/horaerpg>/gi)];
        
        if (!match && !eventMatch && tableMatches.length === 0 && rpgMatches.length === 0) return null;
        
        const content = match ? match[1].trim() : '';
        const eventContent = eventMatch ? eventMatch[1].trim() : '';
        const lines = content.split('\n').concat(eventContent.split('\n'));
        
        const result = {
            timestamp: {},
            costumes: {},
            items: {},
            deletedItems: [],
            events: [],
            affection: {},
            npcs: {},
            scene: {},
            agenda: [],
            deletedAgenda: [],
            mood: {},
            relationships: [],
        };
        
        for (const line of lines) {
            const trimmedLine = line.trim();
            if (!trimmedLine) continue;
            
            // time:10/1 15:00 或 time:小鎮歷永夜2931年 2月1日(五) 20:30
            if (trimmedLine.startsWith('time:')) {
                const timeStr = trimmedLine.substring(5).trim();
                // 從末尾分離 HH:MM 時鐘時間
                const clockMatch = timeStr.match(/\b(\d{1,2}:\d{2})\s*$/);
                if (clockMatch) {
                    result.timestamp.story_time = clockMatch[1];
                    result.timestamp.story_date = timeStr.substring(0, timeStr.lastIndexOf(clockMatch[1])).trim();
                } else {
                    // 無時鐘時間，整個字元串作為日期
                    result.timestamp.story_date = timeStr;
                    result.timestamp.story_time = '';
                }
            }
            // location:咖啡館二樓
            else if (trimmedLine.startsWith('location:')) {
                result.scene.location = trimmedLine.substring(9).trim();
            }
            // atmosphere:輕鬆
            else if (trimmedLine.startsWith('atmosphere:')) {
                result.scene.atmosphere = trimmedLine.substring(11).trim();
            }
            // scene_desc:地點的固定物理特徵描述（支援同一回覆多場景配對）
            else if (trimmedLine.startsWith('scene_desc:')) {
                const desc = trimmedLine.substring(11).trim();
                result.scene.scene_desc = desc;
                if (result.scene.location && desc) {
                    if (!result.scene._descPairs) result.scene._descPairs = [];
                    result.scene._descPairs.push({ location: result.scene.location, desc });
                }
            }
            // characters:愛麗絲,鮑勃
            else if (trimmedLine.startsWith('characters:')) {
                const chars = trimmedLine.substring(11).trim();
                result.scene.characters_present = chars.split(/[,，]/).map(c => c.trim()).filter(Boolean);
            }
            // costume:愛麗絲=白色連衣裙
            else if (trimmedLine.startsWith('costume:')) {
                const costumeStr = trimmedLine.substring(8).trim();
                const eqIndex = costumeStr.indexOf('=');
                if (eqIndex > 0) {
                    const char = costumeStr.substring(0, eqIndex).trim();
                    const costume = costumeStr.substring(eqIndex + 1).trim();
                    result.costumes[char] = costume;
                }
            }
            // item-:物品名 表示物品已消耗/刪除
            else if (trimmedLine.startsWith('item-:')) {
                const itemName = trimmedLine.substring(6).trim();
                const cleanName = itemName.replace(/^[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u, '').trim();
                if (cleanName) {
                    result.deletedItems.push(cleanName);
                }
            }
            // item:🍺劣質麥酒|描述=酒館@吧檯 / item!:📜重要物品|特殊功能描述=角色@位置 / item!!:💎關鍵物品=@位置
            else if (trimmedLine.startsWith('item!!:') || trimmedLine.startsWith('item!:') || trimmedLine.startsWith('item:')) {
                let importance = '';  // 一般用空字元串
                let itemStr;
                if (trimmedLine.startsWith('item!!:')) {
                    importance = '!!';  // 關鍵
                    itemStr = trimmedLine.substring(7).trim();
                } else if (trimmedLine.startsWith('item!:')) {
                    importance = '!';   // 重要
                    itemStr = trimmedLine.substring(6).trim();
                } else {
                    itemStr = trimmedLine.substring(5).trim();
                }
                
                const eqIndex = itemStr.indexOf('=');
                if (eqIndex > 0) {
                    let itemNamePart = itemStr.substring(0, eqIndex).trim();
                    const rest = itemStr.substring(eqIndex + 1).trim();
                    
                    let icon = null;
                    let itemName = itemNamePart;
                    let description = undefined;  // undefined = 合併時不覆蓋原有描述
                    
                    const emojiMatch = itemNamePart.match(/^([\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F600}-\u{1F64F}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA00}-\u{1FA6F}]|[\u{1FA70}-\u{1FAFF}]|[\u{231A}-\u{231B}]|[\u{23E9}-\u{23F3}]|[\u{23F8}-\u{23FA}]|[\u{25AA}-\u{25AB}]|[\u{25B6}]|[\u{25C0}]|[\u{25FB}-\u{25FE}]|[\u{2614}-\u{2615}]|[\u{2648}-\u{2653}]|[\u{267F}]|[\u{2693}]|[\u{26A1}]|[\u{26AA}-\u{26AB}]|[\u{26BD}-\u{26BE}]|[\u{26C4}-\u{26C5}]|[\u{26CE}]|[\u{26D4}]|[\u{26EA}]|[\u{26F2}-\u{26F3}]|[\u{26F5}]|[\u{26FA}]|[\u{26FD}]|[\u{2702}]|[\u{2705}]|[\u{2708}-\u{270D}]|[\u{270F}]|[\u{2712}]|[\u{2714}]|[\u{2716}]|[\u{271D}]|[\u{2721}]|[\u{2728}]|[\u{2733}-\u{2734}]|[\u{2744}]|[\u{2747}]|[\u{274C}]|[\u{274E}]|[\u{2753}-\u{2755}]|[\u{2757}]|[\u{2763}-\u{2764}]|[\u{2795}-\u{2797}]|[\u{27A1}]|[\u{27B0}]|[\u{27BF}]|[\u{2934}-\u{2935}]|[\u{2B05}-\u{2B07}]|[\u{2B1B}-\u{2B1C}]|[\u{2B50}]|[\u{2B55}]|[\u{3030}]|[\u{303D}]|[\u{3297}]|[\u{3299}])/u);
                    if (emojiMatch) {
                        icon = emojiMatch[1];
                        itemNamePart = itemNamePart.substring(icon.length).trim();
                    }
                    
                    const pipeIndex = itemNamePart.indexOf('|');
                    if (pipeIndex > 0) {
                        itemName = itemNamePart.substring(0, pipeIndex).trim();
                        const descText = itemNamePart.substring(pipeIndex + 1).trim();
                        if (descText) description = descText;
                    } else {
                        itemName = itemNamePart;
                    }
                    
                    // 去掉無意義的數量標記
                    itemName = itemName.replace(/[\(（]1[\)）]$/, '').trim();
                    itemName = itemName.replace(new RegExp(`[\\(（]1[${COUNTING_CLASSIFIERS}][\\)）]$`), '').trim();
                    itemName = itemName.replace(new RegExp(`[\\(（][${COUNTING_CLASSIFIERS}][\\)）]$`), '').trim();
                    
                    const atIndex = rest.indexOf('@');
                    const itemInfo = {
                        icon: icon,
                        importance: importance,
                        holder: atIndex >= 0 ? (rest.substring(0, atIndex).trim() || null) : (rest || null),
                        location: atIndex >= 0 ? (rest.substring(atIndex + 1).trim() || '') : ''
                    };
                    if (description !== undefined) itemInfo.description = description;
                    result.items[itemName] = itemInfo;
                }
            }
            // event:重要|愛麗絲坦白了秘密
            else if (trimmedLine.startsWith('event:')) {
                const eventStr = trimmedLine.substring(6).trim();
                const parts = eventStr.split('|');
                if (parts.length >= 2) {
                    const levelRaw = parts[0].trim();
                    const summary = parts.slice(1).join('|').trim();
                    
                    let level = '一般';
                    if (levelRaw === '關鍵' || levelRaw.toLowerCase() === 'critical') {
                        level = '關鍵';
                    } else if (levelRaw === '重要' || levelRaw.toLowerCase() === 'important') {
                        level = '重要';
                    }
                    
                    result.events.push({
                        is_important: level === '重要' || level === '關鍵',
                        level: level,
                        summary: summary
                    });
                }
            }
            // affection:鮑勃=65 或 affection:鮑勃+5（相容新舊格式）
            // 容忍AI附加註解如 affection:湯姆=18(+0)|觀察到xxx，只提取名字和數值
            else if (trimmedLine.startsWith('affection:')) {
                const affStr = trimmedLine.substring(10).trim();
                // 新格式：角色名=數值（絕對值，允許帶正負號如 =+28 或 =-15）
                const absoluteMatch = affStr.match(/^(.+?)=\s*([+\-]?\d+\.?\d*)/);
                if (absoluteMatch) {
                    const key = absoluteMatch[1].trim();
                    const value = parseFloat(absoluteMatch[2]);
                    result.affection[key] = { type: 'absolute', value: value };
                } else {
                    // 舊格式：角色名+/-數值（相對值，無=號）— 允許數值後跟任意註解
                    const relativeMatch = affStr.match(/^(.+?)([+\-]\d+\.?\d*)/);
                    if (relativeMatch) {
                        const key = relativeMatch[1].trim();
                        const value = relativeMatch[2];
                        result.affection[key] = { type: 'relative', value: value };
                    }
                }
            }
            // npc:名|外貌=個性@關係~性別:男~年齡:25~種族:人類~職業:傭兵~補充:xxx
            // 使用 ~ 分隔擴展資料欄（key:value），不依賴順序
            else if (trimmedLine.startsWith('npc:')) {
                const npcStr = trimmedLine.substring(4).trim();
                const npcInfo = this._parseNpcFields(npcStr);
                const name = npcInfo._name;
                delete npcInfo._name;
                
                if (name) {
                    npcInfo.last_seen = new Date().toISOString();
                    if (!result.npcs[name]) {
                        npcInfo.first_seen = new Date().toISOString();
                    }
                    result.npcs[name] = npcInfo;
                }
            }
            // agenda-:已完成待辦內容 / agenda:訂立日期|內容
            else if (trimmedLine.startsWith('agenda-:')) {
                const delStr = trimmedLine.substring(8).trim();
                if (delStr) {
                    const pipeIdx = delStr.indexOf('|');
                    const text = pipeIdx > 0 ? delStr.substring(pipeIdx + 1).trim() : delStr;
                    if (text) {
                        result.deletedAgenda.push(text);
                    }
                }
            }
            else if (trimmedLine.startsWith('agenda:')) {
                const agendaStr = trimmedLine.substring(7).trim();
                const pipeIdx = agendaStr.indexOf('|');
                let dateStr = '', text = '';
                if (pipeIdx > 0) {
                    dateStr = agendaStr.substring(0, pipeIdx).trim();
                    text = agendaStr.substring(pipeIdx + 1).trim();
                } else {
                    text = agendaStr;
                }
                if (text) {
                    // 檢測 AI 用括號標記完成的情況，自動歸入 deletedAgenda
                    const doneMatch = text.match(/[\(（](完成|已完成|done|finished|completed|失效|取消|已取消)[\)）]\s*$/i);
                    if (doneMatch) {
                        const cleanText = text.substring(0, text.length - doneMatch[0].length).trim();
                        if (cleanText) result.deletedAgenda.push(cleanText);
                    } else {
                        result.agenda.push({ date: dateStr, text, source: 'ai', done: false });
                    }
                }
            }
            // rel:角色A>角色B=關係類型|備註
            else if (trimmedLine.startsWith('rel:')) {
                const relStr = trimmedLine.substring(4).trim();
                const arrowIdx = relStr.indexOf('>');
                const eqIdx = relStr.indexOf('=');
                if (arrowIdx > 0 && eqIdx > arrowIdx) {
                    const from = relStr.substring(0, arrowIdx).trim();
                    const to = relStr.substring(arrowIdx + 1, eqIdx).trim();
                    const rest = relStr.substring(eqIdx + 1).trim();
                    const pipeIdx = rest.indexOf('|');
                    const type = pipeIdx > 0 ? rest.substring(0, pipeIdx).trim() : rest;
                    const note = pipeIdx > 0 ? rest.substring(pipeIdx + 1).trim() : '';
                    if (from && to && type) {
                        result.relationships.push({ from, to, type, note });
                    }
                }
            }
            // mood:角色名=情緒狀態
            else if (trimmedLine.startsWith('mood:')) {
                const moodStr = trimmedLine.substring(5).trim();
                const eqIdx = moodStr.indexOf('=');
                if (eqIdx > 0) {
                    const charName = moodStr.substring(0, eqIdx).trim();
                    const emotion = moodStr.substring(eqIdx + 1).trim();
                    if (charName && emotion) {
                        result.mood[charName] = emotion;
                    }
                }
            }
        }

        // 解析客製化表格數據
        if (tableMatches.length > 0) {
            result.tableUpdates = [];
            for (const tm of tableMatches) {
                const tableName = tm[1].trim();
                const tableContent = tm[2].trim();
                const updates = this._parseTableCellEntries(tableContent);
                
                if (Object.keys(updates).length > 0) {
                    result.tableUpdates.push({ name: tableName, updates });
                }
            }
        }

        // 解析 RPG 數據
        if (rpgMatches.length > 0) {
            result.rpg = { bars: {}, status: {}, skills: [], removedSkills: [], attributes: {}, reputation: {}, equipment: [], unequip: [], levels: {}, xp: {}, currency: [], baseChanges: [] };
            for (const rm of rpgMatches) {
                const rpgContent = rm[1].trim();
                for (const rpgLine of rpgContent.split('\n')) {
                    const trimmed = rpgLine.trim();
                    if (trimmed) this._parseRpgLine(trimmed, result.rpg);
                }
            }
        }

        return result;
    }

    /** 將解析結果合併到元數據 */
    mergeParsedToMeta(baseMeta, parsed) {
        const meta = baseMeta ? JSON.parse(JSON.stringify(baseMeta)) : createEmptyMeta();
        
        if (parsed.timestamp?.story_date) {
            meta.timestamp.story_date = parsed.timestamp.story_date;
        }
        if (parsed.timestamp?.story_time) {
            meta.timestamp.story_time = parsed.timestamp.story_time;
        }
        meta.timestamp.absolute = new Date().toISOString();
        
        if (parsed.scene?.location) {
            meta.scene.location = parsed.scene.location;
        }
        if (parsed.scene?.atmosphere) {
            meta.scene.atmosphere = parsed.scene.atmosphere;
        }
        if (parsed.scene?.scene_desc) {
            meta.scene.scene_desc = parsed.scene.scene_desc;
        }
        if (parsed.scene?.characters_present?.length > 0) {
            meta.scene.characters_present = parsed.scene.characters_present;
        }
        
        if (parsed.costumes) {
            Object.assign(meta.costumes, parsed.costumes);
        }
        
        if (parsed.items) {
            Object.assign(meta.items, parsed.items);
        }
        
        if (parsed.deletedItems && parsed.deletedItems.length > 0) {
            if (!meta.deletedItems) meta.deletedItems = [];
            meta.deletedItems = [...new Set([...meta.deletedItems, ...parsed.deletedItems])];
        }
        
        // 支援新格式（events數組）和舊格式（單個event）
        if (parsed.events && parsed.events.length > 0) {
            meta.events = parsed.events;
        } else if (parsed.event) {
            // 相容舊格式：轉換為數組
            meta.events = [parsed.event];
        }
        
        if (parsed.affection) {
            Object.assign(meta.affection, parsed.affection);
        }
        
        if (parsed.npcs) {
            Object.assign(meta.npcs, parsed.npcs);
        }
        
        // 追加AI寫入的待辦（跳過用戶已手動刪除的）
        if (parsed.agenda && parsed.agenda.length > 0) {
            if (!meta.agenda) meta.agenda = [];
            const chat0 = this.getChat()?.[0];
            const deletedSet = new Set(chat0?.horae_meta?._deletedAgendaTexts || []);
            for (const item of parsed.agenda) {
                if (deletedSet.has(item.text)) continue;
                const isDupe = meta.agenda.some(a => a.text === item.text);
                if (!isDupe) {
                    meta.agenda.push(item);
                }
            }
        }
        
        // 關係網路：存入目前訊息（後續由 processAIResponse 合併到 chat[0]）
        if (parsed.relationships && parsed.relationships.length > 0) {
            if (!meta.relationships) meta.relationships = [];
            meta.relationships = parsed.relationships;
        }
        
        // 情緒狀態
        if (parsed.mood && Object.keys(parsed.mood).length > 0) {
            if (!meta.mood) meta.mood = {};
            Object.assign(meta.mood, parsed.mood);
        }
        
        // tableUpdates 作為副屬性傳遞
        if (parsed.tableUpdates) {
            meta._tableUpdates = parsed.tableUpdates;
        }
        
        if (parsed.rpg) {
            meta._rpgChanges = parsed.rpg;
        }
        
        return meta;
    }

    /** 解析單行 RPG 數據 */
    _parseRpgLine(line, rpg) {
        const _uoName = this.context?.name1 || '主角';
        const _uoB = !!this.settings?.rpgBarsUserOnly;
        const _uoS = !!this.settings?.rpgSkillsUserOnly;
        const _uoA = !!this.settings?.rpgAttrsUserOnly;
        const _uoE = !!this.settings?.rpgEquipmentUserOnly;
        const _uoR = !!this.settings?.rpgReputationUserOnly;
        const _uoL = !!this.settings?.rpgLevelUserOnly;
        const _uoC = !!this.settings?.rpgCurrencyUserOnly;

        // 通用：檢測行是否為無owner的userOnly格式（首段含=即正常格式，否則可能是UO格式）
        // 屬性條: 正常 key:owner=cur/max 或 userOnly key:cur/max(顯示名)
        const barNormal = line.match(/^([a-zA-Z]\w*):(.+?)=(\d+)\s*\/\s*(\d+)(?:\((.+?)\))?$/i);
        const barUo = _uoB ? line.match(/^([a-zA-Z]\w*):(\d+)\s*\/\s*(\d+)(?:\((.+?)\))?$/i) : null;
        if (barNormal && !/^(status|skill)$/i.test(barNormal[1])) {
            const type = barNormal[1].toLowerCase();
            const owner = _uoB ? _uoName : barNormal[2].trim();
            const current = parseInt(barNormal[3]);
            const max = parseInt(barNormal[4]);
            const label = barNormal[5]?.trim() || null;
            if (!rpg.bars[owner]) rpg.bars[owner] = {};
            rpg.bars[owner][type] = label ? [current, max, label] : [current, max];
            return;
        }
        if (barUo && !/^(status|skill)$/i.test(barUo[1])) {
            const type = barUo[1].toLowerCase();
            const current = parseInt(barUo[2]);
            const max = parseInt(barUo[3]);
            const label = barUo[4]?.trim() || null;
            if (!rpg.bars[_uoName]) rpg.bars[_uoName] = {};
            rpg.bars[_uoName][type] = label ? [current, max, label] : [current, max];
            return;
        }
        // status
        if (line.startsWith('status:')) {
            const str = line.substring(7).trim();
            const eq = str.indexOf('=');
            if (_uoB && eq < 0) {
                rpg.status[_uoName] = (!str || /^(正常|无|none)$/i.test(str))
                    ? [] : str.split('/').map(s => s.trim()).filter(Boolean);
            } else if (eq > 0) {
                const owner = _uoB ? _uoName : str.substring(0, eq).trim();
                const val = str.substring(eq + 1).trim();
                rpg.status[owner] = (!val || /^(正常|无|none)$/i.test(val))
                    ? [] : val.split('/').map(s => s.trim()).filter(Boolean);
            }
            return;
        }
        // skill
        if (line.startsWith('skill:')) {
            const parts = line.substring(6).trim().split('|').map(s => s.trim());
            if (_uoS && parts.length >= 1) {
                rpg.skills.push({ owner: _uoName, name: parts[0], level: parts[1] || '', desc: parts[2] || '' });
            } else if (parts.length >= 2) {
                rpg.skills.push({ owner: parts[0], name: parts[1], level: parts[2] || '', desc: parts[3] || '' });
            }
            return;
        }
        // skill-
        if (line.startsWith('skill-:')) {
            const parts = line.substring(7).trim().split('|').map(s => s.trim());
            if (_uoS && parts.length >= 1) {
                rpg.removedSkills.push({ owner: _uoName, name: parts[0] });
            } else if (parts.length >= 2) {
                rpg.removedSkills.push({ owner: parts[0], name: parts[1] });
            }
            return;
        }
        // equip
        if (line.startsWith('equip:')) {
            const parts = line.substring(6).trim().split('|').map(s => s.trim());
            const minParts = _uoE ? 2 : 3;
            if (parts.length >= minParts) {
                const owner = _uoE ? _uoName : parts[0];
                const slot = _uoE ? parts[0] : parts[1];
                const itemName = _uoE ? parts[1] : parts[2];
                const attrPart = _uoE ? parts[2] : parts[3];
                const attrs = {};
                if (attrPart) {
                    for (const kv of attrPart.split(',')) {
                        const m = kv.trim().match(/^(.+?)=(-?\d+)$/);
                        if (m) attrs[m[1].trim()] = parseInt(m[2]);
                    }
                }
                if (!rpg.equipment) rpg.equipment = [];
                rpg.equipment.push({ owner, slot, name: itemName, attrs });
            }
            return;
        }
        // unequip
        if (line.startsWith('unequip:')) {
            const parts = line.substring(8).trim().split('|').map(s => s.trim());
            const minParts = _uoE ? 2 : 3;
            if (parts.length >= minParts) {
                if (!rpg.unequip) rpg.unequip = [];
                if (_uoE) {
                    rpg.unequip.push({ owner: _uoName, slot: parts[0], name: parts[1] });
                } else {
                    rpg.unequip.push({ owner: parts[0], slot: parts[1], name: parts[2] });
                }
            }
            return;
        }
        // rep
        if (line.startsWith('rep:')) {
            const parts = line.substring(4).trim().split('|').map(s => s.trim());
            if (_uoR && parts.length >= 1) {
                const kv = parts[0].match(/^(.+?)=(-?\d+)$/);
                if (kv) {
                    if (!rpg.reputation) rpg.reputation = {};
                    if (!rpg.reputation[_uoName]) rpg.reputation[_uoName] = {};
                    rpg.reputation[_uoName][kv[1].trim()] = parseInt(kv[2]);
                }
            } else if (parts.length >= 2) {
                const owner = parts[0];
                const kv = parts[1].match(/^(.+?)=(-?\d+)$/);
                if (kv) {
                    if (!rpg.reputation) rpg.reputation = {};
                    if (!rpg.reputation[owner]) rpg.reputation[owner] = {};
                    rpg.reputation[owner][kv[1].trim()] = parseInt(kv[2]);
                }
            }
            return;
        }
        // level
        if (line.startsWith('level:')) {
            const str = line.substring(6).trim();
            if (_uoL) {
                const val = parseInt(str);
                if (!isNaN(val)) {
                    if (!rpg.levels) rpg.levels = {};
                    rpg.levels[_uoName] = val;
                }
            } else {
                const eq = str.indexOf('=');
                if (eq > 0) {
                    const owner = str.substring(0, eq).trim();
                    const val = parseInt(str.substring(eq + 1).trim());
                    if (!isNaN(val)) {
                        if (!rpg.levels) rpg.levels = {};
                        rpg.levels[owner] = val;
                    }
                }
            }
            return;
        }
        // xp
        if (line.startsWith('xp:')) {
            const str = line.substring(3).trim();
            if (_uoL) {
                const m = str.match(/^(\d+)\s*\/\s*(\d+)$/);
                if (m) {
                    if (!rpg.xp) rpg.xp = {};
                    rpg.xp[_uoName] = [parseInt(m[1]), parseInt(m[2])];
                }
            } else {
                const eq = str.indexOf('=');
                if (eq > 0) {
                    const owner = str.substring(0, eq).trim();
                    const valStr = str.substring(eq + 1).trim();
                    const m = valStr.match(/^(\d+)\s*\/\s*(\d+)$/);
                    if (m) {
                        if (!rpg.xp) rpg.xp = {};
                        rpg.xp[owner] = [parseInt(m[1]), parseInt(m[2])];
                    }
                }
            }
            return;
        }
        // currency
        if (line.startsWith('currency:')) {
            const parts = line.substring(9).trim().split('|').map(s => s.trim());
            if (_uoC && parts.length >= 1) {
                const kvStr = parts.length >= 2 ? parts[1] : parts[0];
                const kv = kvStr.match(/^(.+?)=([+-]?\d+)$/);
                if (kv) {
                    if (!rpg.currency) rpg.currency = [];
                    const rawVal = kv[2];
                    const isDelta = rawVal.startsWith('+') || rawVal.startsWith('-');
                    rpg.currency.push({ owner: _uoName, name: kv[1].trim(), value: parseInt(rawVal), isDelta });
                }
            } else if (parts.length >= 2) {
                const owner = parts[0];
                const kv = parts[1].match(/^(.+?)=([+-]?\d+)$/);
                if (kv) {
                    if (!rpg.currency) rpg.currency = [];
                    const rawVal = kv[2];
                    const isDelta = rawVal.startsWith('+') || rawVal.startsWith('-');
                    rpg.currency.push({ owner, name: kv[1].trim(), value: parseInt(rawVal), isDelta });
                }
            }
            return;
        }
        // attr
        if (line.startsWith('attr:')) {
            const parts = line.substring(5).trim().split('|').map(s => s.trim());
            if (parts.length >= 1) {
                let owner, startIdx;
                if (_uoA) {
                    owner = _uoName;
                    startIdx = 0;
                } else {
                    owner = parts[0];
                    startIdx = 1;
                }
                const vals = {};
                for (let i = startIdx; i < parts.length; i++) {
                    const kv = parts[i].match(/^(\w+)=(\d+)$/);
                    if (kv) vals[kv[1].toLowerCase()] = parseInt(kv[2]);
                }
                if (Object.keys(vals).length) {
                    if (!rpg.attributes) rpg.attributes = {};
                    rpg.attributes[owner] = vals;
                }
            }
            return;
        }
        // base:據點路徑=等級 或 base:據點路徑|desc=描述
        // 路徑用 > 分隔層級，如 base:主角莊園>鍛造區>鍛造爐=2
        if (line.startsWith('base:')) {
            if (!rpg.baseChanges) rpg.baseChanges = [];
            const raw = line.substring(5).trim();
            const pipeIdx = raw.indexOf('|');
            if (pipeIdx >= 0) {
                const path = raw.substring(0, pipeIdx).trim();
                const rest = raw.substring(pipeIdx + 1).trim();
                const kv = rest.match(/^(desc|level)=(.+)$/);
                if (kv) {
                    rpg.baseChanges.push({ path, field: kv[1], value: kv[2].trim() });
                }
            } else {
                const eqIdx = raw.indexOf('=');
                if (eqIdx >= 0) {
                    const path = raw.substring(0, eqIdx).trim();
                    const val = raw.substring(eqIdx + 1).trim();
                    const numVal = parseInt(val);
                    if (!isNaN(numVal)) {
                        rpg.baseChanges.push({ path, field: 'level', value: numVal });
                    } else {
                        rpg.baseChanges.push({ path, field: 'desc', value: val });
                    }
                }
            }
        }
    }

    /** 透過 N編號 解析歸屬者的規範名稱 */
    _resolveRpgOwner(ownerStr) {
        const m = ownerStr.match(/^N(\d+)\s+(.+)$/);
        if (m) {
            const npcId = m[1];
            const padded = padItemId(parseInt(npcId, 10));
            const chat = this.getChat();
            for (let i = chat.length - 1; i >= 0; i--) {
                const npcs = chat[i]?.horae_meta?.npcs;
                if (!npcs) continue;
                for (const [name, info] of Object.entries(npcs)) {
                    if (String(info._id) === npcId || info._id === padded) return name;
                }
            }
            return m[2].trim();
        }
        return ownerStr.trim();
    }

    /** 合併 RPG 變更到 chat[0].horae_meta.rpg */
    _mergeRpgData(changes) {
        const chat = this.getChat();
        if (!chat?.length || !changes) return;
        const first = chat[0];
        if (!first.horae_meta) first.horae_meta = createEmptyMeta();
        if (!first.horae_meta.rpg) first.horae_meta.rpg = { bars: {}, status: {}, skills: {} };
        const rpg = first.horae_meta.rpg;

        const _mUN = this.context?.name1 || '';

        for (const [raw, barData] of Object.entries(changes.bars || {})) {
            const owner = this._resolveRpgOwner(raw);
            if (this.settings?.rpgBarsUserOnly && owner !== _mUN) continue;
            if (!rpg.bars[owner]) rpg.bars[owner] = {};
            Object.assign(rpg.bars[owner], barData);
        }
        for (const [raw, effects] of Object.entries(changes.status || {})) {
            const owner = this._resolveRpgOwner(raw);
            if (this.settings?.rpgBarsUserOnly && owner !== _mUN) continue;
            if (!rpg.status) rpg.status = {};
            rpg.status[owner] = effects;
        }
        for (const sk of (changes.skills || [])) {
            const owner = this._resolveRpgOwner(sk.owner);
            if (this.settings?.rpgSkillsUserOnly && owner !== _mUN) continue;
            if (!rpg.skills[owner]) rpg.skills[owner] = [];
            const idx = rpg.skills[owner].findIndex(s => s.name === sk.name);
            if (idx >= 0) {
                if (sk.level) rpg.skills[owner][idx].level = sk.level;
                if (sk.desc) rpg.skills[owner][idx].desc = sk.desc;
            } else {
                rpg.skills[owner].push({ name: sk.name, level: sk.level, desc: sk.desc });
            }
        }
        for (const sk of (changes.removedSkills || [])) {
            const owner = this._resolveRpgOwner(sk.owner);
            if (this.settings?.rpgSkillsUserOnly && owner !== _mUN) continue;
            if (rpg.skills[owner]) {
                rpg.skills[owner] = rpg.skills[owner].filter(s => s.name !== sk.name);
            }
        }
        // 多維屬性
        for (const [raw, vals] of Object.entries(changes.attributes || {})) {
            const owner = this._resolveRpgOwner(raw);
            if (this.settings?.rpgAttrsUserOnly && owner !== _mUN) continue;
            if (!rpg.attributes) rpg.attributes = {};
            rpg.attributes[owner] = { ...(rpg.attributes[owner] || {}), ...vals };
        }
        // 裝備：按角色獨立格位配置
        if (changes.equipment?.length > 0 || changes.unequip?.length > 0) {
            if (!rpg.equipmentConfig) rpg.equipmentConfig = { locked: false, perChar: {} };
            if (!rpg.equipmentConfig.perChar) rpg.equipmentConfig.perChar = {};
            if (!rpg.equipment) rpg.equipment = {};
            const _getOwnerSlots = (owner) => {
                const pc = rpg.equipmentConfig.perChar[owner];
                if (!pc || !Array.isArray(pc.slots)) return { valid: new Set(), deleted: new Set(), maxMap: {} };
                return {
                    valid: new Set(pc.slots.map(s => s.name)),
                    deleted: new Set(pc._deletedSlots || []),
                    maxMap: Object.fromEntries(pc.slots.map(s => [s.name, s.maxCount ?? 1])),
                };
            };
            const _findAndTakeItem = (name) => {
                const state = this.getLatestState();
                const itemInfo = state?.items?.[name];
                if (!itemInfo) return null;
                const meta = { icon: itemInfo.icon || '', description: itemInfo.description || '', importance: itemInfo.importance || '', _id: itemInfo._id || '', _locked: itemInfo._locked || false };
                for (let k = chat.length - 1; k >= 0; k--) {
                    if (chat[k]?.horae_meta?.items?.[name]) { delete chat[k].horae_meta.items[name]; break; }
                }
                return meta;
            };
            const _returnItemFromEquip = (entry, owner) => {
                if (!first.horae_meta.items) first.horae_meta.items = {};
                const m = entry._itemMeta || {};
                first.horae_meta.items[entry.name] = {
                    icon: m.icon || '📦', description: m.description || '', importance: m.importance || '',
                    holder: owner, location: '', _id: m._id || '', _locked: m._locked || false,
                };
            };
            for (const u of (changes.unequip || [])) {
                const owner = this._resolveRpgOwner(u.owner);
                if (this.settings?.rpgEquipmentUserOnly && owner !== _mUN) continue;
                if (!rpg.equipment[owner]?.[u.slot]) continue;
                const removed = rpg.equipment[owner][u.slot].find(e => e.name === u.name);
                rpg.equipment[owner][u.slot] = rpg.equipment[owner][u.slot].filter(e => e.name !== u.name);
                if (removed) _returnItemFromEquip(removed, owner);
                if (!rpg.equipment[owner][u.slot].length) delete rpg.equipment[owner][u.slot];
                if (rpg.equipment[owner] && !Object.keys(rpg.equipment[owner]).length) delete rpg.equipment[owner];
            }
            for (const eq of (changes.equipment || [])) {
                const slotName = eq.slot;
                const owner = this._resolveRpgOwner(eq.owner);
                if (this.settings?.rpgEquipmentUserOnly && owner !== _mUN) continue;
                const { valid, deleted, maxMap } = _getOwnerSlots(owner);
                if (valid.size > 0 && (!valid.has(slotName) || deleted.has(slotName))) continue;
                if (!rpg.equipment[owner]) rpg.equipment[owner] = {};
                if (!rpg.equipment[owner][slotName]) rpg.equipment[owner][slotName] = [];
                const existing = rpg.equipment[owner][slotName].findIndex(e => e.name === eq.name);
                if (existing >= 0) {
                    rpg.equipment[owner][slotName][existing].attrs = eq.attrs;
                } else {
                    const maxCount = maxMap[slotName] ?? 1;
                    if (rpg.equipment[owner][slotName].length >= maxCount) {
                        const bumped = rpg.equipment[owner][slotName].shift();
                        if (bumped) _returnItemFromEquip(bumped, owner);
                    }
                    const itemMeta = _findAndTakeItem(eq.name);
                    rpg.equipment[owner][slotName].push({ name: eq.name, attrs: eq.attrs || {}, ...(itemMeta ? { _itemMeta: itemMeta } : {}) });
                }
            }
        }
        // 聲望：只接受 reputationConfig 中已定義且未刪除的分類
        if (changes.reputation && Object.keys(changes.reputation).length > 0) {
            if (!rpg.reputationConfig) rpg.reputationConfig = { categories: [], _deletedCategories: [] };
            if (!rpg.reputation) rpg.reputation = {};
            const validNames = new Set((rpg.reputationConfig.categories || []).map(c => c.name));
            const deleted = new Set(rpg.reputationConfig._deletedCategories || []);
            for (const [raw, cats] of Object.entries(changes.reputation)) {
                const owner = this._resolveRpgOwner(raw);
                if (this.settings?.rpgReputationUserOnly && owner !== _mUN) continue;
                if (!rpg.reputation[owner]) rpg.reputation[owner] = {};
                for (const [catName, val] of Object.entries(cats)) {
                    if (!validNames.has(catName) || deleted.has(catName)) continue;
                    const cfg = rpg.reputationConfig.categories.find(c => c.name === catName);
                    const clamped = Math.max(cfg?.min ?? -100, Math.min(cfg?.max ?? 100, val));
                    if (!rpg.reputation[owner][catName]) {
                        rpg.reputation[owner][catName] = { value: clamped, subItems: {} };
                    } else {
                        rpg.reputation[owner][catName].value = clamped;
                    }
                }
            }
        }
        // 等級
        for (const [raw, val] of Object.entries(changes.levels || {})) {
            const owner = this._resolveRpgOwner(raw);
            if (this.settings?.rpgLevelUserOnly && owner !== _mUN) continue;
            if (!rpg.levels) rpg.levels = {};
            rpg.levels[owner] = val;
        }
        // 經驗值
        for (const [raw, val] of Object.entries(changes.xp || {})) {
            const owner = this._resolveRpgOwner(raw);
            if (this.settings?.rpgLevelUserOnly && owner !== _mUN) continue;
            if (!rpg.xp) rpg.xp = {};
            rpg.xp[owner] = val;
        }
        // 貨幣：只接受 currencyConfig 中已定義的幣種
        if (changes.currency?.length > 0) {
            if (!rpg.currencyConfig) rpg.currencyConfig = { denominations: [] };
            if (!rpg.currency) rpg.currency = {};
            const validDenoms = new Set((rpg.currencyConfig.denominations || []).map(d => d.name));
            for (const c of changes.currency) {
                const owner = this._resolveRpgOwner(c.owner);
                if (this.settings?.rpgCurrencyUserOnly && owner !== _mUN) continue;
                if (!validDenoms.has(c.name)) continue;
                if (!rpg.currency[owner]) rpg.currency[owner] = {};
                if (c.isDelta) {
                    rpg.currency[owner][c.name] = (rpg.currency[owner][c.name] || 0) + c.value;
                } else {
                    rpg.currency[owner][c.name] = c.value;
                }
            }
        }
        // 據點變更
        if (changes.baseChanges?.length > 0) {
            if (!rpg.strongholds) rpg.strongholds = [];
            for (const bc of changes.baseChanges) {
                const pathParts = bc.path.split('>').map(s => s.trim()).filter(Boolean);
                let parentId = null;
                let targetNode = null;
                for (const part of pathParts) {
                    targetNode = rpg.strongholds.find(n => n.name === part && (n.parent || null) === parentId);
                    if (!targetNode) {
                        targetNode = { id: 'sh_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name: part, level: null, desc: '', parent: parentId };
                        rpg.strongholds.push(targetNode);
                    }
                    parentId = targetNode.id;
                }
                if (targetNode) {
                    if (bc.field === 'level') targetNode.level = typeof bc.value === 'number' ? bc.value : parseInt(bc.value);
                    else if (bc.field === 'desc') targetNode.desc = String(bc.value);
                }
            }
        }
    }

    /** 從所有訊息重建 RPG 全域數據（保留用戶手動編輯） */
    rebuildRpgData() {
        const chat = this.getChat();
        if (!chat?.length) return;
        const first = chat[0];
        if (!first.horae_meta) first.horae_meta = createEmptyMeta();
        const old = first.horae_meta.rpg || {};
        // 保留用戶手動添加的技能
        const userSkills = {};
        for (const [owner, arr] of Object.entries(old.skills || {})) {
            const ua = (arr || []).filter(s => s._userAdded);
            if (ua.length) userSkills[owner] = ua;
        }
        // 保留用戶手動刪除記錄和手動填寫的屬性
        const deletedSkills = old._deletedSkills || [];
        const userAttrs = old.attributes || {};
        // 保留聲望配置和用戶設定的細項
        const oldRepConfig = old.reputationConfig || { categories: [], _deletedCategories: [] };
        const oldReputation = old.reputation ? JSON.parse(JSON.stringify(old.reputation)) : {};
        // 保留裝備配置
        const oldEquipConfig = old.equipmentConfig || { locked: false, perChar: {} };
        // 保留貨幣配置
        const oldCurrencyConfig = old.currencyConfig || { denominations: [] };

        first.horae_meta.rpg = {
            bars: {}, status: {}, skills: {}, attributes: { ...userAttrs }, _deletedSkills: deletedSkills,
            reputationConfig: oldRepConfig, reputation: {},
            equipmentConfig: oldEquipConfig, equipment: {},
            levels: {}, xp: {},
            currencyConfig: oldCurrencyConfig, currency: {},
        };
        for (let i = 1; i < chat.length; i++) {
            const changes = chat[i]?.horae_meta?._rpgChanges;
            if (changes) this._mergeRpgData(changes);
        }
        // 回填用戶手動添加的技能
        const rpg = first.horae_meta.rpg;
        for (const [owner, arr] of Object.entries(userSkills)) {
            if (!rpg.skills[owner]) rpg.skills[owner] = [];
            for (const sk of arr) {
                if (!rpg.skills[owner].some(s => s.name === sk.name)) rpg.skills[owner].push(sk);
            }
        }
        // 過濾用戶手動刪除的技能
        for (const del of deletedSkills) {
            if (rpg.skills[del.owner]) {
                rpg.skills[del.owner] = rpg.skills[del.owner].filter(s => s.name !== del.name);
                if (!rpg.skills[del.owner].length) delete rpg.skills[del.owner];
            }
        }
        // 回填用戶設定的聲望細項（AI只寫主數值，細項是純用戶數據）
        const deletedRepCats = new Set(rpg.reputationConfig?._deletedCategories || []);
        const validRepCats = new Set((rpg.reputationConfig?.categories || []).map(c => c.name));
        for (const [owner, cats] of Object.entries(oldReputation)) {
            if (!rpg.reputation[owner]) rpg.reputation[owner] = {};
            for (const [catName, data] of Object.entries(cats)) {
                if (deletedRepCats.has(catName) || !validRepCats.has(catName)) continue;
                if (!rpg.reputation[owner][catName]) {
                    rpg.reputation[owner][catName] = data;
                } else {
                    rpg.reputation[owner][catName].subItems = data.subItems || {};
                }
            }
        }
    }

    /** 獲取 RPG 全域數據（chat[0] 累積） */
    getRpgData() {
        return this.getChat()?.[0]?.horae_meta?.rpg || {
            bars: {}, status: {}, skills: {}, attributes: {},
            reputation: {}, reputationConfig: { categories: [], _deletedCategories: [] },
            equipment: {}, equipmentConfig: { locked: false, perChar: {} },
            levels: {}, xp: {},
            currency: {}, currencyConfig: { denominations: [] },
        };
    }

    /**
     * 構建到指定訊息位置的 RPG 快照（不修改 chat[0]）
     * @param {number} skipLast - 跳過末尾N條訊息（swipe時=1）
     */
    getRpgStateAt(skipLast = 0) {
        const chat = this.getChat();
        if (!chat?.length) return { bars: {}, status: {}, skills: {}, attributes: {}, reputation: {}, equipment: {}, levels: {}, xp: {}, currency: {} };
        const end = Math.max(1, chat.length - skipLast);
        const first = chat[0];
        const rpgMeta = first?.horae_meta?.rpg || {};
        const snapshot = {
            bars: {}, status: {}, skills: {}, attributes: {}, reputation: {}, equipment: {},
            levels: JSON.parse(JSON.stringify(rpgMeta.levels || {})),
            xp: JSON.parse(JSON.stringify(rpgMeta.xp || {})),
            currency: JSON.parse(JSON.stringify(rpgMeta.currency || {})),
        };

        // 用戶手動編輯的數據
        const userSkills = {};
        for (const [owner, arr] of Object.entries(rpgMeta.skills || {})) {
            const ua = (arr || []).filter(s => s._userAdded);
            if (ua.length) userSkills[owner] = ua;
        }
        const deletedSkills = rpgMeta._deletedSkills || [];
        const userAttrs = {};
        for (const [owner, vals] of Object.entries(rpgMeta.attributes || {})) {
            userAttrs[owner] = { ...vals };
        }

        // 裝備格位配置（提前獲取，用於循環內校驗 maxCount）
        const _eqCfg = rpgMeta.equipmentConfig || { locked: false, perChar: {} };
        const _eqPerChar = _eqCfg.perChar || {};

        // 從訊息中累積屬性（snapshot 是獨立對象，不汙染 chat[0]）
        const _resolve = (raw) => this._resolveRpgOwner(raw);
        for (let i = 1; i < end; i++) {
            const changes = chat[i]?.horae_meta?._rpgChanges;
            if (!changes) continue;
            for (const [raw, barData] of Object.entries(changes.bars || {})) {
                const owner = _resolve(raw);
                if (!snapshot.bars[owner]) snapshot.bars[owner] = {};
                Object.assign(snapshot.bars[owner], barData);
            }
            for (const [raw, effects] of Object.entries(changes.status || {})) {
                const owner = _resolve(raw);
                snapshot.status[owner] = effects;
            }
            for (const sk of (changes.skills || [])) {
                const owner = _resolve(sk.owner);
                if (!snapshot.skills[owner]) snapshot.skills[owner] = [];
                const idx = snapshot.skills[owner].findIndex(s => s.name === sk.name);
                if (idx >= 0) {
                    if (sk.level) snapshot.skills[owner][idx].level = sk.level;
                    if (sk.desc) snapshot.skills[owner][idx].desc = sk.desc;
                } else {
                    snapshot.skills[owner].push({ name: sk.name, level: sk.level, desc: sk.desc });
                }
            }
            for (const sk of (changes.removedSkills || [])) {
                const owner = _resolve(sk.owner);
                if (snapshot.skills[owner]) {
                    snapshot.skills[owner] = snapshot.skills[owner].filter(s => s.name !== sk.name);
                }
            }
            for (const [raw, vals] of Object.entries(changes.attributes || {})) {
                const owner = _resolve(raw);
                snapshot.attributes[owner] = { ...(snapshot.attributes[owner] || {}), ...vals };
            }
            for (const [raw, cats] of Object.entries(changes.reputation || {})) {
                const owner = _resolve(raw);
                if (!snapshot.reputation[owner]) snapshot.reputation[owner] = {};
                for (const [catName, val] of Object.entries(cats)) {
                    if (!snapshot.reputation[owner][catName]) {
                        snapshot.reputation[owner][catName] = { value: val, subItems: {} };
                    } else {
                        snapshot.reputation[owner][catName].value = val;
                    }
                }
            }
            // 裝備
            for (const u of (changes.unequip || [])) {
                const owner = _resolve(u.owner);
                if (!snapshot.equipment[owner]?.[u.slot]) continue;
                snapshot.equipment[owner][u.slot] = snapshot.equipment[owner][u.slot].filter(e => e.name !== u.name);
                if (!snapshot.equipment[owner][u.slot].length) delete snapshot.equipment[owner][u.slot];
                if (!Object.keys(snapshot.equipment[owner] || {}).length) delete snapshot.equipment[owner];
            }
            for (const eq of (changes.equipment || [])) {
                const owner = _resolve(eq.owner);
                const ownerCfg = _eqPerChar[owner];
                const maxCount = (ownerCfg && Array.isArray(ownerCfg.slots))
                    ? (ownerCfg.slots.find(s => s.name === eq.slot)?.maxCount ?? 1) : 1;
                if (!snapshot.equipment[owner]) snapshot.equipment[owner] = {};
                if (!snapshot.equipment[owner][eq.slot]) snapshot.equipment[owner][eq.slot] = [];
                const idx = snapshot.equipment[owner][eq.slot].findIndex(e => e.name === eq.name);
                if (idx >= 0) {
                    snapshot.equipment[owner][eq.slot][idx].attrs = eq.attrs;
                } else {
                    while (snapshot.equipment[owner][eq.slot].length >= maxCount) snapshot.equipment[owner][eq.slot].shift();
                    snapshot.equipment[owner][eq.slot].push({ name: eq.name, attrs: eq.attrs || {} });
                }
            }
            // 等級/經驗
            for (const [raw, val] of Object.entries(changes.levels || {})) {
                snapshot.levels[_resolve(raw)] = val;
            }
            for (const [raw, val] of Object.entries(changes.xp || {})) {
                snapshot.xp[_resolve(raw)] = val;
            }
            // 貨幣（過濾已刪除/未註冊的幣種）
            const validDenoms = new Set(
                (rpgMeta.currencyConfig?.denominations || []).map(d => d.name)
            );
            for (const c of (changes.currency || [])) {
                if (validDenoms.size && !validDenoms.has(c.name)) continue;
                const owner = _resolve(c.owner);
                if (!snapshot.currency[owner]) snapshot.currency[owner] = {};
                if (c.isDelta) {
                    snapshot.currency[owner][c.name] = (snapshot.currency[owner][c.name] || 0) + c.value;
                } else {
                    snapshot.currency[owner][c.name] = c.value;
                }
            }
        }

        // 合入用戶手動屬性（AI數據優先覆蓋）
        for (const [owner, vals] of Object.entries(userAttrs)) {
            if (!snapshot.attributes[owner]) snapshot.attributes[owner] = {};
            for (const [k, v] of Object.entries(vals)) {
                if (snapshot.attributes[owner][k] === undefined) snapshot.attributes[owner][k] = v;
            }
        }
        // 回填用戶手動技能
        for (const [owner, arr] of Object.entries(userSkills)) {
            if (!snapshot.skills[owner]) snapshot.skills[owner] = [];
            for (const sk of arr) {
                if (!snapshot.skills[owner].some(s => s.name === sk.name)) snapshot.skills[owner].push(sk);
            }
        }
        // 過濾用戶手動刪除
        for (const del of deletedSkills) {
            if (snapshot.skills[del.owner]) {
                snapshot.skills[del.owner] = snapshot.skills[del.owner].filter(s => s.name !== del.name);
                if (!snapshot.skills[del.owner].length) delete snapshot.skills[del.owner];
            }
        }
        // 聲望：合入用戶細項，過濾已刪除分類
        const repConfig = rpgMeta.reputationConfig || { categories: [], _deletedCategories: [] };
        const validRepNames = new Set((repConfig.categories || []).map(c => c.name));
        const deletedRepNames = new Set(repConfig._deletedCategories || []);
        const userRep = rpgMeta.reputation || {};
        for (const [owner, cats] of Object.entries(userRep)) {
            if (!snapshot.reputation[owner]) snapshot.reputation[owner] = {};
            for (const [catName, data] of Object.entries(cats)) {
                if (deletedRepNames.has(catName) || !validRepNames.has(catName)) continue;
                if (!snapshot.reputation[owner][catName]) {
                    snapshot.reputation[owner][catName] = { ...data };
                } else {
                    snapshot.reputation[owner][catName].subItems = data.subItems || {};
                }
            }
        }
        // 移除快照中已刪除的聲望分類
        for (const [owner, cats] of Object.entries(snapshot.reputation)) {
            for (const catName of Object.keys(cats)) {
                if (deletedRepNames.has(catName) || !validRepNames.has(catName)) {
                    delete cats[catName];
                }
            }
            if (!Object.keys(cats).length) delete snapshot.reputation[owner];
        }
        snapshot.reputationConfig = repConfig;
        // 裝備：按角色過濾已刪除格位
        for (const [owner, slots] of Object.entries(snapshot.equipment)) {
            const ownerCfg = _eqPerChar[owner];
            if (!ownerCfg || !Array.isArray(ownerCfg.slots)) continue;
            const validEqSlots = new Set(ownerCfg.slots.map(s => s.name));
            const deletedEqSlots = new Set(ownerCfg._deletedSlots || []);
            for (const slotName of Object.keys(slots)) {
                if (deletedEqSlots.has(slotName) || (validEqSlots.size > 0 && !validEqSlots.has(slotName))) {
                    delete slots[slotName];
                }
            }
            if (!Object.keys(slots).length) delete snapshot.equipment[owner];
        }
        snapshot.equipmentConfig = _eqCfg;
        // 貨幣配置
        snapshot.currencyConfig = rpgMeta.currencyConfig || { denominations: [] };
        return snapshot;
    }

    /** 合併關係數據到 chat[0].horae_meta */
    _mergeRelationships(newRels) {
        const chat = this.getChat();
        if (!chat?.length || !newRels?.length) return;
        const firstMsg = chat[0];
        if (!firstMsg.horae_meta) firstMsg.horae_meta = createEmptyMeta();
        if (!firstMsg.horae_meta.relationships) firstMsg.horae_meta.relationships = [];
        const existing = firstMsg.horae_meta.relationships;
        for (const rel of newRels) {
            const idx = existing.findIndex(r => r.from === rel.from && r.to === rel.to);
            if (idx >= 0) {
                if (existing[idx]._userEdited) continue;
                existing[idx].type = rel.type;
                if (rel.note) existing[idx].note = rel.note;
            } else {
                existing.push({ ...rel });
            }
        }
    }

    /** 從所有訊息重建 chat[0] 的關係網路（用於編輯/刪除後回推） */
    rebuildRelationships() {
        const chat = this.getChat();
        if (!chat?.length) return;
        const firstMsg = chat[0];
        if (!firstMsg.horae_meta) firstMsg.horae_meta = createEmptyMeta();
        // 保留用戶手動編輯的關係，其餘重建
        const userEdited = (firstMsg.horae_meta.relationships || []).filter(r => r._userEdited);
        firstMsg.horae_meta.relationships = [...userEdited];
        for (let i = 1; i < chat.length; i++) {
            const rels = chat[i]?.horae_meta?.relationships;
            if (rels?.length) this._mergeRelationships(rels);
        }
    }

    /** 從所有訊息重建 chat[0] 的場景記憶（用於編輯/刪除/重新生成後回推） */
    rebuildLocationMemory() {
        const chat = this.getChat();
        if (!chat?.length) return;
        const firstMsg = chat[0];
        if (!firstMsg.horae_meta) firstMsg.horae_meta = createEmptyMeta();
        const existing = firstMsg.horae_meta.locationMemory || {};
        const rebuilt = {};
        const deletedNames = new Set();
        // 保留用戶手動建立/編輯的條目，記錄已刪除的條目
        for (const [name, info] of Object.entries(existing)) {
            if (info._deleted) {
                deletedNames.add(name);
                rebuilt[name] = { ...info };
                continue;
            }
            if (info._userEdited) rebuilt[name] = { ...info };
        }
        // 從訊息重放 AI 寫入的 scene_desc（按時間順序，後覆蓋前），跳過已刪除/用戶編輯的
        for (let i = 1; i < chat.length; i++) {
            const meta = chat[i]?.horae_meta;
            const pairs = meta?.scene?._descPairs;
            if (pairs?.length > 0) {
                for (const p of pairs) {
                    if (deletedNames.has(p.location)) continue;
                    if (rebuilt[p.location]?._userEdited) continue;
                    rebuilt[p.location] = {
                        desc: p.desc,
                        firstSeen: rebuilt[p.location]?.firstSeen || new Date().toISOString(),
                        lastUpdated: new Date().toISOString()
                    };
                }
            } else if (meta?.scene?.scene_desc && meta?.scene?.location) {
                const loc = meta.scene.location;
                if (deletedNames.has(loc)) continue;
                if (rebuilt[loc]?._userEdited) continue;
                rebuilt[loc] = {
                    desc: meta.scene.scene_desc,
                    firstSeen: rebuilt[loc]?.firstSeen || new Date().toISOString(),
                    lastUpdated: new Date().toISOString()
                };
            }
        }
        firstMsg.horae_meta.locationMemory = rebuilt;
    }

    getRelationships() {
        const chat = this.getChat();
        return chat?.[0]?.horae_meta?.relationships || [];
    }

    /** 設定關係網路（用戶手動編輯時） */
    setRelationships(relationships) {
        const chat = this.getChat();
        if (!chat?.length) return;
        const firstMsg = chat[0];
        if (!firstMsg.horae_meta) firstMsg.horae_meta = createEmptyMeta();
        firstMsg.horae_meta.relationships = relationships;
    }

    /** 獲取指定角色相關的關係（無在場角色時返回空數組） */
    getRelationshipsForCharacters(charNames) {
        if (!charNames?.length) return [];
        const rels = this.getRelationships();
        const nameSet = new Set(charNames);
        return rels.filter(r => nameSet.has(r.from) || nameSet.has(r.to));
    }

    /** 全域刪除已完成的待辦事項 */
    removeCompletedAgenda(deletedTexts) {
        const chat = this.getChat();
        if (!chat || deletedTexts.length === 0) return;

        const isMatch = (agendaText, deleteText) => {
            if (!agendaText || !deleteText) return false;
            // 精確配對 或 互相包含（允許AI縮寫/擴寫）
            return agendaText === deleteText ||
                   agendaText.includes(deleteText) ||
                   deleteText.includes(agendaText);
        };

        if (chat[0]?.horae_meta?.agenda) {
            chat[0].horae_meta.agenda = chat[0].horae_meta.agenda.filter(
                a => !deletedTexts.some(dt => isMatch(a.text, dt))
            );
        }

        for (let i = 1; i < chat.length; i++) {
            if (chat[i]?.horae_meta?.agenda?.length > 0) {
                chat[i].horae_meta.agenda = chat[i].horae_meta.agenda.filter(
                    a => !deletedTexts.some(dt => isMatch(a.text, dt))
                );
            }
        }
    }

    /** 寫入/更新場景記憶到 chat[0] */
    _updateLocationMemory(locationName, desc) {
        const chat = this.getChat();
        if (!chat?.length || !locationName || !desc) return;
        const firstMsg = chat[0];
        if (!firstMsg.horae_meta) firstMsg.horae_meta = createEmptyMeta();
        if (!firstMsg.horae_meta.locationMemory) firstMsg.horae_meta.locationMemory = {};
        const mem = firstMsg.horae_meta.locationMemory;
        const now = new Date().toISOString();

        // 子級地點去重：若子級描述的"位於"部分重複了父級的地理資訊，則自動縮減
        const sepMatch = locationName.match(/[·・\-\/\|]/);
        if (sepMatch) {
            const parentName = locationName.substring(0, sepMatch.index).trim();
            const parentEntry = mem[parentName];
            if (parentEntry?.desc) {
                desc = this._deduplicateChildDesc(desc, parentEntry.desc, parentName);
            }
        }

        if (mem[locationName]) {
            if (mem[locationName]._userEdited || mem[locationName]._deleted) return;
            mem[locationName].desc = desc;
            mem[locationName].lastUpdated = now;
        } else {
            mem[locationName] = { desc, firstSeen: now, lastUpdated: now };
        }
        console.log(`[Horae] 場景記憶已更新: ${locationName}`);
    }

    /**
     * 子級描述去重：檢測子級描述是否包含父級的地理位置資訊，若包含則替換為相對位置
     */
    _deduplicateChildDesc(childDesc, parentDesc, parentName) {
        if (!childDesc || !parentDesc) return childDesc;
        // 提取父級的"位於"部分
        const parentLocMatch = parentDesc.match(/^位于(.+?)[。\.]/);
        if (!parentLocMatch) return childDesc;
        const parentLocInfo = parentLocMatch[1].trim();
        // 若子級描述也包含父級的地理位置關鍵詞（超過一半的字重合），則認為冗餘
        const parentKeywords = parentLocInfo.replace(/[，,、的]/g, ' ').split(/\s+/).filter(k => k.length >= 2);
        if (parentKeywords.length === 0) return childDesc;
        const childLocMatch = childDesc.match(/^位于(.+?)[。\.]/);
        if (!childLocMatch) return childDesc;
        const childLocInfo = childLocMatch[1].trim();
        let matchCount = 0;
        for (const kw of parentKeywords) {
            if (childLocInfo.includes(kw)) matchCount++;
        }
        // 超過一半關鍵詞重合，判定子級抄了父級地理位置
        if (matchCount >= Math.ceil(parentKeywords.length / 2)) {
            const shortName = parentName.length > 4 ? parentName.substring(0, 4) + '…' : parentName;
            const restDesc = childDesc.substring(childLocMatch[0].length).trim();
            return `位於${shortName}內。${restDesc}`;
        }
        return childDesc;
    }

    /** 獲取場景記憶 */
    getLocationMemory() {
        const chat = this.getChat();
        return chat?.[0]?.horae_meta?.locationMemory || {};
    }

    /**
     * 智慧配對場景記憶（複合地名支援）
     * 優先級：精確配對 → 拆分回退父級 → 上下文推斷 → 放棄
     */
    _findLocationMemory(currentLocation, locMem, previousLocation = '') {
        if (!currentLocation || !locMem || Object.keys(locMem).length === 0) return null;

        const tag = (name) => ({ ...locMem[name], _matchedName: name });

        if (locMem[currentLocation]) return tag(currentLocation);

        // 曾用名配對：檢查所有條目的 _aliases 數組
        for (const [name, info] of Object.entries(locMem)) {
            if (info._aliases?.includes(currentLocation)) return tag(name);
        }

        const SEP = /[·・\-\/|]/;
        const parts = currentLocation.split(SEP).map(s => s.trim()).filter(Boolean);

        if (parts.length > 1) {
            for (let i = parts.length - 1; i >= 1; i--) {
                const partial = parts.slice(0, i).join('·');
                if (locMem[partial]) return tag(partial);
                for (const [name, info] of Object.entries(locMem)) {
                    if (info._aliases?.includes(partial)) return tag(name);
                }
            }
        }

        if (previousLocation) {
            const prevParts = previousLocation.split(SEP).map(s => s.trim()).filter(Boolean);
            const prevParent = prevParts[0] || previousLocation;
            const curParent = parts[0] || currentLocation;

            if (prevParent !== curParent && prevParent.includes(curParent)) {
                if (locMem[prevParent]) return tag(prevParent);
            }
        }

        return null;
    }

    /**
     * 獲取全域表格的目前卡片數據（per-card overlay）
     * 全域表格的結構（表頭、名稱、提示詞、鎖定）共享，數據按角色卡分離
     */
    _getResolvedGlobalTables() {
        const templates = this.settings?.globalTables || [];
        const chat = this.getChat();
        if (!chat?.[0] || templates.length === 0) return [];

        const firstMsg = chat[0];
        if (!firstMsg.horae_meta) firstMsg.horae_meta = createEmptyMeta();
        if (!firstMsg.horae_meta.globalTableData) firstMsg.horae_meta.globalTableData = {};
        const perCardData = firstMsg.horae_meta.globalTableData;

        const result = [];
        for (const template of templates) {
            const name = (template.name || '').trim();
            if (!name) continue;

            if (!perCardData[name]) {
                // 首次在此卡使用：從模範初始化（含遷移舊數據）
                const initData = JSON.parse(JSON.stringify(template.data || {}));
                perCardData[name] = {
                    data: initData,
                    rows: template.rows || 2,
                    cols: template.cols || 2,
                    baseData: JSON.parse(JSON.stringify(initData)),
                    baseRows: template.rows || 2,
                    baseCols: template.cols || 2,
                };
            } else {
                // 同步全域模範的表頭到 per-card（用戶可能在別處改了表頭）
                const templateData = template.data || {};
                for (const key of Object.keys(templateData)) {
                    const [r, c] = key.split('-').map(Number);
                    if (r === 0 || c === 0) {
                        perCardData[name].data[key] = templateData[key];
                    }
                }
            }

            const overlay = perCardData[name];
            result.push({
                name: template.name,
                prompt: template.prompt,
                lockedRows: template.lockedRows || [],
                lockedCols: template.lockedCols || [],
                lockedCells: template.lockedCells || [],
                data: overlay.data,
                rows: overlay.rows,
                cols: overlay.cols,
                baseData: overlay.baseData,
                baseRows: overlay.baseRows,
                baseCols: overlay.baseCols,
            });
        }
        return result;
    }

    /** 處理AI回覆，解析標籤並存儲元數據 */
    processAIResponse(messageIndex, messageContent) {
        // 根據用戶配置的剔除標籤，整塊移除小劇場等客製化區塊，防止其內部的 horae 標籤汙染正文解析
        const cleanedContent = this._stripCustomTags(messageContent, this.settings?.vectorStripTags);
        let parsed = this.parseHoraeTag(cleanedContent);
        
        // 標籤解析失敗時，自動 fallback 到寬鬆格式解析
        if (!parsed) {
            parsed = this.parseLooseFormat(cleanedContent);
            if (parsed) {
                console.log(`[Horae] #${messageIndex} 未檢測到標籤，已透過寬鬆解析提取數據`);
            }
        }
        
        if (parsed) {
            const existingMeta = this.getMessageMeta(messageIndex);
            const newMeta = this.mergeParsedToMeta(existingMeta, parsed);
            
            // 處理表格更新
            if (newMeta._tableUpdates) {
                // 記錄表格貢獻，用於回退
                newMeta.tableContributions = newMeta._tableUpdates;
                this.applyTableUpdates(newMeta._tableUpdates);
                delete newMeta._tableUpdates;
            }
            
            // 處理AI標記已完成的待辦
            if (parsed.deletedAgenda && parsed.deletedAgenda.length > 0) {
                this.removeCompletedAgenda(parsed.deletedAgenda);
            }

            // 場景記憶：將 scene_desc 存入 locationMemory（支援同一回覆多場景配對）
            const descPairs = parsed.scene?._descPairs;
            if (descPairs?.length > 0) {
                for (const p of descPairs) {
                    this._updateLocationMemory(p.location, p.desc);
                }
            } else if (parsed.scene?.scene_desc && parsed.scene?.location) {
                this._updateLocationMemory(parsed.scene.location, parsed.scene.scene_desc);
            }
            
            // 關係網路：合併到 chat[0].horae_meta.relationships
            if (parsed.relationships && parsed.relationships.length > 0) {
                this._mergeRelationships(parsed.relationships);
            }
            
            this.setMessageMeta(messageIndex, newMeta);
            
            // RPG 數據：合併到 chat[0].horae_meta.rpg
            if (newMeta._rpgChanges) {
                this._mergeRpgData(newMeta._rpgChanges);
            }
            return true;
        } else {
            // 無標籤，建立空元數據
            if (!this.getMessageMeta(messageIndex)) {
                this.setMessageMeta(messageIndex, createEmptyMeta());
            }
            return false;
        }
    }

    /**
     * 解析NPC資料欄
     * 格式: 名|外貌=個性@關係~性別:男~年齡:25~種族:人類~職業:傭兵~補充:xxx
     */
    _parseNpcFields(npcStr) {
        const info = {};
        if (!npcStr) return { _name: '' };
        
        // 1. 分離擴展資料欄
        const tildeParts = npcStr.split('~');
        const mainPart = tildeParts[0].trim(); // 名|外貌=個性@關係
        
        for (let i = 1; i < tildeParts.length; i++) {
            const kv = tildeParts[i].trim();
            if (!kv) continue;
            const colonIdx = kv.indexOf(':');
            if (colonIdx <= 0) continue;
            const key = kv.substring(0, colonIdx).trim();
            const value = kv.substring(colonIdx + 1).trim();
            if (!value) continue;
            
            // 關鍵詞配對
            if (/^(性别|gender|sex)$/i.test(key)) info.gender = value;
            else if (/^(年龄|age|年纪)$/i.test(key)) info.age = value;
            else if (/^(种族|race|族裔|族群)$/i.test(key)) info.race = value;
            else if (/^(职业|job|class|职务|身份)$/i.test(key)) info.job = value;
            else if (/^(生日|birthday|birth)$/i.test(key)) info.birthday = value;
            else if (/^(补充|note|备注|其他)$/i.test(key)) info.note = value;
        }
        
        // 2. 解析主體
        let name = '';
        const pipeIdx = mainPart.indexOf('|');
        if (pipeIdx > 0) {
            name = mainPart.substring(0, pipeIdx).trim();
            const descPart = mainPart.substring(pipeIdx + 1).trim();
            
            const hasNewFormat = descPart.includes('=') || descPart.includes('@');
            
            if (hasNewFormat) {
                const atIdx = descPart.indexOf('@');
                let beforeAt = atIdx >= 0 ? descPart.substring(0, atIdx) : descPart;
                const relationship = atIdx >= 0 ? descPart.substring(atIdx + 1).trim() : '';
                
                const eqIdx = beforeAt.indexOf('=');
                const appearance = eqIdx >= 0 ? beforeAt.substring(0, eqIdx).trim() : beforeAt.trim();
                const personality = eqIdx >= 0 ? beforeAt.substring(eqIdx + 1).trim() : '';
                
                if (appearance) info.appearance = appearance;
                if (personality) info.personality = personality;
                if (relationship) info.relationship = relationship;
            } else {
                const parts = descPart.split('|').map(s => s.trim());
                if (parts[0]) info.appearance = parts[0];
                if (parts[1]) info.personality = parts[1];
                if (parts[2]) info.relationship = parts[2];
            }
        } else {
            name = mainPart.trim();
        }
        
        info._name = name;
        return info;
    }

    /**
     * 解析表格單元格數據
     * 格式: 每行一格 1,1:內容 或 單行多格用 | 分隔
     */
    _parseTableCellEntries(text) {
        const updates = {};
        if (!text) return updates;
        
        const cellRegex = /^(\d+)[,\-](\d+)[:：]\s*(.*)$/;
        
        for (const line of text.split('\n')) {
            const trimmed = line.trim();
            if (!trimmed) continue;
            
            // 按 | 分割
            const segments = trimmed.split(/\s*[|｜]\s*/);
            
            for (const seg of segments) {
                const s = seg.trim();
                if (!s) continue;
                
                const m = s.match(cellRegex);
                if (m) {
                    const r = parseInt(m[1]);
                    const c = parseInt(m[2]);
                    const value = m[3].trim();
                    // 過濾空標記
                    if (value && !/^[\(\（]?空[\)\）]?$/.test(value) && !/^[-—]+$/.test(value)) {
                        updates[`${r}-${c}`] = value;
                    }
                }
            }
        }
        
        return updates;
    }

    /** 將表格更新寫入 chat[0]（本地表格）或 per-card overlay（全域表格） */
    applyTableUpdates(tableUpdates) {
        if (!tableUpdates || tableUpdates.length === 0) return;

        const chat = this.getChat();
        if (!chat || chat.length === 0) return;

        const firstMsg = chat[0];
        if (!firstMsg.horae_meta) firstMsg.horae_meta = createEmptyMeta();
        if (!firstMsg.horae_meta.customTables) firstMsg.horae_meta.customTables = [];

        const localTables = firstMsg.horae_meta.customTables;
        const resolvedGlobal = this._getResolvedGlobalTables();

        for (const update of tableUpdates) {
            const updateName = (update.name || '').trim();
            let table = localTables.find(t => (t.name || '').trim() === updateName);
            let isGlobal = false;
            if (!table) {
                table = resolvedGlobal.find(t => (t.name || '').trim() === updateName);
                isGlobal = true;
            }
            if (!table) {
                console.warn(`[Horae] 表格 "${updateName}" 不存在，跳過`);
                continue;
            }

            if (!table.data) table.data = {};
            const lockedRows = new Set(table.lockedRows || []);
            const lockedCols = new Set(table.lockedCols || []);
            const lockedCells = new Set(table.lockedCells || []);

            // 用戶編輯快照：先清除所有數據單元格再整體寫入
            if (update._isUserEdit) {
                for (const key of Object.keys(table.data)) {
                    const [r, c] = key.split('-').map(Number);
                    if (r >= 1 && c >= 1) delete table.data[key];
                }
            }

            let updatedCount = 0;
            let blockedCount = 0;

            for (const [key, value] of Object.entries(update.updates)) {
                const [r, c] = key.split('-').map(Number);

                // 用戶編輯不受 header 保護和鎖定限制
                if (!update._isUserEdit) {
                    if (r === 0 || c === 0) {
                        const existing = table.data[key];
                        if (existing && existing.trim()) continue;
                    }

                    if (lockedRows.has(r) || lockedCols.has(c) || lockedCells.has(key)) {
                        blockedCount++;
                        continue;
                    }
                }

                table.data[key] = value;
                updatedCount++;

                if (r + 1 > (table.rows || 2)) table.rows = r + 1;
                if (c + 1 > (table.cols || 2)) table.cols = c + 1;
            }

            // 全域表格：將維度變更同步回 per-card overlay
            if (isGlobal) {
                const perCardData = firstMsg.horae_meta?.globalTableData;
                if (perCardData?.[updateName]) {
                    perCardData[updateName].rows = table.rows;
                    perCardData[updateName].cols = table.cols;
                }
            }

            if (blockedCount > 0) {
                console.log(`[Horae] 表格 "${updateName}" 攔截 ${blockedCount} 個鎖定單元格的修改`);
            }
            console.log(`[Horae] 表格 "${updateName}" 已更新 ${updatedCount} 個單元格`);
        }
    }

    /** 重建表格數據（訊息刪除/編輯後保持一致性） */
    rebuildTableData(maxIndex = -1) {
        const chat = this.getChat();
        if (!chat || chat.length === 0) return;
        
        const firstMsg = chat[0];
        const limit = maxIndex >= 0 ? Math.min(maxIndex + 1, chat.length) : chat.length;

        // 輔助：重置單個表格到 baseData
        const resetTable = (table) => {
            if (table.baseData) {
                table.data = JSON.parse(JSON.stringify(table.baseData));
            } else {
                if (!table.data) { table.data = {}; return; }
                const keysToDelete = [];
                for (const key of Object.keys(table.data)) {
                    const [r, c] = key.split('-').map(Number);
                    if (r >= 1 && c >= 1) keysToDelete.push(key);
                }
                for (const key of keysToDelete) delete table.data[key];
            }
            if (table.baseRows !== undefined) {
                table.rows = table.baseRows;
            } else if (table.baseData) {
                let calcRows = 2, calcCols = 2;
                for (const key of Object.keys(table.baseData)) {
                    const [r, c] = key.split('-').map(Number);
                    if (r === 0 && c + 1 > calcCols) calcCols = c + 1;
                    if (c === 0 && r + 1 > calcRows) calcRows = r + 1;
                }
                table.rows = calcRows;
                table.cols = calcCols;
            }
            if (table.baseCols !== undefined) {
                table.cols = table.baseCols;
            }
        };

        // 1a. 重置本地表格
        const localTables = firstMsg.horae_meta?.customTables || [];
        for (const table of localTables) {
            resetTable(table);
        }

        // 1b. 重置全域表格的 per-card overlay
        const perCardData = firstMsg.horae_meta?.globalTableData || {};
        for (const overlay of Object.values(perCardData)) {
            resetTable(overlay);
        }
        
        // 2. 預掃描：找到每個表格最後一個 _isUserEdit 所在的訊息索引
        const lastUserEditIdx = new Map();
        for (let i = 0; i < limit; i++) {
            const meta = chat[i]?.horae_meta;
            if (meta?.tableContributions) {
                for (const tc of meta.tableContributions) {
                    if (tc._isUserEdit) {
                        lastUserEditIdx.set((tc.name || '').trim(), i);
                    }
                }
            }
        }

        // 3. 按訊息順序重播 tableContributions（截斷到 limit）
        // 防禦：如果某表格存在用戶編輯快照，跳過該快照之前的所有 AI 貢獻
        let totalApplied = 0;
        for (let i = 0; i < limit; i++) {
            const meta = chat[i]?.horae_meta;
            if (meta?.tableContributions && meta.tableContributions.length > 0) {
                const filtered = meta.tableContributions.filter(tc => {
                    if (tc._isUserEdit) return true;
                    const name = (tc.name || '').trim();
                    const ueIdx = lastUserEditIdx.get(name);
                    if (ueIdx !== undefined && i <= ueIdx) return false;
                    return true;
                });
                if (filtered.length > 0) {
                    this.applyTableUpdates(filtered);
                    totalApplied++;
                }
            }
        }
        
        console.log(`[Horae] 表格數據已重建，重播了 ${totalApplied} 條訊息的表格貢獻（截止到#${limit - 1}）`);
    }

    /** 掃描並注入歷史記錄 */
    async scanAndInjectHistory(progressCallback, analyzeCallback = null) {
        const chat = this.getChat();
        let processed = 0;
        let skipped = 0;

        // 需要在覆寫 meta 時保留的全域/摘要相關資料欄
        const PRESERVE_KEYS = [
            'autoSummaries', 'customTables', 'globalTableData',
            'locationMemory', 'relationships', 'tableContributions'
        ];

        for (let i = 0; i < chat.length; i++) {
            const message = chat[i];
            
            if (message.is_user) {
                skipped++;
                if (progressCallback) {
                    progressCallback(Math.round((i + 1) / chat.length * 100), i + 1, chat.length);
                }
                continue;
            }

            // 跳過已有元數據
            const hasEvents = message.horae_meta?.events?.length > 0 || message.horae_meta?.event?.summary;
            if (message.horae_meta && (
                message.horae_meta.timestamp?.story_date ||
                hasEvents ||
                Object.keys(message.horae_meta.costumes || {}).length > 0
            )) {
                skipped++;
                if (progressCallback) {
                    progressCallback(Math.round((i + 1) / chat.length * 100), i + 1, chat.length);
                }
                continue;
            }

            // 保留已有 meta 上的全域數據和事件標記
            const existing = message.horae_meta;
            const preserved = {};
            if (existing) {
                for (const key of PRESERVE_KEYS) {
                    if (existing[key] !== undefined) preserved[key] = existing[key];
                }
                // 保留事件上的摘要標記（_compressedBy / _summaryId）
                if (existing.events?.length > 0) preserved._existingEvents = existing.events;
            }

            const parsed = this.parseHoraeTag(message.mes);
            
            if (parsed) {
                const meta = this.mergeParsedToMeta(null, parsed);
                if (meta._tableUpdates) {
                    meta.tableContributions = meta._tableUpdates;
                    delete meta._tableUpdates;
                }
                // 恢復保留資料欄
                Object.assign(meta, preserved);
                delete meta._existingEvents;
                this.setMessageMeta(i, meta);
                processed++;
            } else if (analyzeCallback) {
                try {
                    const analyzed = await analyzeCallback(message.mes);
                    if (analyzed) {
                        const meta = this.mergeParsedToMeta(null, analyzed);
                        if (meta._tableUpdates) {
                            meta.tableContributions = meta._tableUpdates;
                            delete meta._tableUpdates;
                        }
                        Object.assign(meta, preserved);
                        delete meta._existingEvents;
                        this.setMessageMeta(i, meta);
                        processed++;
                    }
                } catch (error) {
                    console.error(`[Horae] 分析訊息 #${i} 失敗:`, error);
                }
            } else {
                const meta = createEmptyMeta();
                Object.assign(meta, preserved);
                delete meta._existingEvents;
                this.setMessageMeta(i, meta);
                processed++;
            }

            if (progressCallback) {
                progressCallback(Math.round((i + 1) / chat.length * 100), i + 1, chat.length);
            }
        }

        return { processed, skipped };
    }

    generateSystemPromptAddition() {
        const userName = this.context?.name1 || '主角';
        const charName = this.context?.name2 || '角色';
        
        if (this.settings?.customSystemPrompt) {
            const custom = this.settings.customSystemPrompt
                .replace(/\{\{user\}\}/gi, userName)
                .replace(/\{\{char\}\}/gi, charName);
            return custom + this.generateLocationMemoryPrompt() + this.generateCustomTablesPrompt() + this.generateRelationshipPrompt() + this.generateMoodPrompt() + this.generateRpgPrompt();
        }
        
        const sceneDescLine = this.settings?.sendLocationMemory ? '\nscene_desc:地點固定物理特徵（見場景記憶規則，觸發時才寫）' : '';
        const relLine = this.settings?.sendRelationships ? '\nrel:角色A>角色B=關係類型|備註（見關係網路規則，觸發時才寫）' : '';
        const moodLine = this.settings?.sendMood ? '\nmood:角色名=情緒/心理狀態（見情緒追蹤規則，觸發時才寫）' : '';
        return `
【Horae記憶系統】（以下示例僅為示範，勿直接原句用於正文！）

═══ 核心原則：變化驅動 ═══
★★★ 在寫<horae>標籤前，先判斷本回合哪些資訊發生了實質變化 ★★★
  ① 場景基礎（time/location/characters/costume）→ 每回合必填
  ② 其他所有資料欄 → 嚴格遵守各自的【觸發條件】，無變化則完全不寫該行
  ③ 已記錄的NPC/物品若無新資訊 → 禁止輸出！重複輸出無變化的數據=浪費token
  ④ 部分資料欄變化 → 使用增量更新，只寫變化的部分
  ⑤ NPC首次出場 → npc:和affection:兩行都必須寫！

═══ 標籤格式 ═══
每次回覆末尾必須寫入兩個標籤：
<horae>
time:日期 時間（必填）
location:地點（必填。多級地點用·分隔，如「酒館·大廳」「皇宮·王座間」。同一地點每次必須使用完全一致的名稱）
atmosphere:氛圍${sceneDescLine}
characters:在場角色名,逗號分隔（必填）
costume:角色名=服裝描述（必填，每人一行，禁止分號合併）
item/item!/item!!:見物品規則（觸發時才寫）
item-:物品名（物品消耗/遺失時刪除。見物品規則，觸發時才寫）
affection:角色名=好感度（★NPC首次出場必填初始值！之後僅好感變化時更新）
npc:角色名|外貌=個性@關係~擴展資料欄（★NPC首次出場必填完整資訊！之後僅變化時更新）
agenda:日期|內容（新待辦觸發時才寫）
agenda-:內容關鍵詞（待辦已完成/失效時才寫，系統自動移除配對的待辦）${relLine}${moodLine}
</horae>
<horaeevent>
event:重要程度|事件簡述（30-50字，重要程度：一般/重要/關鍵，記錄本條訊息中的事件摘要，用於劇情追溯）
</horaeevent>

═══ 【物品】觸發條件與規則 ═══
參照[物品清單]中的編號(#ID)，嚴格按以下條件決定是否輸出。

【何時寫】（滿足任一條件才輸出）
  ✦ 獲得新物品 → item:/item!:/item!!:
  ✦ 已有物品的數量/歸屬/位置/性質發生改變 → item:（僅寫變化部分）
  ✦ 物品消耗/遺失/用完 → item-:物品名
【何時不寫】
  ✗ 物品無任何變化 → 禁止輸出任何item行
  ✗ 物品僅被提及但無狀態改變 → 不寫

【格式】
  新獲得：item:emoji物品名(數量)|描述=持有者@精確位置（可省略描述資料欄。除非該物品有特殊含意，如禮物、紀念品，則添加描述）
  新獲得(重要)：item!:emoji物品名(數量)|描述=持有者@精確位置（重要物品，描述必填：外觀+功能+來源）
  新獲得(關鍵)：item!!:emoji物品名(數量)|描述=持有者@精確位置（關鍵道具，描述必須詳細）
  已有物品變化：item:emoji物品名(新數量)=新持有者@新位置（僅更新變化的部分，不寫|則保留原描述）
  消耗/遺失：item-:物品名

【資料欄級規則】
  · 描述：記錄物品本質屬性（外觀/功能/來源），普通物品可省略，重要/關鍵物品首次必填
    ★ 外觀特徵（顏色、材質、大小等，便於後續一致性描寫）
    ★ 功能/用途
    ★ 來源（誰給的/如何獲得）
       - 示例（以下內容中若有示例僅為示範，勿直接原句用於正文！）：
         - 示例1：item!:🌹永生花束|深紅色玫瑰永生花，黑色緞帶束扎，C贈送給U的情人節禮物=U@U房間書桌上
         - 示例2：item!:🎫幸運十連抽券|閃著金光的紙質獎券，可在系統獎池進行一次十連抽的新手福利=U@空間戒指
         - 示例3：item!!:🏧位面貨幣自動兌換機|看起來像個小型的ATM機，能按即時匯率兌換各位面貨幣=U@酒館吧檯
  · 數量：單件不寫(1)/(1個)/(1把)等，只有計量組織才寫括號如(5斤)(1L)(1箱)
  · 位置：必須是精確固定地點
    ❌ 某某人身前地上、某某人腳邊、某某人旁邊、地板、桌子上
    ✅ 酒館大廳地板、餐廳吧檯上、家中廚房、揹包裡、U的房間桌子上
  · 禁止將固定傢俱和建築設施計入物品
  · 臨時借用≠歸屬轉移


示例（麥酒生命週期）：
  獲得：item:🍺陳釀麥酒(50L)|雜物間翻出的麥酒，口感酸澀=U@酒館後廚食材櫃
  量變：item:🍺陳釀麥酒(25L)=U@酒館後廚食材櫃
  用完：item-:陳釀麥酒

═══ 【NPC】觸發條件與規則 ═══
格式：npc:名|外貌=個性@與${userName}的關係~性別:值~年齡:值~種族:值~職業:值~生日:值
分隔符：| 分名字，= 分外貌與個性，@ 分關係，~ 分擴展資料欄(key:value)

【何時寫】（滿足任一條件才輸出該NPC的npc:行）
  ✦ 首次出場 → 完整格式，全部資料欄+全部~擴展資料欄（性別/年齡/種族/職業），缺一不可
  ✦ 外貌永久變化（如受傷留疤、換了髮型、穿戴改變）→ 只寫外貌資料欄
  ✦ 個性發生轉變（如經歷重大事件後個性改變）→ 只寫個性資料欄
  ✦ 與${userName}的關係定位改變（如從客人變成朋友）→ 只寫關係資料欄
  ✦ 獲得關於該NPC的新資訊（之前不知道的身高/體重等）→ 追加到對應資料欄
  ✦ ~擴展資料欄本身發生變化（如職業變了）→ 只寫變化的~擴展資料欄
【何時不寫】
  ✗ NPC在場但無新資訊 → 禁止寫npc:行
  ✗ NPC暫時離場後回來，資訊無變化 → 禁止重寫
  ✗ 想用同義詞/縮寫重寫已有描述 → 嚴禁！
    ❌ "肌肉發達/滿身戰鬥傷痕"→"肌肉強壯/傷疤"（換詞≠更新）
    ✅ "肌肉發達/滿身戰鬥傷痕/重傷"→"肌肉發達/滿身戰鬥傷痕"（傷愈，移除過時狀態）

【增量更新示例】（以NPC沃爾為例）
  首次：npc:沃爾|銀灰色披毛/綠眼睛/身高220cm/滿身戰鬥傷痕=沉默寡言的重裝傭兵@${userName}的第一個客人~性別:男~年齡:約35~種族:狼獸人~職業:傭兵
  只更新關係：npc:沃爾|=@${userName}的男朋友
  只追加外貌：npc:沃爾|銀灰色披毛/綠眼睛/身高220cm/滿身戰鬥傷痕/左臂繃帶
  只更新個性：npc:沃爾|=不再沉默/偶爾微笑
  只改職業：npc:沃爾|~職業:退役傭兵
（注意：未變化的資料欄和~擴展資料欄完全不寫！系統自動保留原有數據！）

【生日資料欄（可選擴展資料欄）】
  格式：~生日:yyyy/mm/dd 或 ~生日:mm/dd（無年份時僅寫月日）
  ⚠ 僅當角色設定/人物描述中明確提及生日日期時才寫！嚴禁猜測或捏造！
  ⚠ 沒有明確出處的生日一律不寫此資料欄——留空由用戶自行填寫。

【關係描述規範】
  必須包含對象名且準確：❌客人 ✅${userName}的新訪客 / ❌債主 ✅持有${userName}欠條的人 / ❌房東 ✅${userName}的房東 / ❌男朋友 ✅${userName}的男朋友 / ❌恩人 ✅救了${userName}一命的人 / ❌霸凌者 ✅欺負${userName}的人 / ❌暗戀者 ✅暗戀${userName}的人 / ❌仇人 ✅被${userName}殺掉了生父
  附屬關係需寫出所屬NPC名：✅伊凡的獵犬; ${userName}客人的寵物 / 伊凡的女朋友; ${userName}的客人 / ${userName}的閨蜜; 伊凡的妻子 / ${userName}的繼父; 伊凡的父親 / ${userName}的情夫; 伊凡的弟弟 / ${userName}的閨蜜; ${userName}的丈夫的情婦; 插足${userName}與伊凡夫妻關係的第三者

═══ 【好感度】觸發條件 ═══
僅記錄NPC對${userName}的好感度（禁止記錄${userName}自己）。每人一行，禁止數值後加註解。

【何時寫】
  ✦ NPC首次出場 → 按關係判定初始值（陌生0-20/熟人30-50/朋友50-70/戀人70-90）
  ✦ 互動導致好感度實質變化 → affection:名=新總值
【何時不寫】
  ✗ 好感度無變化 → 不寫

═══ 【待辦事項】觸發條件 ═══
【何時寫（新增）】
  ✦ 劇情中出現新的約定/計劃/行程/任務/伏筆 → agenda:日期|內容
  格式：agenda:訂立日期|內容（相對時間須括號標註絕對日期）
  示例：agenda:2026/02/10|艾倫邀請${userName}情人節晚上約會(2026/02/14 18:00)
【何時寫（完成刪除）— 極重要！】
  ✦ 待辦事項已完成/已失效/已取消 → 必須用 agenda-: 標記刪除
  格式：agenda-:待辦內容（寫入已完成事項的內容關鍵詞即可自動移除）
  示例：agenda-:艾倫邀請${userName}情人節晚上約會
  ⚠ 嚴禁用 agenda:內容(完成) 這種方式！必須用 agenda-: 前綴！
  ⚠ 嚴禁重複寫入已存在的待辦內容！
【何時不寫】
  ✗ 已有待辦無變化 → 禁止每回合重複已有待辦
  ✗ 待辦已完成 → 禁止用 agenda: 加括號標註完成，必須用 agenda-:

═══ 時間格式規則 ═══
禁止"Day 1"/"第X天"等模糊格式，必須使用具體日曆日期。
- 現代：年/月/日 時:分（如 2026/2/4 15:00）
- 歷史：該年代日期（如 1920/3/15 14:00）
- 奇幻/架空：該世界觀日曆（如 霜降月第三日 黃昏）
${this.generateLocationMemoryPrompt()}${this.generateCustomTablesPrompt()}${this.generateRelationshipPrompt()}${this.generateMoodPrompt()}${this.generateRpgPrompt()}${this._generateAntiParaphrasePrompt()}
═══ 最終強制提醒 ═══
${this._generateMustTagsReminder()}

【每回合必寫資料欄——缺任何一項=不合格！】
  ✅ time: ← 目前日期時間
  ✅ location: ← 目前地點
  ✅ atmosphere: ← 氛圍
  ✅ characters: ← 目前在場所有角色名，逗號分隔（絕對不能省略！）
  ✅ costume: ← 每個在場角色各一行服裝描述
  ✅ event: ← 重要程度|事件摘要

【NPC首次登場時額外必寫——缺一不可！】
  ✅ npc:名|外貌=個性@關係~性別:值~年齡:值~種族:值~職業:值~生日:值(僅已知時寫，未知不寫)
  ✅ affection:該NPC名=初始好感度（陌生0-20/熟人30-50/朋友50-70/戀人70-90）

以上資料欄不存在"可寫可不寫"的情況——它們是強制性的。
`;
    }

    getDefaultSystemPrompt() {
        const sceneDescLine = this.settings?.sendLocationMemory ? '\nscene_desc:地點固定物理特徵（見場景記憶規則，觸發時才寫）' : '';
        const relLine = this.settings?.sendRelationships ? '\nrel:角色A>角色B=關係類型|備註（見關係網路規則，觸發時才寫）' : '';
        const moodLine = this.settings?.sendMood ? '\nmood:角色名=情緒/心理狀態（見情緒追蹤規則，觸發時才寫）' : '';
        return `【Horae記憶系統】（以下示例僅為示範，勿直接原句用於正文！）

═══ 核心原則：變化驅動 ═══
★★★ 在寫<horae>標籤前，先判斷本回合哪些資訊發生了實質變化 ★★★
  ① 場景基礎（time/location/characters/costume）→ 每回合必填
  ② 其他所有資料欄 → 嚴格遵守各自的【觸發條件】，無變化則完全不寫該行
  ③ 已記錄的NPC/物品若無新資訊 → 禁止輸出！重複輸出無變化的數據=浪費token
  ④ 部分資料欄變化 → 使用增量更新，只寫變化的部分
  ⑤ NPC首次出場 → npc:和affection:兩行都必須寫！

═══ 標籤格式 ═══
每次回覆末尾必須寫入兩個標籤：
<horae>
time:日期 時間（必填）
location:地點（必填。多級地點用·分隔，如「酒館·大廳」「皇宮·王座間」。同一地點每次必須使用完全一致的名稱）
atmosphere:氛圍${sceneDescLine}
characters:在場角色名,逗號分隔（必填）
costume:角色名=服裝描述（必填，每人一行，禁止分號合併）
item/item!/item!!:見物品規則（觸發時才寫）
item-:物品名（物品消耗/遺失時刪除。見物品規則，觸發時才寫）
affection:角色名=好感度（★NPC首次出場必填初始值！之後僅好感變化時更新）
npc:角色名|外貌=個性@關係~擴展資料欄（★NPC首次出場必填完整資訊！之後僅變化時更新）
agenda:日期|內容（新待辦觸發時才寫）
agenda-:內容關鍵詞（待辦已完成/失效時才寫，系統自動移除配對的待辦）${relLine}${moodLine}
</horae>
<horaeevent>
event:重要程度|事件簡述（30-50字，重要程度：一般/重要/關鍵，記錄本條訊息中的事件摘要，用於劇情追溯）
</horaeevent>

═══ 【物品】觸發條件與規則 ═══
參照[物品清單]中的編號(#ID)，嚴格按以下條件決定是否輸出。

【何時寫】（滿足任一條件才輸出）
  ✦ 獲得新物品 → item:/item!:/item!!:
  ✦ 已有物品的數量/歸屬/位置/性質發生改變 → item:（僅寫變化部分）
  ✦ 物品消耗/遺失/用完 → item-:物品名
【何時不寫】
  ✗ 物品無任何變化 → 禁止輸出任何item行
  ✗ 物品僅被提及但無狀態改變 → 不寫

【格式】
  新獲得：item:emoji物品名(數量)|描述=持有者@精確位置（可省略描述資料欄。除非該物品有特殊含意，如禮物、紀念品，則添加描述）
  新獲得(重要)：item!:emoji物品名(數量)|描述=持有者@精確位置（重要物品，描述必填：外觀+功能+來源）
  新獲得(關鍵)：item!!:emoji物品名(數量)|描述=持有者@精確位置（關鍵道具，描述必須詳細）
  已有物品變化：item:emoji物品名(新數量)=新持有者@新位置（僅更新變化的部分，不寫|則保留原描述）
  消耗/遺失：item-:物品名

【資料欄級規則】
  · 描述：記錄物品本質屬性（外觀/功能/來源），普通物品可省略，重要/關鍵物品首次必填
    ★ 外觀特徵（顏色、材質、大小等，便於後續一致性描寫）
    ★ 功能/用途
    ★ 來源（誰給的/如何獲得）
       - 示例（以下內容中若有示例僅為示範，勿直接原句用於正文！）：
         - 示例1：item!:🌹永生花束|深紅色玫瑰永生花，黑色緞帶束扎，C贈送給U的情人節禮物=U@U房間書桌上
         - 示例2：item!:🎫幸運十連抽券|閃著金光的紙質獎券，可在系統獎池進行一次十連抽的新手福利=U@空間戒指
         - 示例3：item!!:🏧位面貨幣自動兌換機|看起來像個小型的ATM機，能按即時匯率兌換各位面貨幣=U@酒館吧檯
  · 數量：單件不寫(1)/(1個)/(1把)等，只有計量組織才寫括號如(5斤)(1L)(1箱)
  · 位置：必須是精確固定地點
    ❌ 某某人身前地上、某某人腳邊、某某人旁邊、地板、桌子上
    ✅ 酒館大廳地板、餐廳吧檯上、家中廚房、揹包裡、U的房間桌子上
  · 禁止將固定傢俱和建築設施計入物品
  · 臨時借用≠歸屬轉移


示例（麥酒生命週期）：
  獲得：item:🍺陳釀麥酒(50L)|雜物間翻出的麥酒，口感酸澀=U@酒館後廚食材櫃
  量變：item:🍺陳釀麥酒(25L)=U@酒館後廚食材櫃
  用完：item-:陳釀麥酒

═══ 【NPC】觸發條件與規則 ═══
格式：npc:名|外貌=個性@與{{user}}的關係~性別:值~年齡:值~種族:值~職業:值~生日:值
分隔符：| 分名字，= 分外貌與個性，@ 分關係，~ 分擴展資料欄(key:value)

【何時寫】（滿足任一條件才輸出該NPC的npc:行）
  ✦ 首次出場 → 完整格式，全部資料欄+全部~擴展資料欄（性別/年齡/種族/職業），缺一不可
  ✦ 外貌永久變化（如受傷留疤、換了髮型、穿戴改變）→ 只寫外貌資料欄
  ✦ 個性發生轉變（如經歷重大事件後個性改變）→ 只寫個性資料欄
  ✦ 與{{user}}的關係定位改變（如從客人變成朋友）→ 只寫關係資料欄
  ✦ 獲得關於該NPC的新資訊（之前不知道的身高/體重等）→ 追加到對應資料欄
  ✦ ~擴展資料欄本身發生變化（如職業變了）→ 只寫變化的~擴展資料欄
【何時不寫】
  ✗ NPC在場但無新資訊 → 禁止寫npc:行
  ✗ NPC暫時離場後回來，資訊無變化 → 禁止重寫
  ✗ 想用同義詞/縮寫重寫已有描述 → 嚴禁！
    ❌ "肌肉發達/滿身戰鬥傷痕"→"肌肉強壯/傷疤"（換詞≠更新）
    ✅ "肌肉發達/滿身戰鬥傷痕/重傷"→"肌肉發達/滿身戰鬥傷痕"（傷愈，移除過時狀態）

【增量更新示例】（以NPC沃爾為例）
  首次：npc:沃爾|銀灰色披毛/綠眼睛/身高220cm/滿身戰鬥傷痕=沉默寡言的重裝傭兵@{{user}}的第一個客人~性別:男~年齡:約35~種族:狼獸人~職業:傭兵
  只更新關係：npc:沃爾|=@{{user}}的男朋友
  只追加外貌：npc:沃爾|銀灰色披毛/綠眼睛/身高220cm/滿身戰鬥傷痕/左臂繃帶
  只更新個性：npc:沃爾|=不再沉默/偶爾微笑
  只改職業：npc:沃爾|~職業:退役傭兵
（注意：未變化的資料欄和~擴展資料欄完全不寫！系統自動保留原有數據！）

【生日資料欄（可選擴展資料欄）】
  格式：~生日:yyyy/mm/dd 或 ~生日:mm/dd（無年份時僅寫月日）
  ⚠ 僅當角色設定/人物描述中明確提及生日日期時才寫！嚴禁猜測或捏造！
  ⚠ 沒有明確出處的生日一律不寫此資料欄——留空由用戶自行填寫。

【關係描述規範】
  必須包含對象名且準確：❌客人 ✅{{user}}的新訪客 / ❌債主 ✅持有{{user}}欠條的人 / ❌房東 ✅{{user}}的房東 / ❌男朋友 ✅{{user}}的男朋友 / ❌恩人 ✅救了{{user}}一命的人 / ❌霸凌者 ✅欺負{{user}}的人 / ❌暗戀者 ✅暗戀{{user}}的人 / ❌仇人 ✅被{{user}}殺掉了生父
  附屬關係需寫出所屬NPC名：✅伊凡的獵犬; {{user}}客人的寵物 / 伊凡的女朋友; {{user}}的客人 / {{user}}的閨蜜; 伊凡的妻子 / {{user}}的繼父; 伊凡的父親 / {{user}}的情夫; 伊凡的弟弟 / {{user}}的閨蜜; {{user}}的丈夫的情婦; 插足{{user}}與伊凡夫妻關係的第三者

═══ 【好感度】觸發條件 ═══
僅記錄NPC對{{user}}的好感度（禁止記錄{{user}}自己）。每人一行，禁止數值後加註解。

【何時寫】
  ✦ NPC首次出場 → 按關係判定初始值（陌生0-20/熟人30-50/朋友50-70/戀人70-90）
  ✦ 互動導致好感度實質變化 → affection:名=新總值
【何時不寫】
  ✗ 好感度無變化 → 不寫

═══ 【待辦事項】觸發條件 ═══
【何時寫（新增）】
  ✦ 劇情中出現新的約定/計劃/行程/任務/伏筆 → agenda:日期|內容
  格式：agenda:訂立日期|內容（相對時間須括號標註絕對日期）
  示例：agenda:2026/02/10|艾倫邀請{{user}}情人節晚上約會(2026/02/14 18:00)
【何時寫（完成刪除）— 極重要！】
  ✦ 待辦事項已完成/已失效/已取消 → 必須用 agenda-: 標記刪除
  格式：agenda-:待辦內容（寫入已完成事項的內容關鍵詞即可自動移除）
  示例：agenda-:艾倫邀請{{user}}情人節晚上約會
  ⚠ 嚴禁用 agenda:內容(完成) 這種方式！必須用 agenda-: 前綴！
  ⚠ 嚴禁重複寫入已存在的待辦內容！
【何時不寫】
  ✗ 已有待辦無變化 → 禁止每回合重複已有待辦
  ✗ 待辦已完成 → 禁止用 agenda: 加括號標註完成，必須用 agenda-:

═══ 時間格式規則 ═══
禁止"Day 1"/"第X天"等模糊格式，必須使用具體日曆日期。
- 現代：年/月/日 時:分（如 2026/2/4 15:00）
- 歷史：該年代日期（如 1920/3/15 14:00）
- 奇幻/架空：該世界觀日曆（如 霜降月第三日 黃昏）

═══ 最終強制提醒 ═══
你的回覆末尾必須包含 <horae>...</horae> 和 <horaeevent>...</horaeevent> 兩個標籤。
缺少任何一個標籤=不合格。

【每回合必寫資料欄——缺任何一項=不合格！】
  ✅ time: ← 目前日期時間
  ✅ location: ← 目前地點
  ✅ atmosphere: ← 氛圍
  ✅ characters: ← 目前在場所有角色名，逗號分隔（絕對不能省略！）
  ✅ costume: ← 每個在場角色各一行服裝描述
  ✅ event: ← 重要程度|事件摘要

【NPC首次登場時額外必寫——缺一不可！】
  ✅ npc:名|外貌=個性@關係~性別:值~年齡:值~種族:值~職業:值~生日:值(僅已知時寫，未知不寫)
  ✅ affection:該NPC名=初始好感度（陌生0-20/熟人30-50/朋友50-70/戀人70-90）

以上資料欄不存在"可寫可不寫"的情況——它們是強制性的。`;
    }

    getDefaultTablesPrompt() {
        return `═══ 客製化表格規則 ═══
上方有用戶客製化表格，根據"填寫要求"填寫數據。
★ 格式：<horaetable:表格名> 標籤內，每行一個單元格 → 行,列:內容
★★ 座標說明：第0行和第0列是表頭，數據從1,1開始。行號=數據行序號，列號=數據列序號
★★★ 填寫原則 ★★★
  - 空單元格且劇情中已有對應資訊 → 必須填寫！不要遺漏！
  - 已有內容且無變化 → 不重複寫
  - 該行/列確實無對應劇情資訊 → 留空
  - 禁止輸出"(空)""-""無"等佔位符
  - 🔒標記的行/列為只讀數據，禁止修改其內容
  - 新增行請在現有最大行號之後追加，新增列請在現有最大列號之後追加`;
    }

    getDefaultLocationPrompt() {
        return `═══ 【場景記憶】觸發條件 ═══
格式：scene_desc:位於…。該地點的固定物理特徵描述（50-150字）
場景記憶記錄地點的核心格局和永久性特徵（建築結構、固定傢俱、空間特點），用於保持跨回合的場景描寫一致性。

【地點／位於 格式】★★★ 嚴格遵守層級規則 ★★★
  · 描述開頭先寫「位於」標明該地點相對於直接上級的方位，再寫該地點自身的物理特徵
  · 子級地點（含·分隔符的地名）：「位於」只寫相對於父級建築內部的方位（如哪一樓、哪個方向），絕對禁止包含父級的外部地理位置
  · 父級/頂級地點：「位於」才寫外部地理位置（如哪個大陸、哪片森林旁）
  · 系統會自動同時發送父級描述給AI，子級無需也不應重複父級資訊
    ✓ 無名酒館·客房203 → scene_desc:位於2樓東側。邊間，採光佳，單人木床靠牆，窗戶朝東
    ✓ 無名酒館·大廳 → scene_desc:位於1樓。挑高木質空間，正中是長吧檯，散落數張圓桌
    ✓ 無名酒館 → scene_desc:位於OO大陸北方XX森林邊上。兩層木石結構，一樓大廳和吧檯，二樓客房區
    ✗ 無名酒館·客房203 → scene_desc:位於OO大陸北方XX森林邊上的無名酒館2樓…（❌ 子級禁止寫父級的外部地理資訊）
    ✗ 無名酒館·大廳 → scene_desc:位於森林邊上的無名酒館1樓…（❌ 同上）
【地名規範】
  · 多級地點用·分隔：建築·區域（如「無名酒館·大廳」「皇宮·地牢」）
  · 同一地點必須始終使用與上方[場景|...]中完全一致的名稱，禁止縮寫或改寫
  · 不同建築的同名區域各自獨立記錄（如「無名酒館·大廳」和「皇宮·大廳」是不同地點）
【何時寫】
  ✦ 首次到達一個新地點 → 必須寫scene_desc，描述該地點的固定物理特徵
  ✦ 地點發生永久性物理變化（如被破壞、重新裝修）→ 寫更新後的scene_desc
【何時不寫】
  ✗ 回到已記錄的舊地點且無物理變化 → 不寫
  ✗ 季節/天氣/氛圍變化 → 不寫（這些是臨時變化，不屬於固定特徵）
【描述規範】
  · 只寫固定/永久性的物理特徵：空間結構、建築材質、固定傢俱、窗戶朝向、標誌性裝飾
  · 不寫臨時性狀態：目前光照、天氣、人群、季節裝飾、臨時擺放的物品
  · 禁止照搬場景記憶原文到正文，將其作為背景參考，以目前時間/天氣/光線/角色視角重新描寫
  · 上方[場景記憶|...]是系統已記錄的該地點特徵，描寫該場景時保持這些核心要素不變，同時根據時間/季節/劇情自由發揮變化細節`;
    }

    generateLocationMemoryPrompt() {
        if (!this.settings?.sendLocationMemory) return '';
        const custom = this.settings?.customLocationPrompt;
        if (custom) {
            const userName = this.context?.name1 || '主角';
            const charName = this.context?.name2 || '角色';
            return '\n' + custom.replace(/\{\{user\}\}/gi, userName).replace(/\{\{char\}\}/gi, charName);
        }
        return '\n' + this.getDefaultLocationPrompt();
    }

    generateCustomTablesPrompt() {
        const chat = this.getChat();
        const firstMsg = chat?.[0];
        const localTables = firstMsg?.horae_meta?.customTables || [];
        const resolvedGlobal = this._getResolvedGlobalTables();
        const allTables = [...resolvedGlobal, ...localTables];
        if (allTables.length === 0) return '';

        let prompt = '\n' + (this.settings?.customTablesPrompt || this.getDefaultTablesPrompt());

        // 為每個表格生成帶座標的示例
        for (const table of allTables) {
            const tableName = table.name || '客製化表格';
            const rows = table.rows || 2;
            const cols = table.cols || 2;
            prompt += `\n★ 表格「${tableName}」尺寸：${rows - 1}行×${cols - 1}列（數據區行號1-${rows - 1}，列號1-${cols - 1}）`;
            prompt += `\n示例（填寫空單元格或更新有變化的單元格）：
<horaetable:${tableName}>
1,1:內容A
1,2:內容B
2,1:內容C
</horaetable>`;
            break;
        }

        return prompt;
    }

    getDefaultRelationshipPrompt() {
        const userName = this.context?.name1 || '{{user}}';
        return `═══ 【關係網路】觸發條件 ═══
格式：rel:角色A>角色B=關係類型|備註
系統會自動記錄和顯示角色間的關係網路，當角色間關係發生變化時輸出。

【何時寫】（滿足任一條件才輸出）
  ✦ 兩個角色之間確立/定義了新關係 → rel:角色A>角色B=關係類型
  ✦ 已有關係發生變化（如從同事變成朋友）→ rel:角色A>角色B=新關係類型
  ✦ 關係中有重要細節需要備註 → 加|備註
【何時不寫】
  ✗ 關係無變化 → 不寫
  ✗ 已記錄過的關係且無更新 → 不寫

【規範】
  · 角色A和角色B都必須使用準確全名
  · 關係類型用簡潔詞描述：朋友、戀人、上下級、師徒、宿敵、合作伙伴等
  · 備註資料欄可選，記錄關係的特殊細節
  · 包含${userName}的關係也要記錄
  示例：
    rel:${userName}>沃爾=僱傭關係|${userName}經營酒館，沃爾是常客
    rel:沃爾>艾拉=暗戀|沃爾對艾拉有好感但未表白
    rel:${userName}>艾拉=閨蜜`;
    }

    getDefaultMoodPrompt() {
        return `═══ 【情緒/心理狀態追蹤】觸發條件 ═══
格式：mood:角色名=情緒狀態（簡潔詞組，如"緊張/不安"、"開心/期待"、"憤怒"、"平靜但警惕"）
系統會追蹤在場角色的情緒變化，幫助保持角色心理狀態的連貫性。

【何時寫】（滿足任一條件才輸出）
  ✦ 角色情緒發生明顯變化（如從平靜變為憤怒）→ mood:角色名=新情緒
  ✦ 角色首次出場時有明顯的情緒特徵 → mood:角色名=目前情緒
【何時不寫】
  ✗ 角色情緒無變化 → 不寫
  ✗ 角色不在場 → 不寫
【規範】
  · 情緒描述用1-4個詞，用/分隔複合情緒
  · 只記錄在場角色的情緒`;
    }

    generateRelationshipPrompt() {
        if (!this.settings?.sendRelationships) return '';
        const custom = this.settings?.customRelationshipPrompt;
        if (custom) {
            const userName = this.context?.name1 || '主角';
            const charName = this.context?.name2 || '角色';
            return '\n' + custom.replace(/\{\{user\}\}/gi, userName).replace(/\{\{char\}\}/gi, charName);
        }
        return '\n' + this.getDefaultRelationshipPrompt();
    }

    _generateAntiParaphrasePrompt() {
        if (!this.settings?.antiParaphraseMode) return '';
        const userName = this.context?.name1 || '主角';
        return `
═══ 反轉述模式（Anti-Paraphrase） ═══
目前用戶使用反轉述寫法：${userName}的行動/對話由${userName}自行在USER訊息中描寫，你（AI）不再重複描述${userName}的部分。
因此，你在撰寫本回合的<horae>標籤時，必須把"緊接在你這條回覆之前的那條USER訊息"中發生的情節也一併納入結算：
  ✦ USER訊息中出現的物品獲取/消耗 → 寫入對應item:/item-:行
  ✦ USER訊息中出現的場景轉移 → 更新location:
  ✦ USER訊息中出現的NPC互動/好感變化 → 更新affection:
  ✦ USER訊息中出現的情節推進 → 在<horaeevent>中一併概括
  ✦ 總之：本條<horae>應同時覆蓋"上一條USER訊息"和"你本條AI回覆"兩部分的所有變化
`;
    }

    generateMoodPrompt() {
        if (!this.settings?.sendMood) return '';
        const custom = this.settings?.customMoodPrompt;
        if (custom) {
            const userName = this.context?.name1 || '主角';
            const charName = this.context?.name2 || '角色';
            return '\n' + custom.replace(/\{\{user\}\}/gi, userName).replace(/\{\{char\}\}/gi, charName);
        }
        return '\n' + this.getDefaultMoodPrompt();
    }

    /** RPG 提示詞（rpgMode 開啟才注入） */
    generateRpgPrompt() {
        if (!this.settings?.rpgMode) return '';
        // 客製化提示詞優先
        if (this.settings.customRpgPrompt) {
            return '\n' + this.settings.customRpgPrompt
                .replace(/\{\{user\}\}/gi, this.context?.name1 || '主角')
                .replace(/\{\{char\}\}/gi, this.context?.name2 || 'AI');
        }
        return '\n' + this.getDefaultRpgPrompt();
    }

    /** RPG 預設提示詞 */
    getDefaultRpgPrompt() {
        const sendBars = this.settings?.sendRpgBars !== false;
        const sendSkills = this.settings?.sendRpgSkills !== false;
        const sendAttrs = this.settings?.sendRpgAttributes !== false;
        const sendEq = !!this.settings?.sendRpgEquipment;
        const sendRep = !!this.settings?.sendRpgReputation;
        const sendLvl = !!this.settings?.sendRpgLevel;
        const sendCur = !!this.settings?.sendRpgCurrency;
        const sendSh = !!this.settings?.sendRpgStronghold;
        if (!sendBars && !sendSkills && !sendAttrs && !sendEq && !sendRep && !sendLvl && !sendCur && !sendSh) return '';
        const userName = this.context?.name1 || '主角';
        const uoBars = !!this.settings?.rpgBarsUserOnly;
        const uoSkills = !!this.settings?.rpgSkillsUserOnly;
        const uoAttrs = !!this.settings?.rpgAttrsUserOnly;
        const uoEq = !!this.settings?.rpgEquipmentUserOnly;
        const uoRep = !!this.settings?.rpgReputationUserOnly;
        const uoLvl = !!this.settings?.rpgLevelUserOnly;
        const uoCur = !!this.settings?.rpgCurrencyUserOnly;
        const anyUo = uoBars || uoSkills || uoAttrs || uoEq || uoRep || uoLvl || uoCur;
        const allUo = uoBars && uoSkills && uoAttrs && uoEq && uoRep && uoLvl && uoCur;
        const barCfg = this.settings?.rpgBarConfig || [
            { key: 'hp', name: 'HP' }, { key: 'mp', name: 'MP' }, { key: 'sp', name: 'SP' }
        ];
        const attrCfg = this.settings?.rpgAttributeConfig || [];
        let p = `═══ 【RPG】 ═══\n你的回覆末尾必須包含<horaerpg>標籤。`;
        if (allUo) {
            p += `所有RPG數據僅追蹤${userName}一人，格式中不含歸屬資料欄。禁止為NPC輸出任何RPG行。\n`;
        } else if (anyUo) {
            p += `歸屬格式同NPC編號：N編號 全名，${userName}直接寫名字不加N。部分模組僅追蹤${userName}（以下會標註）。\n`;
        } else {
            p += `歸屬格式同NPC編號：N編號 全名，${userName}直接寫名字不加N。\n`;
        }
        if (sendBars) {
            p += `\n【屬性條——每回合必寫，缺少=不合格！】\n`;
            if (uoBars) {
                p += `僅輸出${userName}的屬性條和狀態：\n`;
                for (const bar of barCfg) {
                    p += `  ✅ ${bar.key}:目前/最大(${bar.name})  ← 首次必須標註顯示名\n`;
                }
                p += `  ✅ status:效果1/效果2  ← 無異常寫 正常\n`;
            } else {
                p += `必須為 characters: 中每個在場角色輸出全部屬性條和狀態：\n`;
                for (const bar of barCfg) {
                    p += `  ✅ ${bar.key}:歸屬=目前/最大(${bar.name})  ← 首次必須標註顯示名\n`;
                }
                p += `  ✅ status:歸屬=效果1/效果2  ← 無異常寫 =正常\n`;
            }
            p += `規則：\n`;
            p += `  - 戰鬥/受傷/施法/消耗 → 合理扣減；恢復/休息 → 合理回增\n`;
            if (!uoBars) {
                p += `  - 每個在場角色的每個屬性條都必須寫，漏寫任何一人=不合格\n`;
            }
            p += `  - 即使本回合數值無變化，也必須寫出目前值\n`;
        }
        if (sendAttrs && attrCfg.length > 0) {
            p += `\n【多維屬性】僅首次登場或屬性變化時寫，無變化可省略\n`;
            if (uoAttrs) {
                p += `  attr:${attrCfg.map(a => `${a.key}=数值`).join('|')}\n`;
            } else {
                p += `  attr:歸屬|${attrCfg.map(a => `${a.key}=数值`).join('|')}\n`;
            }
            p += `  數值範圍0-100。屬性含義：${attrCfg.map(a => `${a.key}(${a.name})`).join('、')}\n`;
        }
        if (sendSkills) {
            p += `\n【技能】僅習得/更新/失去時寫，無變化可省略\n`;
            if (uoSkills) {
                p += `  skill:技能名|等級|效果描述\n`;
                p += `  skill-:技能名\n`;
            } else {
                p += `  skill:歸屬|技能名|等級|效果描述\n`;
                p += `  skill-:歸屬|技能名\n`;
            }
        }
        if (sendEq) {
            const eqCfg = this._getRpgEquipmentConfig();
            const perChar = eqCfg.perChar || {};
            const present = new Set(this.getLatestState()?.scene?.characters_present || []);
            const hasAnySlots = Object.values(perChar).some(c => c.slots?.length > 0);
            if (hasAnySlots) {
                p += `\n【裝備】角色穿戴/卸下裝備時寫，無變化可省略\n`;
                if (uoEq) {
                    p += `  equip:格位名|裝備名|屬性1=值,屬性2=值\n`;
                    p += `  unequip:格位名|裝備名\n`;
                    const userCfg = perChar[userName];
                    if (userCfg?.slots?.length) {
                        const slotNames = userCfg.slots.map(s => `${s.name}(×${s.maxCount ?? 1})`).join('、');
                        p += `  格位: ${slotNames}\n`;
                    }
                } else {
                    p += `  equip:歸屬|格位名|裝備名|屬性1=值,屬性2=值\n`;
                    p += `  unequip:歸屬|格位名|裝備名\n`;
                    for (const [owner, cfg] of Object.entries(perChar)) {
                        if (!cfg.slots?.length) continue;
                        if (present.size > 0 && !present.has(owner)) continue;
                        const slotNames = cfg.slots.map(s => `${s.name}(×${s.maxCount ?? 1})`).join('、');
                        p += `  ${owner} 格位: ${slotNames}\n`;
                    }
                }
                p += `  ⚠ 每個角色只能使用其已註冊的格位。屬性值為整數。\n`;
                p += `  ⚠ 普通衣物非賦魔或特殊材料不應有高屬性值。\n`;
            }
        }
        if (sendRep) {
            const repConfig = this._getRpgReputationConfig();
            if (repConfig.categories.length > 0) {
                const catNames = repConfig.categories.map(c => c.name).join('、');
                p += `\n【聲望】僅聲望變化時寫，無變化可省略\n`;
                if (uoRep) {
                    p += `  rep:聲望分類名=目前值\n`;
                } else {
                    p += `  rep:歸屬|聲望分類名=目前值\n`;
                }
                p += `  已註冊的聲望分類: ${catNames}\n`;
                p += `  ⚠ 禁止創造新的聲望分類。只允許使用上述已註冊的分類名。\n`;
            }
        }
        if (sendLvl) {
            p += `\n【等級與經驗值】僅更新/降級或經驗變化時寫，無變化可省略\n`;
            if (uoLvl) {
                p += `  level:等級數值\n`;
                p += `  xp:目前經驗/更新所需\n`;
            } else {
                p += `  level:歸屬=等級數值\n`;
                p += `  xp:歸屬=目前經驗/更新所需\n`;
            }
            p += `  經驗值獲取參考：\n`;
            p += `  - 與角色等級相近或更強的挑戰：獲得較多經驗(10~50+)\n`;
            p += `  - 等級差 ≥10 的低級挑戰：僅得 1 點經驗\n`;
            p += `  - 日常活動/對話/探索：少量經驗(1~5)\n`;
            p += `  - 更新所需經驗隨等級遞增：建議 更新所需 = 等級 × 100\n`;
        }
        if (sendCur) {
            const curConfig = this._getRpgCurrencyConfig();
            if (curConfig.denominations.length > 0) {
                const denomNames = curConfig.denominations.map(d => d.name).join('、');
                p += `\n【貨幣——發生交易/拾取/消費時必寫！】\n`;
                if (uoCur) {
                    p += `格式: currency:幣名=±變化量\n`;
                    p += `示例:\n`;
                    p += `  currency:${curConfig.denominations[0].name}=+10\n`;
                    p += `  currency:${curConfig.denominations[0].name}=-3\n`;
                    if (curConfig.denominations.length > 1) {
                        p += `  currency:${curConfig.denominations[1].name}=+50\n`;
                    }
                    p += `也可寫絕對值: currency:幣名=數量\n`;
                } else {
                    p += `格式: currency:歸屬|幣名=±變化量\n`;
                    p += `示例:\n`;
                    p += `  currency:${userName}|${curConfig.denominations[0].name}=+10\n`;
                    p += `  currency:${userName}|${curConfig.denominations[0].name}=-3\n`;
                    if (curConfig.denominations.length > 1) {
                        p += `  currency:${userName}|${curConfig.denominations[1].name}=+50\n`;
                    }
                    p += `也可寫絕對值: currency:歸屬|幣名=數量\n`;
                }
                p += `已註冊幣種: ${denomNames}\n`;
                p += `⚠ 禁止使用未註冊的幣種名。任何涉及金錢的行為（買賣/拾取/獎賞/偷竊）都必須寫 currency 行。\n`;
            }
        }
        if (!!this.settings?.sendRpgStronghold) {
            const rpg = this.getChat()?.[0]?.horae_meta?.rpg;
            const nodes = rpg?.strongholds || [];
            p += `\n【據點/基地】據點狀態變化時寫（更新/建造/損毀/描述變更），無變化可省略\n`;
            p += `格式: base:據點路徑=等級 或 base:據點路徑|desc=描述\n`;
            p += `路徑用 > 分隔層級\n`;
            p += `示例:\n`;
            p += `  base:主角莊園=3\n`;
            p += `  base:主角莊園>鍛造區>鍛造爐=2\n`;
            p += `  base:主角莊園|desc=坐落於河谷的石砌莊園，配有圍牆和瞭望塔\n`;
            if (nodes.length > 0) {
                const rootNodes = nodes.filter(n => !n.parent);
                const summary = rootNodes.map(r => {
                    const kids = nodes.filter(n => n.parent === r.id);
                    const kidStr = kids.length > 0 ? `(${kids.map(k => k.name).join('、')})` : '';
                    return `${r.name}${r.level != null ? ' Lv.' + r.level : ''}${kidStr}`;
                }).join('；');
                p += `目前據點: ${summary}\n`;
            }
        }
        return p;
    }

    /** 獲取目前對話的裝備配置 */
    _getRpgEquipmentConfig() {
        const rpg = this.getChat()?.[0]?.horae_meta?.rpg;
        return rpg?.equipmentConfig || { locked: false, perChar: {} };
    }

    /** 獲取目前對話的聲望配置 */
    _getRpgReputationConfig() {
        const rpg = this.getChat()?.[0]?.horae_meta?.rpg;
        return rpg?.reputationConfig || { categories: [], _deletedCategories: [] };
    }

    /** 獲取目前對話的貨幣配置 */
    _getRpgCurrencyConfig() {
        const rpg = this.getChat()?.[0]?.horae_meta?.rpg;
        return rpg?.currencyConfig || { denominations: [] };
    }

    /** 動態生成必須包含的標籤提醒（RPG 開啟時追加 <horaerpg>） */
    _generateMustTagsReminder() {
        const tags = ['<horae>...</horae>', '<horaeevent>...</horaeevent>'];
        const rpgActive = this.settings?.rpgMode &&
            (this.settings.sendRpgBars !== false || this.settings.sendRpgSkills !== false ||
             this.settings.sendRpgAttributes !== false || !!this.settings.sendRpgReputation ||
             !!this.settings.sendRpgEquipment || !!this.settings.sendRpgLevel || !!this.settings.sendRpgCurrency ||
             !!this.settings.sendRpgStronghold);
        if (rpgActive) tags.push('<horaerpg>...</horaerpg>');
        const count = tags.length === 2 ? '兩個' : `${tags.length}個`;
        return `你的回覆末尾必須包含 ${tags.join(' 和 ')} ${count}標籤。\n缺少任何一個標籤=不合格。`;
    }

    /** 寬鬆正則解析（不需要標籤包裹） */
    parseLooseFormat(message) {
        const result = {
            timestamp: {},
            costumes: {},
            items: {},
            deletedItems: [],
            events: [],  // 支援多個事件
            affection: {},
            npcs: {},
            scene: {},
            agenda: [],   // 待辦事項
            deletedAgenda: []  // 已完成的待辦事項
        };

        let hasAnyData = false;

        const patterns = {
            time: /time[:：]\s*(.+?)(?:\n|$)/gi,
            location: /location[:：]\s*(.+?)(?:\n|$)/gi,
            atmosphere: /atmosphere[:：]\s*(.+?)(?:\n|$)/gi,
            characters: /characters[:：]\s*(.+?)(?:\n|$)/gi,
            costume: /costume[:：]\s*(.+?)(?:\n|$)/gi,
            item: /item(!{0,2})[:：]\s*(.+?)(?:\n|$)/gi,
            itemDelete: /item-[:：]\s*(.+?)(?:\n|$)/gi,
            event: /event[:：]\s*(.+?)(?:\n|$)/gi,
            affection: /affection[:：]\s*(.+?)(?:\n|$)/gi,
            npc: /npc[:：]\s*(.+?)(?:\n|$)/gi,
            agendaDelete: /agenda-[:：]\s*(.+?)(?:\n|$)/gi,
            agenda: /agenda[:：]\s*(.+?)(?:\n|$)/gi
        };

        // time
        let match;
        while ((match = patterns.time.exec(message)) !== null) {
            const timeStr = match[1].trim();
            const clockMatch = timeStr.match(/\b(\d{1,2}:\d{2})\s*$/);
            if (clockMatch) {
                result.timestamp.story_time = clockMatch[1];
                result.timestamp.story_date = timeStr.substring(0, timeStr.lastIndexOf(clockMatch[1])).trim();
            } else {
                result.timestamp.story_date = timeStr;
                result.timestamp.story_time = '';
            }
            hasAnyData = true;
        }

        // location
        while ((match = patterns.location.exec(message)) !== null) {
            result.scene.location = match[1].trim();
            hasAnyData = true;
        }

        // atmosphere
        while ((match = patterns.atmosphere.exec(message)) !== null) {
            result.scene.atmosphere = match[1].trim();
            hasAnyData = true;
        }

        // characters
        while ((match = patterns.characters.exec(message)) !== null) {
            result.scene.characters_present = match[1].trim().split(/[,，]/).map(c => c.trim()).filter(Boolean);
            hasAnyData = true;
        }

        // costume
        while ((match = patterns.costume.exec(message)) !== null) {
            const costumeStr = match[1].trim();
            const eqIndex = costumeStr.indexOf('=');
            if (eqIndex > 0) {
                const char = costumeStr.substring(0, eqIndex).trim();
                const costume = costumeStr.substring(eqIndex + 1).trim();
                result.costumes[char] = costume;
                hasAnyData = true;
            }
        }

        // item
        while ((match = patterns.item.exec(message)) !== null) {
            const exclamations = match[1] || '';
            const itemStr = match[2].trim();
            let importance = '';  // 一般用空字元串
            if (exclamations === '!!') importance = '!!';  // 關鍵
            else if (exclamations === '!') importance = '!';  // 重要
            
            const eqIndex = itemStr.indexOf('=');
            if (eqIndex > 0) {
                let itemNamePart = itemStr.substring(0, eqIndex).trim();
                const rest = itemStr.substring(eqIndex + 1).trim();
                
                let icon = null;
                let itemName = itemNamePart;
                const emojiMatch = itemNamePart.match(/^([\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F600}-\u{1F64F}]|[\u{1F680}-\u{1F6FF}])/u);
                if (emojiMatch) {
                    icon = emojiMatch[1];
                    itemName = itemNamePart.substring(icon.length).trim();
                }
                
                let description = undefined;  // undefined = 沒有描述資料欄，合併時不覆蓋原有描述
                const pipeIdx = itemName.indexOf('|');
                if (pipeIdx > 0) {
                    const descText = itemName.substring(pipeIdx + 1).trim();
                    if (descText) description = descText;  // 只有非空才設定
                    itemName = itemName.substring(0, pipeIdx).trim();
                }
                
                // 去掉無意義的數量標記
                itemName = itemName.replace(/[\(（]1[\)）]$/, '').trim();
                itemName = itemName.replace(new RegExp(`[\\(（]1[${COUNTING_CLASSIFIERS}][\\)）]$`), '').trim();
                itemName = itemName.replace(new RegExp(`[\\(（][${COUNTING_CLASSIFIERS}][\\)）]$`), '').trim();
                
                const atIndex = rest.indexOf('@');
                const itemInfo = {
                    icon: icon,
                    importance: importance,
                    holder: atIndex >= 0 ? (rest.substring(0, atIndex).trim() || null) : (rest || null),
                    location: atIndex >= 0 ? (rest.substring(atIndex + 1).trim() || '') : ''
                };
                if (description !== undefined) itemInfo.description = description;
                result.items[itemName] = itemInfo;
                hasAnyData = true;
            }
        }

        // item-
        while ((match = patterns.itemDelete.exec(message)) !== null) {
            const itemName = match[1].trim().replace(/^[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u, '').trim();
            if (itemName) {
                result.deletedItems.push(itemName);
                hasAnyData = true;
            }
        }

        // event
        while ((match = patterns.event.exec(message)) !== null) {
            const eventStr = match[1].trim();
            const parts = eventStr.split('|');
            if (parts.length >= 2) {
                const levelRaw = parts[0].trim();
                const summary = parts.slice(1).join('|').trim();
                
                let level = '一般';
                if (levelRaw === '關鍵' || levelRaw.toLowerCase() === 'critical') {
                    level = '關鍵';
                } else if (levelRaw === '重要' || levelRaw.toLowerCase() === 'important') {
                    level = '重要';
                }
                
                result.events.push({
                    is_important: level === '重要' || level === '關鍵',
                    level: level,
                    summary: summary
                });
                hasAnyData = true;
            }
        }

        // affection
        while ((match = patterns.affection.exec(message)) !== null) {
            const affStr = match[1].trim();
            // 絕對值格式
            const absMatch = affStr.match(/^(.+?)=\s*([+\-]?\d+\.?\d*)/);
            if (absMatch) {
                result.affection[absMatch[1].trim()] = { type: 'absolute', value: parseFloat(absMatch[2]) };
                hasAnyData = true;
            } else {
                // 相對值格式 name+/-數值（無=號）
                const relMatch = affStr.match(/^(.+?)([+\-]\d+\.?\d*)/);
                if (relMatch) {
                    result.affection[relMatch[1].trim()] = { type: 'relative', value: relMatch[2] };
                    hasAnyData = true;
                }
            }
        }

        // npc
        while ((match = patterns.npc.exec(message)) !== null) {
            const npcStr = match[1].trim();
            const npcInfo = this._parseNpcFields(npcStr);
            const name = npcInfo._name;
            delete npcInfo._name;
            
            if (name) {
                npcInfo.last_seen = new Date().toISOString();
                result.npcs[name] = npcInfo;
                hasAnyData = true;
            }
        }

        // agenda-:（須在 agenda 之前解析）
        while ((match = patterns.agendaDelete.exec(message)) !== null) {
            const delStr = match[1].trim();
            if (delStr) {
                const pipeIdx = delStr.indexOf('|');
                const text = pipeIdx > 0 ? delStr.substring(pipeIdx + 1).trim() : delStr;
                if (text) {
                    result.deletedAgenda.push(text);
                    hasAnyData = true;
                }
            }
        }

        // agenda
        while ((match = patterns.agenda.exec(message)) !== null) {
            const agendaStr = match[1].trim();
            const pipeIdx = agendaStr.indexOf('|');
            let dateStr = '', text = '';
            if (pipeIdx > 0) {
                dateStr = agendaStr.substring(0, pipeIdx).trim();
                text = agendaStr.substring(pipeIdx + 1).trim();
            } else {
                text = agendaStr;
            }
            if (text) {
                const doneMatch = text.match(/[\(（](完成|已完成|done|finished|completed|失效|取消|已取消)[\)）]\s*$/i);
                if (doneMatch) {
                    const cleanText = text.substring(0, text.length - doneMatch[0].length).trim();
                    if (cleanText) { result.deletedAgenda.push(cleanText); hasAnyData = true; }
                } else {
                    result.agenda.push({ date: dateStr, text, source: 'ai', done: false });
                    hasAnyData = true;
                }
            }
        }

        // 表格更新
        const tableMatches = [...message.matchAll(/<horaetable[:：]\s*(.+?)>([\s\S]*?)<\/horaetable>/gi)];
        if (tableMatches.length > 0) {
            result.tableUpdates = [];
            for (const tm of tableMatches) {
                const tableName = tm[1].trim();
                const tableContent = tm[2].trim();
                const updates = this._parseTableCellEntries(tableContent);
                
                if (Object.keys(updates).length > 0) {
                    result.tableUpdates.push({ name: tableName, updates });
                    hasAnyData = true;
                }
            }
        }

        return hasAnyData ? result : null;
    }
}

// 導出單例
export const horaeManager = new HoraeManager();
