/**
 * Horae - 向量記憶管理器
 * 基於 Transformers.js 的本地向量檢索系統
 *
 * 數據按 chatId 隔離，向量存 IndexedDB，輕量索引存 chat[0].horae_meta.vectorIndex
 */

import { calculateDetailedRelativeTime } from '../utils/timeUtils.js';

const DB_NAME = 'HoraeVectors';
const DB_VERSION = 1;
const STORE_NAME = 'vectors';

const MODEL_CONFIG = {
    'Xenova/bge-small-zh-v1.5': { dimensions: 512, prefix: null },
    'Xenova/multilingual-e5-small': { dimensions: 384, prefix: { query: 'query: ', passage: 'passage: ' } },
};

const TERM_CATEGORIES = {
    medical: ['包紮', '傷口', '治療', '救治', '處理傷', '療傷', '敷藥', '上藥', '受傷', '負傷', '照料', '護理', '急救', '止血', '繃帶', '縫合', '卸甲', '療養', '中毒', '解毒', '昏迷', '甦醒'],
    combat: ['打架', '打鬥', '戰鬥', '衝突', '交手', '攻擊', '擊敗', '斬殺', '對抗', '格鬥', '廝殺', '砍', '劈', '刺', '伏擊', '圍攻', '決鬥', '比武', '防禦', '撤退', '逃跑', '追擊'],
    cooking: ['做飯', '烹飪', '煮', '炒', '烤', '餵食', '吃飯', '喝粥', '餐', '料理', '膳食', '廚房', '食材', '美食', '下廚', '烘焙'],
    clothing: ['換衣', '更衣', '穿衣', '脫衣', '衣物', '換裝', '浴袍', '內衣', '連衣裙', '襯衫'],
    emotion_positive: ['開心', '高興', '快樂', '歡喜', '喜悅', '愉快', '滿足', '感動', '溫馨', '幸福'],
    emotion_negative: ['生氣', '憤怒', '暴怒', '發火', '惱怒', '難過', '傷心', '悲傷', '哭泣', '落淚', '害怕', '恐懼', '驚恐', '委屈', '失落', '焦慮', '羞恥', '愧疚', '崩潰'],
    movement: ['拖', '搬', '抱', '背', '扶', '抬', '推', '拉', '帶走', '轉移', '攙扶', '安頓'],
    social: ['告白', '表白', '道歉', '擁抱', '親吻', '握手', '初次', '重逢', '求婚', '訂婚', '結婚'],
    gift: ['禮物', '贈送', '送給', '信物', '定情', '戒指', '項鍊', '手鍊', '花束', '巧克力', '賀卡', '紀念品', '嫁妝', '聘禮', '徽章', '勳章', '寶石', '收下', '轉贈'],
    ceremony: ['婚禮', '葬禮', '儀式', '典禮', '慶典', '節日', '祭祀', '加冕', '冊封', '宣誓', '洗禮', '成人禮', '畢業', '慶祝', '紀念日', '生日', '週年', '祭典', '開幕', '閉幕', '慶功', '宴會', '舞會'],
    revelation: ['秘密', '真相', '揭露', '坦白', '暴露', '發現', '真實身份', '隱瞞', '謊言', '欺騙', '偽裝', '冒充', '真名', '血統', '身世', '臥底', '間諜', '告密', '揭穿', '拆穿'],
    promise: ['承諾', '誓言', '約定', '保證', '發誓', '立誓', '契約', '盟約', '許諾', '約好', '守護', '效忠', '誓約'],
    loss: ['死亡', '去世', '犧牲', '離別', '分離', '告別', '失去', '消失', '隕落', '凋零', '永別', '喪失', '陣亡', '殉職', '送別', '訣別', '夭折'],
    power: ['覺醒', '更新', '進化', '突破', '衰退', '失去能力', '解封', '封印', '變身', '異變', '獲得力量', '魔力', '能力', '天賦', '血脈', '繼承', '傳承', '修煉', '領悟'],
    intimate: ['親熱', '纏綿', '情事', '春宵', '歡愛', '共度', '同床', '肌膚之親', '親密', '曖昧', '挑逗', '誘惑', '勾引', '撩撥', '調情', '情動', '動情', '慾望', '渴望', '貪戀', '索求', '迎合', '糾纏', '痴纏', '沉淪', '迷戀', '沉溺', '喘息', '顫抖', '呻吟', '嬌喘', '低吟', '求饒', '失控', '隱忍', '剋制', '放縱', '貪婪', '溫存', '餘韻', '繾綣', '旖旎', '性交', '內射', '顏射', '性行為', '中出', '射精', '性器', '交配', '野合', '歡愛', '高潮'],
    body_contact: ['撫摸', '觸碰', '貼近', '依偎', '摟抱', '吻', '啃咬', '舔', '吮', '摩挲', '揉捏', '按壓', '握住', '牽手', '十指相扣', '額頭相抵', '耳鬢廝磨', '臉紅', '心跳', '身體', '肌膚', '鎖骨', '脖頸', '耳垂', '嘴唇', '腰肢', '後背', '髮絲', '指尖', '掌心'],
};

export class VectorManager {
    constructor() {
        this.worker = null;
        this.db = null;
        this.chatId = null;
        this.vectors = new Map();
        this.isReady = false;
        this.isLoading = false;
        this.isApiMode = false;
        this.dimensions = 0;
        this.modelName = '';
        this._apiUrl = '';
        this._apiKey = '';
        this._apiModel = '';
        this.termCounts = new Map();
        this.totalDocuments = 0;
        this._pendingCallbacks = new Map();
        this._callId = 0;
    }

    // ========================================
    // 生命週期
    // ========================================

    async initModel(model, dtype, onProgress) {
        if (this.isLoading) return;
        this.isLoading = true;
        this.isReady = false;
        this.modelName = model;

        try {
            await this._disposeWorker();

            const workerUrl = new URL('../utils/embeddingWorker.js', import.meta.url);
            this.worker = new Worker(workerUrl, { type: 'module' });

            await new Promise((resolve, reject) => {
                const timeout = setTimeout(() => reject(new Error('模型加載超時（5分鐘）')), 300000);

                this.worker.onmessage = (e) => {
                    const { type, data, dimensions: dims } = e.data;
                    if (type === 'progress' && onProgress) {
                        onProgress(data);
                    } else if (type === 'ready') {
                        this.dimensions = dims;
                        this.isReady = true;
                        clearTimeout(timeout);
                        resolve();
                    } else if (type === 'error') {
                        clearTimeout(timeout);
                        reject(new Error(e.data.message));
                    } else if (type === 'result' || type === 'disposed') {
                        const cb = this._pendingCallbacks.get(e.data.id);
                        if (cb) {
                            this._pendingCallbacks.delete(e.data.id);
                            cb.resolve(e.data);
                        }
                    }
                };

                this.worker.onerror = (err) => {
                    clearTimeout(timeout);
                    reject(new Error(err.message || 'Worker 加載失敗'));
                };

                this.worker.postMessage({ type: 'init', data: { model, dtype: dtype || 'q8' } });
            });

            this.worker.onmessage = (e) => {
                const msg = e.data;
                if (msg.type === 'result' || msg.type === 'error' || msg.type === 'disposed') {
                    const cb = this._pendingCallbacks.get(msg.id);
                    if (cb) {
                        this._pendingCallbacks.delete(msg.id);
                        if (msg.type === 'error') cb.reject(new Error(msg.message));
                        else cb.resolve(msg);
                    }
                }
            };

            console.log(`[Horae Vector] 模型已加載: ${model} (${this.dimensions}維)`);
        } finally {
            this.isLoading = false;
        }
    }

    /**
     * 初始化 API 模式（OpenAI 相容的 embedding endpoint）
     */
    async initApi(url, key, model) {
        if (this.isLoading) return;
        this.isLoading = true;
        this.isReady = false;

        try {
            await this._disposeWorker();

            this.isApiMode = true;
            this._apiUrl = url.replace(/\/+$/, '');
            this._apiKey = key;
            this._apiModel = model;
            this.modelName = model;

            // 探測維度：發一條測試文字
            const testResult = await this._embedApi(['test']);
            if (!testResult?.vectors?.[0]) {
                throw new Error('API 連接失敗或返回格式異常，請檢查地址、密鑰和模型名稱是否正確');
            }
            this.dimensions = testResult.vectors[0].length;
            this.isReady = true;
            console.log(`[Horae Vector] API 模式已就緒: ${model} (${this.dimensions}維)`);
        } finally {
            this.isLoading = false;
        }
    }

    async dispose() {
        await this._disposeWorker();
        this.vectors.clear();
        this.termCounts.clear();
        this.totalDocuments = 0;
        this.chatId = null;
        this.isReady = false;
        this.isApiMode = false;
        this._apiUrl = '';
        this._apiKey = '';
        this._apiModel = '';
    }

    async _disposeWorker() {
        if (this.worker) {
            try {
                this.worker.postMessage({ type: 'dispose' });
                await new Promise(r => setTimeout(r, 200));
            } catch (_) { /* ignore */ }
            this.worker.terminate();
            this.worker = null;
        }
        this._pendingCallbacks.clear();
    }

    /**
     * 切換聊天：加載對應 chatId 的向量索引
     */
    async loadChat(chatId, chat) {
        this.chatId = chatId;
        this.vectors.clear();
        this.termCounts.clear();
        this.totalDocuments = 0;

        if (!chatId) return;

        try {
            await this._openDB();
            const stored = await this._loadAllVectors();
            const staleKeys = [];
            for (const item of stored) {
                if (item.messageIndex >= chat.length) {
                    staleKeys.push(item.messageIndex);
                    continue;
                }
                const doc = this.buildVectorDocument(chat[item.messageIndex]?.horae_meta);
                if (doc && this._hashString(doc) !== item.hash) {
                    staleKeys.push(item.messageIndex);
                    continue;
                }
                this.vectors.set(item.messageIndex, {
                    vector: item.vector,
                    hash: item.hash,
                    document: item.document,
                });
                this._updateTermCounts(item.document, 1);
                this.totalDocuments++;
            }
            if (staleKeys.length > 0) {
                for (const idx of staleKeys) await this._deleteVector(idx);
                console.log(`[Horae Vector] 清理了 ${staleKeys.length} 條過期/分支外向量`);
            }
            console.log(`[Horae Vector] 已加載 ${this.vectors.size} 條向量 (chatId: ${chatId})`);
        } catch (err) {
            console.warn('[Horae Vector] 加載向量索引失敗:', err);
        }
    }

    // ========================================
    // 說明文件構建
    // ========================================

    /**
     * 將 horae_meta 序列化為檢索文字
     * 事件摘要為核心（佔主要權重），場景/角色/NPC 為輔
     * 去掉物品、服裝、心情等噪音，讓 embedding 集中在語義關鍵內容
     */
    buildVectorDocument(meta) {
        if (!meta) return '';

        const eventTexts = [];
        if (meta.events?.length > 0) {
            for (const evt of meta.events) {
                if (evt.isSummary || evt.level === '摘要' || evt._summaryId) continue;
                if (evt.summary) eventTexts.push(evt.summary);
            }
        }

        const npcTexts = [];
        if (meta.npcs) {
            for (const [name, info] of Object.entries(meta.npcs)) {
                let s = name;
                if (info.appearance) s += ` ${info.appearance}`;
                if (info.relationship) s += ` ${info.relationship}`;
                npcTexts.push(s);
            }
        }

        if (eventTexts.length === 0 && npcTexts.length === 0) return '';

        const parts = [];

        for (const t of eventTexts) parts.push(t);

        for (const t of npcTexts) parts.push(t);

        if (meta.scene?.location) parts.push(meta.scene.location);

        const chars = meta.scene?.characters_present || [];
        if (chars.length > 0) parts.push(chars.join(' '));

        if (meta.timestamp?.story_date) {
            parts.push(meta.timestamp.story_time
                ? `${meta.timestamp.story_date} ${meta.timestamp.story_time}`
                : meta.timestamp.story_date);
        }

        // RPG milestones: level changes, equipment events, stronghold changes
        const rpg = meta._rpgChanges;
        if (rpg) {
            if (rpg.levels && Object.keys(rpg.levels).length > 0) {
                for (const [owner, lv] of Object.entries(rpg.levels)) {
                    parts.push(`${owner} 更新至Lv.${lv}`);
                }
            }
            for (const eq of (rpg.equipment || [])) {
                parts.push(`${eq.owner} 裝備了 ${eq.name}(${eq.slot})`);
            }
            for (const u of (rpg.unequip || [])) {
                parts.push(`${u.owner} 卸下 ${u.name}(${u.slot})`);
            }
            for (const bc of (rpg.baseChanges || [])) {
                if (bc.field === 'level') parts.push(`據點 ${bc.path} 升至Lv.${bc.value}`);
            }
        }

        return parts.join(' | ');
    }

    // ========================================
    // 索引操作
    // ========================================

    async addMessage(messageIndex, meta) {
        if (!this.isReady || !this.chatId) return;
        if (meta?._skipHorae) return;

        const doc = this.buildVectorDocument(meta);
        if (!doc) return;

        const hash = this._hashString(doc);
        const existing = this.vectors.get(messageIndex);
        if (existing && existing.hash === hash) return;

        const text = this._prepareText(doc, false);
        const result = await this._embed([text]);
        if (!result || !result.vectors?.[0]) return;

        const vector = result.vectors[0];

        if (existing) {
            this._updateTermCounts(existing.document, -1);
        } else {
            this.totalDocuments++;
        }

        this.vectors.set(messageIndex, { vector, hash, document: doc });
        this._updateTermCounts(doc, 1);
        await this._saveVector(messageIndex, { vector, hash, document: doc });
    }

    async removeMessage(messageIndex) {
        const existing = this.vectors.get(messageIndex);
        if (!existing) return;

        this._updateTermCounts(existing.document, -1);
        this.totalDocuments--;
        this.vectors.delete(messageIndex);
        await this._deleteVector(messageIndex);
    }

    /**
     * 批量建索引（用於歷史記錄）
     * @returns {{ indexed: number, skipped: number }}
     */
    async batchIndex(chat, onProgress) {
        if (!this.isReady || !this.chatId) return { indexed: 0, skipped: 0 };

        const tasks = [];
        for (let i = 0; i < chat.length; i++) {
            const meta = chat[i].horae_meta;
            if (!meta || chat[i].is_user) continue;
            if (meta._skipHorae) continue;
            const doc = this.buildVectorDocument(meta);
            if (!doc) continue;
            const hash = this._hashString(doc);
            const existing = this.vectors.get(i);
            if (existing && existing.hash === hash) continue;
            tasks.push({ messageIndex: i, document: doc, hash });
        }

        if (tasks.length === 0) return { indexed: 0, skipped: chat.length };

        const batchSize = this.isApiMode ? 8 : 16;
        let indexed = 0;

        for (let b = 0; b < tasks.length; b += batchSize) {
            const batch = tasks.slice(b, b + batchSize);
            const texts = batch.map(t => this._prepareText(t.document, false));
            const result = await this._embed(texts);
            if (!result?.vectors) continue;

            for (let j = 0; j < batch.length; j++) {
                const task = batch[j];
                const vector = result.vectors[j];
                if (!vector) continue;

                const old = this.vectors.get(task.messageIndex);
                if (old) {
                    this._updateTermCounts(old.document, -1);
                } else {
                    this.totalDocuments++;
                }

                this.vectors.set(task.messageIndex, {
                    vector,
                    hash: task.hash,
                    document: task.document,
                });
                this._updateTermCounts(task.document, 1);
                await this._saveVector(task.messageIndex, { vector, hash: task.hash, document: task.document });
                indexed++;
            }

            if (onProgress) {
                onProgress({ current: Math.min(b + batchSize, tasks.length), total: tasks.length });
            }
        }

        return { indexed, skipped: chat.length - tasks.length };
    }

    async clearIndex() {
        this.vectors.clear();
        this.termCounts.clear();
        this.totalDocuments = 0;
        if (this.chatId) await this._clearVectors();
    }

    // ========================================
    // 查詢與召回
    // ========================================

    /**
     * 構建狀態查詢文字（目前場景/角色/事件）
     */
    buildStateQuery(currentState, lastMeta) {
        const parts = [];

        if (currentState.scene?.location) parts.push(currentState.scene.location);

        const chars = currentState.scene?.characters_present || [];
        for (const c of chars) {
            parts.push(c);
            if (currentState.costumes?.[c]) parts.push(currentState.costumes[c]);
        }

        if (lastMeta?.events?.length > 0) {
            for (const evt of lastMeta.events) {
                if (evt.summary) parts.push(evt.summary);
            }
        }

        return parts.filter(Boolean).join(' ');
    }

    /**
     * 清理用戶訊息為查詢文字
     */
    cleanUserMessage(rawMessage) {
        if (!rawMessage) return '';
        return rawMessage
            .replace(/<[^>]*>/g, '')
            .replace(/[\[\]]/g, '')
            .trim()
            .substring(0, 300);
    }

    /**
     * 向量檢索
     * @param {string} queryText
     * @param {number} topK
     * @param {number} threshold
     * @param {Set<number>} excludeIndices - 排除的訊息索引（已在上下文中）
     * @returns {Promise<Array<{messageIndex: number, similarity: number, document: string}>>}
     */
    async search(queryText, topK = 5, threshold = 0.72, excludeIndices = new Set(), pureMode = false) {
        if (!this.isReady || !queryText || this.vectors.size === 0) return [];

        const prepared = this._prepareText(queryText, true);
        console.log('[Horae Vector] 開始 embedding 查詢...');
        const result = await this._embed([prepared]);
        if (!result?.vectors?.[0]) {
            console.warn('[Horae Vector] embedding 返回空結果:', result);
            return [];
        }

        const queryVec = result.vectors[0];
        console.log(`[Horae Vector] 查詢向量維度: ${queryVec.length}，開始對比 ${this.vectors.size} 條...`);

        const scored = [];
        const allScored = [];
        let searchedCount = 0;

        for (const [msgIdx, entry] of this.vectors) {
            if (excludeIndices.has(msgIdx)) continue;
            searchedCount++;
            const sim = this._dotProduct(queryVec, entry.vector);
            allScored.push({ messageIndex: msgIdx, similarity: sim, document: entry.document });
            if (sim >= threshold) {
                scored.push({ messageIndex: msgIdx, similarity: sim, document: entry.document });
            }
        }

        allScored.sort((a, b) => b.similarity - a.similarity);
        const bestSim = allScored.length > 0 ? allScored[0].similarity : 0;
        console.log(`[Horae Vector] 搜索了 ${searchedCount} 條 | 最高相似度=${bestSim.toFixed(4)} | 超過閾值(${threshold}): ${scored.length} 條`);
        if (scored.length === 0 && allScored.length > 0) {
            console.log(`[Horae Vector] 閾值下 Top-5 候選:`);
            for (const c of allScored.slice(0, 5)) {
                console.log(`  #${c.messageIndex} sim=${c.similarity.toFixed(4)} | ${c.document.substring(0, 60)}`);
            }
        }

        scored.sort((a, b) => b.similarity - a.similarity);

        const adjusted = pureMode ? scored : this._adjustThresholdByFrequency(scored, threshold);
        if (!pureMode) console.log(`[Horae Vector] 頻率過濾後: ${adjusted.length} 條`);

        const deduped = this._deduplicateResults(adjusted);
        console.log(`[Horae Vector] 去重後: ${deduped.length} 條`);

        return deduped.slice(0, topK);
    }

    /**
     * 策略B：高頻內容懲罰
     * 只在說明文件中 >80% 的詞都是公共詞（出現在 >60% 說明文件中）時才輕微提高閾值，
     * 避免角色名等必然高頻詞誤殺有效結果。
     */
    _adjustThresholdByFrequency(results, baseThreshold) {
        if (results.length < 2 || this.totalDocuments < 10) return results;

        return results.filter(r => {
            const terms = this._extractKeyTerms(r.document);
            if (terms.length === 0) return true;

            let commonCount = 0;
            for (const term of terms) {
                const count = this.termCounts.get(term) || 0;
                if (count / this.totalDocuments > 0.6) commonCount++;
            }
            const commonRatio = commonCount / terms.length;

            if (commonRatio > 0.8) {
                const penalty = (commonRatio - 0.8) * 0.1;
                return r.similarity >= baseThreshold + penalty;
            }
            return true;
        });
    }

    /**
     * 策略C：摺疊高度相似的結果
     */
    _deduplicateResults(results) {
        if (results.length <= 1) return results;

        const kept = [results[0]];
        for (let i = 1; i < results.length; i++) {
            const candidate = results[i];
            let isDuplicate = false;
            for (const existing of kept) {
                const mutualSim = this._dotProduct(
                    this.vectors.get(existing.messageIndex)?.vector || [],
                    this.vectors.get(candidate.messageIndex)?.vector || []
                );
                if (mutualSim > 0.92) {
                    isDuplicate = true;
                    break;
                }
            }
            if (!isDuplicate) kept.push(candidate);
        }
        return kept;
    }

    // ========================================
    // 召回 Prompt 構建
    // ========================================

    /**
     * 智慧召回：結構化查詢 + 向量搜索並行，合併結果
     */
    async generateRecallPrompt(horaeManager, skipLast, settings) {
        const chat = horaeManager.getChat();
        const state = horaeManager.getLatestState(skipLast);
        const topK = settings.vectorTopK || 5;
        const threshold = settings.vectorThreshold ?? 0.72;

        let rawUserMsg = '';
        for (let i = chat.length - 1; i >= 0; i--) {
            if (chat[i].is_user) { rawUserMsg = chat[i].mes || ''; break; }
        }
        const userQuery = this.cleanUserMessage(rawUserMsg);

        const EXCLUDE_RECENT = 5;
        const excludeIndices = new Set();
        for (let i = Math.max(0, chat.length - EXCLUDE_RECENT); i < chat.length; i++) {
            excludeIndices.add(i);
        }

        const merged = new Map();

        const pureMode = !!settings.vectorPureMode;
        if (pureMode) console.log('[Horae Vector] 純向量模式已啟用，跳過關鍵詞啟發式');

        const structuredResults = this._structuredQuery(userQuery, chat, state, excludeIndices, topK, pureMode);
        console.log(`[Horae Vector] 結構化查詢: ${structuredResults.length} 條命中`);
        for (const r of structuredResults) {
            merged.set(r.messageIndex, r);
        }

        const hybridResults = await this._hybridSearch(userQuery, state, horaeManager, skipLast, settings, excludeIndices, topK, threshold, pureMode);
        console.log(`[Horae Vector] 向量混合搜索: ${hybridResults.length} 條命中`);
        for (const r of hybridResults) {
            if (!merged.has(r.messageIndex)) {
                merged.set(r.messageIndex, r);
            }
        }

        // 多人卡角色相關性加權：
        // 收集"相關角色" = 用戶訊息中提到的角色 + 目前在場角色
        // 對涉及相關角色的結果施加小幅正向加權，優先召回相關事件
        // 不過濾任何結果，確保跨角色引用（如向A提起B）仍能召回
        const relevantChars = new Set(state.scene?.characters_present || []);
        const allKnownChars = new Set();
        for (let i = 0; i < chat.length; i++) {
            const m = chat[i].horae_meta;
            if (!m) continue;
            (m.scene?.characters_present || []).forEach(c => allKnownChars.add(c));
            if (m.npcs) Object.keys(m.npcs).forEach(c => allKnownChars.add(c));
        }
        for (const c of allKnownChars) {
            if (userQuery && userQuery.includes(c)) relevantChars.add(c);
        }

        let results = Array.from(merged.values());
        if (relevantChars.size > 0) {
            for (const r of results) {
                const meta = chat[r.messageIndex]?.horae_meta;
                if (!meta) continue;
                const docChars = new Set([
                    ...(meta.scene?.characters_present || []),
                    ...Object.keys(meta.npcs || {}),
                ]);
                let hasRelevant = false;
                for (const c of relevantChars) {
                    if (docChars.has(c)) { hasRelevant = true; break; }
                }
                if (hasRelevant) {
                    r.similarity += 0.03;
                }
            }
            console.log(`[Horae Vector] 角色加權: 相關角色=[${[...relevantChars].join(',')}]`);
        }

        results.sort((a, b) => b.similarity - a.similarity);

        // Rerank：對候選結果做二次精排
        if (settings.vectorRerankEnabled && settings.vectorRerankModel && results.length > 1) {
            const rerankCandidates = results.slice(0, topK * 3);
            const rerankQuery = userQuery || this.buildStateQuery(state, null);
            if (rerankQuery) {
                try {
                    const useFullText = !!settings.vectorRerankFullText;
                    const _stripTags = settings.vectorStripTags || '';
                    const rerankDocs = rerankCandidates.map(r => {
                        if (useFullText) {
                            const fullText = this._extractCleanText(chat[r.messageIndex]?.mes, _stripTags);
                            return fullText || r.document;
                        }
                        return r.document;
                    });
                    console.log(`[Horae Vector] Rerank 模式: ${useFullText ? '全文精排' : '摘要排序'}`);

                    const reranked = await this._rerank(
                        rerankQuery,
                        rerankDocs,
                        topK,
                        settings
                    );
                    if (reranked && reranked.length > 0) {
                        console.log(`[Horae Vector] Rerank 完成: ${reranked.length} 條`);
                        results = reranked.map(rr => {
                            const original = rerankCandidates[rr.index];
                            return {
                                ...original,
                                similarity: rr.relevance_score,
                                source: original.source + (useFullText ? '+rerank-full' : '+rerank'),
                            };
                        });
                    }
                } catch (err) {
                    console.warn('[Horae Vector] Rerank 失敗，使用原始排序:', err.message);
                }
            }
        }

        results = results.slice(0, topK);

        console.log(`[Horae Vector] === 最終合併: ${results.length} 條 ===`);
        for (const r of results) {
            console.log(`  #${r.messageIndex} sim=${r.similarity.toFixed(3)} [${r.source}]`);
        }

        if (results.length === 0) return '';

        const currentDate = state.timestamp?.story_date;
        const fullTextCount = Math.min(settings.vectorFullTextCount ?? 3, topK);
        const fullTextThreshold = settings.vectorFullTextThreshold ?? 0.9;
        const recallText = this._buildRecallText(results, currentDate, chat, fullTextCount, fullTextThreshold, settings.vectorStripTags || '');
        console.log(`[Horae Vector] 召回文字 (${recallText.length}字):\n${recallText}`);
        return recallText;
    }

    // ========================================
    // 結構化查詢（精準，不需要向量）
    // ========================================

    /**
     * 從用戶訊息解析意圖，直接查詢 horae_meta 結構化數據
     */
    _structuredQuery(userQuery, chat, state, excludeIndices, topK, pureMode = false) {
        if (!userQuery || chat.length === 0) return [];

        const knownChars = new Set();
        for (let i = 0; i < chat.length; i++) {
            const m = chat[i].horae_meta;
            if (!m) continue;
            (m.scene?.characters_present || []).forEach(c => knownChars.add(c));
            if (m.npcs) Object.keys(m.npcs).forEach(c => knownChars.add(c));
        }

        const mentionedChars = [];
        for (const c of knownChars) {
            if (userQuery.includes(c)) mentionedChars.push(c);
        }

        const isFirst = /第一次|初次|首次|初见|初遇|最早|一开始/.test(userQuery);
        const isLast = /上次|上一次|最后一次|最近一次|之前/.test(userQuery);

        const hasCostumeKw = /穿|戴|换|衣|裙|裤|袍|衫|装|鞋/.test(userQuery);
        const hasMoodKw = /生气|愤怒|开心|高兴|难过|伤心|哭|害怕|恐惧|害羞|羞耻|得意|满足|嫉妒|悲伤|焦虑|紧张|兴奋|感动|温柔|冷漠/.test(userQuery);
        const hasGiftKw = /礼物|赠送|送给|送的|信物|定情|收到|收下|转赠|聘礼|嫁妆|纪念品|贺卡/.test(userQuery);
        const hasImportantItemKw = /重要.{0,2}(物品|东西|道具|宝物)|关键.{0,2}(物品|东西|道具|宝物)|珍贵|宝贝|宝物|神器|秘宝|圣物/.test(userQuery);
        const hasImportantEventKw = /重要.{0,2}(事|事件|经历)|关键.{0,2}(事|事件|转折)|大事|转折|里程碑/.test(userQuery);
        const hasCeremonyKw = /婚礼|葬礼|仪式|典礼|庆典|节日|祭祀|加冕|册封|宣誓|洗礼|成人礼|庆祝|宴会|舞会|祭典/.test(userQuery);
        const hasPromiseKw = /承诺|誓言|约定|保证|发誓|立誓|契约|盟约|许诺/.test(userQuery);
        const hasLossKw = /死亡|去世|牺牲|离别|分离|告别|失去|消失|陨落|永别|诀别|阵亡/.test(userQuery);
        const hasRevelationKw = /秘密|真相|揭露|坦白|暴露|真实身份|隐瞒|谎言|欺骗|伪装|冒充|真名|血统|身世|揭穿/.test(userQuery);
        const hasPowerKw = /觉醒|升级|进化|突破|衰退|失去能力|解封|封印|变身|异变|获得力量|血脉|继承|传承|领悟/.test(userQuery);

        const results = [];

        if (isFirst && mentionedChars.length > 0) {
            for (const charName of mentionedChars) {
                const idx = this._findFirstAppearance(chat, charName, excludeIndices);
                if (idx !== -1) {
                    results.push({ messageIndex: idx, similarity: 1.0, document: `[結構化] ${charName}首次出現`, source: 'structured' });
                    console.log(`[Horae Vector] 結構化查詢: "${charName}" 首次出現於 #${idx}`);
                }
            }
        }

        if (isLast && mentionedChars.length > 0 && hasCostumeKw) {
            const costumeKw = this._extractCostumeKeywords(userQuery, mentionedChars);
            if (costumeKw) {
                for (const charName of mentionedChars) {
                    const idx = this._findLastCostume(chat, charName, costumeKw, excludeIndices);
                    if (idx !== -1) {
                        results.push({ messageIndex: idx, similarity: 1.0, document: `[結構化] ${charName}穿${costumeKw}`, source: 'structured' });
                        console.log(`[Horae Vector] 結構化查詢: "${charName}" 上次穿 "${costumeKw}" 於 #${idx}`);
                    }
                }
            }
        }

        if (hasCostumeKw && !isFirst && !isLast && mentionedChars.length === 0) {
            const costumeKw = this._extractCostumeKeywords(userQuery, []);
            if (costumeKw) {
                const matches = this._findCostumeMatches(chat, costumeKw, excludeIndices, topK);
                for (const m of matches) {
                    results.push({ messageIndex: m.idx, similarity: 0.95, document: `[結構化] 服裝配對:${costumeKw}`, source: 'structured' });
                }
            }
        }

        if (isLast && hasMoodKw) {
            const moodKw = this._extractMoodKeyword(userQuery);
            if (moodKw) {
                const targetChar = mentionedChars[0] || null;
                const idx = this._findLastMood(chat, targetChar, moodKw, excludeIndices);
                if (idx !== -1) {
                    results.push({ messageIndex: idx, similarity: 1.0, document: `[結構化] 情緒配對:${moodKw}`, source: 'structured' });
                    console.log(`[Horae Vector] 結構化查詢: 上次 "${moodKw}" 於 #${idx}`);
                }
            }
        }

        if (hasGiftKw) {
            const giftResults = this._findGiftItems(chat, mentionedChars, excludeIndices, topK);
            for (const r of giftResults) {
                results.push(r);
                console.log(`[Horae Vector] 結構化查詢: 禮物/贈品 #${r.messageIndex} [${r.document}]`);
            }
        }

        if (hasImportantItemKw) {
            const impResults = this._findImportantItems(chat, excludeIndices, topK);
            for (const r of impResults) {
                results.push(r);
                console.log(`[Horae Vector] 結構化查詢: 重要物品 #${r.messageIndex} [${r.document}]`);
            }
        }

        if (hasImportantEventKw) {
            const evtResults = this._findImportantEvents(chat, excludeIndices, topK);
            for (const r of evtResults) {
                results.push(r);
                console.log(`[Horae Vector] 結構化查詢: 重要事件 #${r.messageIndex} [${r.document}]`);
            }
        }

        // 純向量模式下跳過關鍵詞啟發式（主題事件搜索、事件詞組配對），完全依賴向量語義
        if (!pureMode) {
            if (hasCeremonyKw || hasPromiseKw || hasLossKw || hasRevelationKw || hasPowerKw) {
                const thematicResults = this._findThematicEvents(chat, {
                    ceremony: hasCeremonyKw, promise: hasPromiseKw,
                    loss: hasLossKw, revelation: hasRevelationKw, power: hasPowerKw,
                }, excludeIndices, topK);
                for (const r of thematicResults) {
                    results.push(r);
                    console.log(`[Horae Vector] 結構化查詢: 主題事件 #${r.messageIndex} [${r.document}]`);
                }
            }

            const existingIds = new Set(results.map(r => r.messageIndex));
            const eventMatches = this._eventKeywordSearch(userQuery, chat, mentionedChars, existingIds, excludeIndices, topK);
            for (const m of eventMatches) {
                results.push(m);
            }
        }

        const withContext = this._expandContextWindow(results, chat, excludeIndices);
        return withContext.slice(0, topK);
    }

    /**
     * 上下文視窗擴展：對每個命中訊息，把前後相鄰的 AI 訊息也加進來
     * RP 中相鄰訊息是連續事件，天然相關
     */
    _expandContextWindow(results, chat, excludeIndices) {
        const resultIds = new Set(results.map(r => r.messageIndex));
        const contextToAdd = [];

        for (const r of results) {
            const idx = r.messageIndex;

            for (let i = idx - 1; i >= Math.max(0, idx - 3); i--) {
                if (excludeIndices.has(i) || resultIds.has(i)) continue;
                const m = chat[i].horae_meta;
                if (!chat[i].is_user && this._hasOriginalEvents(m)) {
                    contextToAdd.push({
                        messageIndex: i,
                        similarity: r.similarity * 0.85,
                        document: `[上文] #${idx}的前置事件`,
                        source: 'context',
                    });
                    resultIds.add(i);
                    break;
                }
            }

            for (let i = idx + 1; i <= Math.min(chat.length - 1, idx + 3); i++) {
                if (excludeIndices.has(i) || resultIds.has(i)) continue;
                const m = chat[i].horae_meta;
                if (!chat[i].is_user && this._hasOriginalEvents(m)) {
                    contextToAdd.push({
                        messageIndex: i,
                        similarity: r.similarity * 0.85,
                        document: `[下文] #${idx}的後續事件`,
                        source: 'context',
                    });
                    resultIds.add(i);
                    break;
                }
            }
        }

        if (contextToAdd.length > 0) {
            console.log(`[Horae Vector] 上下文擴展: +${contextToAdd.length} 條`);
            for (const c of contextToAdd) console.log(`  #${c.messageIndex} [${c.document}]`);
        }

        const all = [...results, ...contextToAdd];
        all.sort((a, b) => b.similarity - a.similarity);
        return all;
    }

    /**
     * 事件關鍵詞搜索：從用戶文字直接掃描已知類別詞彙，擴展後搜索事件摘要
     */
    _eventKeywordSearch(userQuery, chat, mentionedChars, skipIds, excludeIndices, limit) {
        const detected = this._detectCategoryTerms(userQuery);
        if (detected.length === 0) return [];

        const expanded = this._expandByCategory(detected);
        console.log(`[Horae Vector] 事件搜索: 檢測到=[${detected.join(',')}] 擴展後=[${expanded.join(',')}]`);

        const scored = [];
        for (let i = 0; i < chat.length; i++) {
            if (excludeIndices.has(i) || skipIds.has(i)) continue;
            const meta = chat[i].horae_meta;
            if (!meta) continue;

            const searchText = this._buildSearchableText(meta);
            if (!searchText) continue;

            let matchCount = 0;
            const matched = [];
            for (const kw of expanded) {
                if (searchText.includes(kw)) {
                    matchCount++;
                    matched.push(kw);
                }
            }

            if (matchCount >= 2 || (matchCount >= 1 && mentionedChars.some(c => searchText.includes(c)))) {
                scored.push({
                    messageIndex: i,
                    similarity: 0.85 + matchCount * 0.02,
                    document: `[事件配對] ${matched.join(',')}`,
                    source: 'structured',
                    _matchCount: matchCount,
                });
            }
        }

        scored.sort((a, b) => b._matchCount - a._matchCount || b.similarity - a.similarity);
        const top = scored.slice(0, limit);
        if (top.length > 0) {
            console.log(`[Horae Vector] 事件搜索命中 ${top.length} 條:`);
            for (const r of top) console.log(`  #${r.messageIndex} matches=${r._matchCount} [${r.document}]`);
        }
        return top;
    }

    _buildSearchableText(meta) {
        const parts = [];
        if (meta.events) {
            for (const evt of meta.events) {
                if (evt.isSummary || evt.level === '摘要' || evt._summaryId) continue;
                if (evt.summary) parts.push(evt.summary);
            }
        }
        if (meta.scene?.location) parts.push(meta.scene.location);
        if (meta.npcs) {
            for (const [name, info] of Object.entries(meta.npcs)) {
                parts.push(name);
                if (info.description) parts.push(info.description);
            }
        }
        if (meta.items) {
            for (const [name, info] of Object.entries(meta.items)) {
                parts.push(name);
                if (info.location) parts.push(info.location);
            }
        }
        return parts.join(' ');
    }

    /**
     * 直接從用戶文字中掃描 TERM_CATEGORIES 中的已知詞彙（無需分詞）
     */
    _detectCategoryTerms(text) {
        const found = [];
        for (const terms of Object.values(TERM_CATEGORIES)) {
            for (const term of terms) {
                if (text.includes(term)) {
                    found.push(term);
                }
            }
        }
        return [...new Set(found)];
    }

    /**
     * 將檢測到的詞擴展到同類別的所有詞
     */
    _expandByCategory(keywords) {
        const expanded = new Set(keywords);
        for (const kw of keywords) {
            for (const terms of Object.values(TERM_CATEGORIES)) {
                if (terms.includes(kw)) {
                    for (const t of terms) expanded.add(t);
                }
            }
        }
        return [...expanded];
    }

    _findFirstAppearance(chat, charName, excludeIndices) {
        for (let i = 0; i < chat.length; i++) {
            if (excludeIndices.has(i)) continue;
            const m = chat[i].horae_meta;
            if (!m) continue;
            if (m.npcs && m.npcs[charName]) return i;
            if (m.scene?.characters_present?.includes(charName)) return i;
        }
        return -1;
    }

    _findLastCostume(chat, charName, costumeKw, excludeIndices) {
        for (let i = chat.length - 1; i >= 0; i--) {
            if (excludeIndices.has(i)) continue;
            const costume = chat[i].horae_meta?.costumes?.[charName];
            if (costume && costume.includes(costumeKw)) return i;
        }
        return -1;
    }

    _findCostumeMatches(chat, costumeKw, excludeIndices, limit) {
        const matches = [];
        for (let i = chat.length - 1; i >= 0 && matches.length < limit; i--) {
            if (excludeIndices.has(i)) continue;
            const costumes = chat[i].horae_meta?.costumes;
            if (!costumes) continue;
            for (const v of Object.values(costumes)) {
                if (v && v.includes(costumeKw)) { matches.push({ idx: i }); break; }
            }
        }
        return matches;
    }

    _findLastMood(chat, charName, moodKw, excludeIndices) {
        for (let i = chat.length - 1; i >= 0; i--) {
            if (excludeIndices.has(i)) continue;
            const mood = chat[i].horae_meta?.mood;
            if (!mood) continue;
            if (charName) {
                if (mood[charName] && mood[charName].includes(moodKw)) return i;
            } else {
                for (const v of Object.values(mood)) {
                    if (v && v.includes(moodKw)) return i;
                }
            }
        }
        return -1;
    }

    _extractCostumeKeywords(query, chars) {
        let cleaned = query;
        for (const c of chars) cleaned = cleaned.replace(c, '');
        cleaned = cleaned.replace(/上次|上一次|最后一次|之前|穿|戴|换|的|了|过|着|那件|那套|那个/g, '').trim();
        return cleaned.length >= 2 ? cleaned : '';
    }

    _extractMoodKeyword(query) {
        const moodWords = ['生氣', '憤怒', '開心', '高興', '難過', '傷心', '哭泣', '害怕', '恐懼', '害羞', '羞恥', '得意', '滿足', '嫉妒', '悲傷', '焦慮', '緊張', '興奮', '感動', '溫柔', '冷漠', '暴怒', '委屈', '失落'];
        for (const w of moodWords) {
            if (query.includes(w)) return w;
        }
        return '';
    }

    /**
     * 搜尋與禮物/贈品相關的訊息
     * 透過 item.holder 變化或事件文字中的贈送關鍵詞定位
     */
    _findGiftItems(chat, mentionedChars, excludeIndices, limit) {
        const giftKws = ['贈送', '送給', '收到', '收下', '轉贈', '信物', '定情', '禮物', '聘禮', '嫁妝'];
        const results = [];
        const seen = new Set();

        for (let i = chat.length - 1; i >= 0 && results.length < limit; i--) {
            if (excludeIndices.has(i) || seen.has(i)) continue;
            const meta = chat[i].horae_meta;
            if (!meta) continue;

            let matched = false;
            const matchedItems = [];

            if (meta.items) {
                for (const [name, info] of Object.entries(meta.items)) {
                    const imp = info.importance || '';
                    const holder = info.holder || '';
                    const holderMatchesChar = mentionedChars.length === 0 || mentionedChars.some(c => holder.includes(c));

                    if ((imp === '!' || imp === '!!') && holderMatchesChar) {
                        matched = true;
                        matchedItems.push(`${imp === '!!' ? '关键' : '重要'}:${name}`);
                    }
                }
            }

            if (!matched && meta.events) {
                for (const evt of meta.events) {
                    if (evt.isSummary || evt.level === '摘要' || evt._summaryId) continue;
                    const text = evt.summary || '';
                    if (giftKws.some(kw => text.includes(kw))) {
                        if (mentionedChars.length === 0 || mentionedChars.some(c => text.includes(c))) {
                            matched = true;
                            matchedItems.push(text.substring(0, 20));
                        }
                    }
                }
            }

            if (matched) {
                seen.add(i);
                results.push({
                    messageIndex: i,
                    similarity: 0.95,
                    document: `[結構化] 禮物/贈品: ${matchedItems.join('; ')}`,
                    source: 'structured',
                });
            }
        }
        return results;
    }

    /**
     * 搜尋包含重要/關鍵物品的訊息（importance '!' 或 '!!'）
     */
    _findImportantItems(chat, excludeIndices, limit) {
        const results = [];
        for (let i = chat.length - 1; i >= 0 && results.length < limit; i--) {
            if (excludeIndices.has(i)) continue;
            const meta = chat[i].horae_meta;
            if (!meta?.items) continue;

            const importantNames = [];
            for (const [name, info] of Object.entries(meta.items)) {
                if (info.importance === '!' || info.importance === '!!') {
                    importantNames.push(`${info.importance === '!!' ? '★' : '☆'}${info.icon || ''}${name}`);
                }
            }
            if (importantNames.length > 0) {
                results.push({
                    messageIndex: i,
                    similarity: 0.95,
                    document: `[結構化] 重要物品: ${importantNames.join(', ')}`,
                    source: 'structured',
                });
            }
        }
        return results;
    }

    /**
     * 搜尋重要/關鍵層級的事件
     */
    _findImportantEvents(chat, excludeIndices, limit) {
        const results = [];
        for (let i = chat.length - 1; i >= 0 && results.length < limit; i--) {
            if (excludeIndices.has(i)) continue;
            const meta = chat[i].horae_meta;
            if (!meta?.events) continue;

            for (const evt of meta.events) {
                if (evt.isSummary || evt.level === '摘要' || evt._summaryId) continue;
                if (evt.level === '重要' || evt.level === '關鍵') {
                    results.push({
                        messageIndex: i,
                        similarity: evt.level === '關鍵' ? 1.0 : 0.95,
                        document: `[結構化] ${evt.level}事件: ${(evt.summary || '').substring(0, 30)}`,
                        source: 'structured',
                    });
                    break;
                }
            }
        }
        return results;
    }

    /**
     * 主題事件搜索：儀式/承諾/失去/揭露/能力變化
     * 結合事件文字和 TERM_CATEGORIES 做精準配對
     */
    _findThematicEvents(chat, flags, excludeIndices, limit) {
        const activeCategories = [];
        if (flags.ceremony) activeCategories.push('ceremony');
        if (flags.promise) activeCategories.push('promise');
        if (flags.loss) activeCategories.push('loss');
        if (flags.revelation) activeCategories.push('revelation');
        if (flags.power) activeCategories.push('power');

        const searchTerms = new Set();
        for (const cat of activeCategories) {
            if (TERM_CATEGORIES[cat]) {
                for (const t of TERM_CATEGORIES[cat]) searchTerms.add(t);
            }
        }
        if (searchTerms.size === 0) return [];

        const results = [];
        for (let i = chat.length - 1; i >= 0 && results.length < limit; i--) {
            if (excludeIndices.has(i)) continue;
            const meta = chat[i].horae_meta;
            if (!meta?.events) continue;

            for (const evt of meta.events) {
                if (evt.isSummary || evt.level === '摘要' || evt._summaryId) continue;
                const text = evt.summary || '';
                const hits = [...searchTerms].filter(t => text.includes(t));
                if (hits.length > 0) {
                    results.push({
                        messageIndex: i,
                        similarity: 0.90 + Math.min(hits.length, 5) * 0.02,
                        document: `[結構化] 主題事件(${activeCategories.join('+')}): ${hits.join(',')}`,
                        source: 'structured',
                    });
                    break;
                }
            }
        }
        return results;
    }

    // ========================================
    // 向量+關鍵詞混合搜索（兜底）
    // ========================================

    async _hybridSearch(userQuery, state, horaeManager, skipLast, settings, excludeIndices, topK, threshold, pureMode = false) {
        if (!this.isReady || this.vectors.size === 0) return [];

        const lastIdx = Math.max(0, horaeManager.getChat().length - 1 - skipLast);
        const lastMeta = horaeManager.getMessageMeta(lastIdx);
        const stateQuery = this.buildStateQuery(state, lastMeta);

        const merged = new Map();

        if (userQuery) {
            const intentThreshold = Math.max(threshold - 0.25, 0.4);
            const intentResults = await this.search(userQuery, topK * 2, intentThreshold, excludeIndices, pureMode);
            console.log(`[Horae Vector] 意圖搜索: ${intentResults.length} 條`);
            for (const r of intentResults) {
                merged.set(r.messageIndex, { ...r, source: 'intent' });
            }
        }

        if (stateQuery) {
            const stateResults = await this.search(stateQuery, topK * 2, threshold, excludeIndices, pureMode);
            console.log(`[Horae Vector] 狀態搜索: ${stateResults.length} 條`);
            for (const r of stateResults) {
                const existing = merged.get(r.messageIndex);
                if (!existing || r.similarity > existing.similarity) {
                    merged.set(r.messageIndex, { ...r, source: existing ? 'both' : 'state' });
                }
            }
        }

        let results = Array.from(merged.values());
        results.sort((a, b) => b.similarity - a.similarity);
        results = this._deduplicateResults(results).slice(0, topK);

        console.log(`[Horae Vector] 混合搜索結果: ${results.length} 條`);
        for (const r of results) {
            console.log(`  #${r.messageIndex} sim=${r.similarity.toFixed(4)} [${r.source}] | ${r.document.substring(0, 80)}`);
        }

        return results;
    }

    _buildRecallText(results, currentDate, chat, fullTextCount = 3, fullTextThreshold = 0.9, stripTags = '') {
        const lines = ['[記憶回溯——以下為與目前情境相關的歷史片段，僅供參考，非目前上下文]'];

        for (let rank = 0; rank < results.length; rank++) {
            const r = results[rank];
            const meta = chat[r.messageIndex]?.horae_meta;
            if (!meta) continue;

            const isFullText = fullTextCount > 0 && rank < fullTextCount && r.similarity >= fullTextThreshold;

            if (isFullText) {
                const rawText = this._extractCleanText(chat[r.messageIndex]?.mes, stripTags);
                if (rawText) {
                    const timeTag = this._buildTimeTag(meta?.timestamp, currentDate);
                    lines.push(`#${r.messageIndex} ${timeTag ? timeTag + ' ' : ''}[全文回顧]\n${rawText}`);
                    continue;
                }
            }

            const parts = [];

            const timeTag = this._buildTimeTag(meta?.timestamp, currentDate);
            if (timeTag) parts.push(timeTag);

            if (meta?.scene?.location) parts.push(`場景:${meta.scene.location}`);

            const chars = meta?.scene?.characters_present || [];
            const costumes = meta?.costumes || {};
            for (const c of chars) {
                parts.push(costumes[c] ? `${c}(${costumes[c]})` : c);
            }

            if (meta?.events?.length > 0) {
                for (const evt of meta.events) {
                    if (evt.isSummary || evt.level === '摘要') continue;
                    const mark = evt.level === '關鍵' ? '★' : evt.level === '重要' ? '●' : '○';
                    if (evt.summary) parts.push(`${mark}${evt.summary}`);
                }
            }

            if (meta?.npcs) {
                for (const [name, info] of Object.entries(meta.npcs)) {
                    let s = `NPC:${name}`;
                    if (info.relationship) s += `(${info.relationship})`;
                    parts.push(s);
                }
            }

            if (meta?.items && Object.keys(meta.items).length > 0) {
                for (const [name, info] of Object.entries(meta.items)) {
                    let s = `${info.icon || ''}${name}`;
                    if (info.holder) s += `=${info.holder}`;
                    parts.push(s);
                }
            }

            if (parts.length > 0) {
                lines.push(`#${r.messageIndex} ${parts.join(' | ')}`);
            }
        }

        return lines.length > 1 ? lines.join('\n') : '';
    }

    _extractCleanText(mes, stripTags) {
        if (!mes) return '';
        let text = mes
            .replace(/<think>[\s\S]*?<\/think>/gi, '')
            .replace(/<thinking>[\s\S]*?<\/thinking>/gi, '')
            .replace(/<!--[\s\S]*?-->/g, '');
        if (stripTags) {
            const tags = stripTags.split(/[,，\s]+/).map(t => t.trim()).filter(Boolean);
            for (const tag of tags) {
                const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                text = text.replace(new RegExp(`<${escaped}(?:\\s[^>]*)?>[\\s\\S]*?</${escaped}>`, 'gi'), '');
            }
        }
        return text.replace(/<[^>]*>/g, '').trim();
    }

    /**
     * 構建時間標籤：(相對時間 絕對日期 時間)
     * 例：(前天 霜降月第一日 19:10) 或 (今天 07:55)
     */
    _buildTimeTag(timestamp, currentDate) {
        if (!timestamp) return '';

        const storyDate = timestamp.story_date;
        const storyTime = timestamp.story_time;
        const parts = [];

        if (storyDate && currentDate) {
            const relDesc = this._getRelativeTimeDesc(storyDate, currentDate);
            if (relDesc) {
                parts.push(relDesc.replace(/[()]/g, ''));
            }
        }

        if (storyDate) parts.push(storyDate);
        if (storyTime) parts.push(storyTime);

        if (parts.length === 0) return '';

        const combined = parts.join(' ');
        return `(${combined})`;
    }

    _getRelativeTimeDesc(eventDate, currentDate) {
        if (!eventDate || !currentDate) return '';
        const result = calculateDetailedRelativeTime(eventDate, currentDate);
        if (result.days === null || result.days === undefined) return '';

        const { days, fromDate, toDate } = result;
        if (days === 0) return '(今天)';
        if (days === 1) return '(昨天)';
        if (days === 2) return '(前天)';
        if (days === 3) return '(大前天)';
        if (days >= 4 && days <= 13 && fromDate) {
            const WD = ['日', '一', '二', '三', '四', '五', '六'];
            return `(上週${WD[fromDate.getDay()]})`;
        }
        if (days >= 20 && days < 60 && fromDate && toDate && fromDate.getMonth() !== toDate.getMonth()) {
            return `(上個月${fromDate.getDate()}號)`;
        }
        if (days >= 300 && fromDate && toDate && fromDate.getFullYear() < toDate.getFullYear()) {
            return `(去年${fromDate.getMonth() + 1}月)`;
        }
        if (days > 0 && days < 30) return `(${days}天前)`;
        if (days > 0) return `(${Math.round(days / 30)}個月前)`;
        return '';
    }

    // ========================================
    // Worker 通訊
    // ========================================

    _embed(texts) {
        if (this.isApiMode) return this._embedApi(texts);
        if (!this.worker) return Promise.resolve(null);
        const id = ++this._callId;
        return new Promise((resolve, reject) => {
            this._pendingCallbacks.set(id, { resolve, reject });
            this.worker.postMessage({ type: 'embed', id, data: { texts } });
            setTimeout(() => {
                if (this._pendingCallbacks.has(id)) {
                    this._pendingCallbacks.delete(id);
                    reject(new Error('Embedding 超時'));
                }
            }, 30000);
        });
    }

    async _embedApi(texts) {
        const endpoint = `${this._apiUrl}/embeddings`;
        try {
            const resp = await fetch(endpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${this._apiKey}`,
                },
                body: JSON.stringify({
                    model: this._apiModel,
                    input: texts,
                }),
            });
            if (!resp.ok) {
                const errText = await resp.text().catch(() => '');
                throw new Error(`API ${resp.status}: ${errText.slice(0, 200)}`);
            }
            const json = await resp.json();
            if (!json.data || !Array.isArray(json.data)) {
                throw new Error('API 返回格式異常：缺少 data 數組');
            }
            const vectors = json.data
                .sort((a, b) => a.index - b.index)
                .map(d => d.embedding);
            return { vectors };
        } catch (err) {
            console.error('[Horae Vector] API embedding 失敗:', err);
            throw err;
        }
    }

    /**
     * Rerank API 呼叫（Cohere/Jina/Qwen 相容格式）
     * @returns {Array<{index: number, relevance_score: number}>}
     */
    async _rerank(query, documents, topN, settings) {
        const baseUrl = (settings.vectorRerankUrl || settings.vectorApiUrl || '').replace(/\/+$/, '');
        const apiKey = settings.vectorRerankKey || settings.vectorApiKey || '';
        const model = settings.vectorRerankModel || '';

        if (!baseUrl || !model) throw new Error('Rerank API 地址或模型未配置');

        const endpoint = `${baseUrl}/rerank`;
        console.log(`[Horae Vector] Rerank 請求: ${documents.length} 條候選 → ${endpoint}`);

        const resp = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`,
            },
            body: JSON.stringify({
                model,
                query,
                documents,
                top_n: topN,
            }),
        });

        if (!resp.ok) {
            const errText = await resp.text().catch(() => '');
            throw new Error(`Rerank API ${resp.status}: ${errText.slice(0, 200)}`);
        }

        const json = await resp.json();
        const results = json.results || json.data;
        if (!Array.isArray(results)) {
            throw new Error('Rerank API 返回格式異常：缺少 results 數組');
        }

        return results.map(r => ({
            index: r.index,
            relevance_score: r.relevance_score ?? r.score ?? 0,
        })).sort((a, b) => b.relevance_score - a.relevance_score);
    }

    // ========================================
    // IndexedDB
    // ========================================

    async _openDB() {
        if (this.db) {
            try {
                this.db.transaction(STORE_NAME, 'readonly');
                return;
            } catch (_) {
                console.warn('[Horae Vector] DB connection stale, reconnecting...');
                try { this.db.close(); } catch (__) {}
                this.db = null;
            }
        }
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(DB_NAME, DB_VERSION);
            req.onupgradeneeded = () => {
                const db = req.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    const store = db.createObjectStore(STORE_NAME, { keyPath: 'key' });
                    store.createIndex('chatId', 'chatId', { unique: false });
                }
            };
            req.onblocked = () => {
                console.warn('[Horae Vector] DB upgrade blocked by another tab, closing old connection');
            };
            req.onsuccess = () => {
                this.db = req.result;
                this.db.onversionchange = () => {
                    this.db.close();
                    this.db = null;
                    console.log('[Horae Vector] DB closed due to version change in another tab');
                };
                this.db.onclose = () => { this.db = null; };
                resolve();
            };
            req.onerror = () => reject(req.error);
        });
    }

    async _saveVector(messageIndex, data) {
        await this._openDB();
        const key = `${this.chatId}_${messageIndex}`;
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction(STORE_NAME, 'readwrite');
            tx.objectStore(STORE_NAME).put({
                key,
                chatId: this.chatId,
                messageIndex,
                vector: data.vector,
                hash: data.hash,
                document: data.document,
            });
            tx.oncomplete = resolve;
            tx.onerror = () => reject(tx.error);
        });
    }

    async _loadAllVectors() {
        await this._openDB();
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction(STORE_NAME, 'readonly');
            const index = tx.objectStore(STORE_NAME).index('chatId');
            const req = index.getAll(this.chatId);
            req.onsuccess = () => resolve(req.result || []);
            req.onerror = () => reject(req.error);
        });
    }

    async _deleteVector(messageIndex) {
        await this._openDB();
        const key = `${this.chatId}_${messageIndex}`;
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction(STORE_NAME, 'readwrite');
            tx.objectStore(STORE_NAME).delete(key);
            tx.oncomplete = resolve;
            tx.onerror = () => reject(tx.error);
        });
    }

    async _clearVectors() {
        await this._openDB();
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction(STORE_NAME, 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            const index = store.index('chatId');
            const req = index.openCursor(this.chatId);
            req.onsuccess = () => {
                const cursor = req.result;
                if (cursor) {
                    cursor.delete();
                    cursor.continue();
                }
            };
            tx.oncomplete = resolve;
            tx.onerror = () => reject(tx.error);
        });
    }

    // ========================================
    // 工具函數
    // ========================================

    _hasOriginalEvents(meta) {
        if (!meta?.events?.length) return false;
        return meta.events.some(e => !e.isSummary && e.level !== '摘要' && !e._summaryId);
    }

    _dotProduct(a, b) {
        if (!a || !b || a.length !== b.length) return 0;
        let sum = 0;
        for (let i = 0; i < a.length; i++) sum += a[i] * b[i];
        return sum;
    }

    _hashString(str) {
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            hash = ((hash << 5) - hash) + str.charCodeAt(i);
            hash |= 0;
        }
        return hash.toString(36);
    }

    _extractKeyTerms(document) {
        return document
            .split(/[\s|,，。！？：；、()\[\]（）\n]+/)
            .filter(t => t.length >= 2 && t.length <= 20);
    }

    _updateTermCounts(document, delta) {
        const terms = this._extractKeyTerms(document);
        const unique = new Set(terms);
        for (const term of unique) {
            const prev = this.termCounts.get(term) || 0;
            const next = prev + delta;
            if (next <= 0) this.termCounts.delete(term);
            else this.termCounts.set(term, next);
        }
    }

    _prepareText(text, isQuery) {
        const cfg = MODEL_CONFIG[this.modelName];
        if (cfg?.prefix) {
            return isQuery ? `${cfg.prefix.query}${text}` : `${cfg.prefix.passage}${text}`;
        }
        return text;
    }
}

export const vectorManager = new VectorManager();
