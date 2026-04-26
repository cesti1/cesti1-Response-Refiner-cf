import * as SillyTavern from "../../../../script.js";
import * as Extensions from "../../../extensions.js";

const {
    chat,
    event_types,
    eventSource,
    saveChatConditional,
    saveSettingsDebounced,
} = SillyTavern;

const extension_settings = Extensions.extension_settings || {};
const renderExtensionTemplateAsync = Extensions.renderExtensionTemplateAsync;

function updateMessageBlockCompat(messageId, message) {
    if (typeof SillyTavern.updateMessageBlock === "function") {
        SillyTavern.updateMessageBlock(messageId, message);
        return;
    }

    const text = String(message?.mes ?? "");
    const $messageBlock = $(`#chat .mes[mesid="${messageId}"]`);
    const $text = $messageBlock.find(".mes_text");
    if ($text.length) {
        $text.text(text);
    }
}

const MODULE_NAME = "cesti1-Response-Refiner-cf";
const SETTINGS_KEY = /** @type {const} */ ("response_refiner");
const VERSION = "0.1.0-beta.1";

const DEFAULT_ENDPOINTS = {
    openrouter: "https://openrouter.ai/api/v1",
    gemini: "https://generativelanguage.googleapis.com/v1beta",
    claude: "https://api.anthropic.com",
    openai: "https://api.openai.com/v1",
    deepseek: "https://api.deepseek.com/v1",
};

const PROVIDERS = {
    cloudflare: {
        label: "Cloudflare Worker 反代",
        endpointLabel: "Cloudflare Worker 地址",
        endpointPlaceholder: "例如: 填写你的 Worker 地址",
        defaultEndpoint: "",
        apiStyle: "openai-compatible",
        modelFetch: "openai-compatible",
        requiresEndpoint: true,
    },
    direct: {
        label: "直接连接（OpenAI 兼容）",
        endpointLabel: "API 接口地址",
        endpointPlaceholder: "例如: 填写兼容接口地址",
        defaultEndpoint: "",
        apiStyle: "openai-compatible",
        modelFetch: "openai-compatible",
        requiresEndpoint: true,
    },
    openrouter: {
        label: "OpenRouter",
        endpointLabel: "OpenRouter 接口地址",
        endpointPlaceholder: DEFAULT_ENDPOINTS.openrouter,
        defaultEndpoint: DEFAULT_ENDPOINTS.openrouter,
        apiStyle: "openai-compatible",
        modelFetch: "openai-compatible",
        requiresEndpoint: true,
    },
    gemini: {
        label: "Gemini",
        endpointLabel: "Gemini API 地址",
        endpointPlaceholder: DEFAULT_ENDPOINTS.gemini,
        defaultEndpoint: DEFAULT_ENDPOINTS.gemini,
        apiStyle: "gemini",
        modelFetch: "gemini",
        requiresEndpoint: true,
    },
    claude: {
        label: "Claude",
        endpointLabel: "Claude API 地址",
        endpointPlaceholder: DEFAULT_ENDPOINTS.claude,
        defaultEndpoint: DEFAULT_ENDPOINTS.claude,
        apiStyle: "claude",
        modelFetch: "claude",
        requiresEndpoint: true,
    },
    openai: {
        label: "OpenAI",
        endpointLabel: "OpenAI API 地址",
        endpointPlaceholder: DEFAULT_ENDPOINTS.openai,
        defaultEndpoint: DEFAULT_ENDPOINTS.openai,
        apiStyle: "openai-compatible",
        modelFetch: "openai-compatible",
        requiresEndpoint: true,
    },
    deepseek: {
        label: "DeepSeek",
        endpointLabel: "DeepSeek API 地址",
        endpointPlaceholder: DEFAULT_ENDPOINTS.deepseek,
        defaultEndpoint: DEFAULT_ENDPOINTS.deepseek,
        apiStyle: "openai-compatible",
        modelFetch: "openai-compatible",
        requiresEndpoint: true,
    },
};

const DEFAULT_PROVIDER_SETTINGS = Object.fromEntries(
    Object.entries(PROVIDERS).map(([key, provider]) => [
        key,
        {
            endpoint: provider.defaultEndpoint,
            model: "",
            apiKey: "",
            models: [],
            allowManualModel: true,
        },
    ]),
);

const DEFAULT_SETTINGS = {
    version: VERSION,
    connectionType: "cloudflare",
    providers: DEFAULT_PROVIDER_SETTINGS,
    prompt: "你是一个文本润色助手。你只能处理用户提供的文本，不引入对话历史、设定扩写、旁白说明或解释。保持原意不变，只提升措辞、流畅度、节奏、文笔与可读性。直接输出处理后的正文，不要添加前后缀说明。",
    userPrompt: "你是一个用户输入润色助手。用户输入通常只是一小段简单描述、指令或台词。你只能在原始输入范围内改善措辞、错别字、语序和清晰度，不得扩写成章节、正文、剧情段落、完整回复或新增设定。保持长度与信息量基本接近原文，直接输出润色后的用户输入，不要添加解释、标题、前后缀或引号。",
    temperature: 0.7,
    maxTokens: 1200,
    filterRegex: "<content>([\\s\\S]*?)</content>",
    filterEnabled: true,
    features: {
        refineEnabled: true,
        formatEnabled: true,
        completionEnabled: true,
    },
    forbiddenPhrases: "",
    formatRules: [
        {
            id: "content",
            enabled: true,
            name: "正文",
            startTag: "<content>",
            endTag: "</content>",
            prompt: "正文内容需要完整、连贯、符合上下文语气。",
            template: "<content>\n这里填写正文内容\n</content>",
        },
    ],
    completionChainRegex: "",
    completionOutlineRegex: "",
    completionPrompt: "如果当前 AI 回复被截断，请根据格式规则、可用思维链或需要补完回复的上一条用户消息继续补完。只输出续写或缺失标签内容，不要解释。",
    completionContextMessages: 0,
    streamStatusEnabled: false,
    uiCollapsedSections: {},
    collapsedFormatRules: {},
};

/** @typedef {{ mes?: string, is_user?: boolean, is_system?: boolean, extra?: Record<string, any> }} RefinerChatMessage */

const state = {
    initialized: false,
    busyMessageIds: new Set(),
    requestControllers: new Map(),
    statusBuffers: new Map(),
    comparisonPanelVisible: false,
    extractedRules: [],
};

function deepClone(value) {
    return JSON.parse(JSON.stringify(value));
}

function mergeDefaults(target, defaults) {
    for (const [key, value] of Object.entries(defaults)) {
        if (target[key] === undefined) {
            target[key] = deepClone(value);
        } else if (
            value &&
            typeof value === "object" &&
            !Array.isArray(value) &&
            target[key] &&
            typeof target[key] === "object" &&
            !Array.isArray(target[key])
        ) {
            mergeDefaults(target[key], value);
        }
    }
    return target;
}

/** @returns {typeof DEFAULT_SETTINGS} */
function getSettings() {
    const root = /** @type {Record<string, any>} */ (extension_settings);
    root[SETTINGS_KEY] = root[SETTINGS_KEY] || {};
    const settings = root[SETTINGS_KEY];
    migrateLegacySettings(settings);
    mergeDefaults(settings, DEFAULT_SETTINGS);
    settings.version = VERSION;
    return settings;
}

function migrateLegacySettings(settings) {
    settings.providers = settings.providers || deepClone(DEFAULT_PROVIDER_SETTINGS);

    if (settings.endpoint !== undefined || settings.cloudflareModel !== undefined || settings.cloudflareApiKey !== undefined) {
        settings.providers.cloudflare = settings.providers.cloudflare || deepClone(DEFAULT_PROVIDER_SETTINGS.cloudflare);
        settings.providers.cloudflare.endpoint = settings.providers.cloudflare.endpoint || settings.endpoint || "";
        settings.providers.cloudflare.model = settings.providers.cloudflare.model || settings.cloudflareModel || settings.model || "";
        settings.providers.cloudflare.apiKey = settings.providers.cloudflare.apiKey || settings.cloudflareApiKey || settings.apiKey || "";
    }

    if (settings.directEndpoint !== undefined || settings.directModel !== undefined || settings.directApiKey !== undefined) {
        settings.providers.direct = settings.providers.direct || deepClone(DEFAULT_PROVIDER_SETTINGS.direct);
        settings.providers.direct.endpoint = settings.providers.direct.endpoint || settings.directEndpoint || "";
        settings.providers.direct.model = settings.providers.direct.model || settings.directModel || "";
        settings.providers.direct.apiKey = settings.providers.direct.apiKey || settings.directApiKey || "";
    }

    if (settings.openrouterEndpoint !== undefined || settings.openrouterModel !== undefined || settings.openrouterApiKey !== undefined) {
        settings.providers.openrouter = settings.providers.openrouter || deepClone(DEFAULT_PROVIDER_SETTINGS.openrouter);
        settings.providers.openrouter.endpoint = settings.providers.openrouter.endpoint || settings.openrouterEndpoint || PROVIDERS.openrouter.defaultEndpoint;
        settings.providers.openrouter.model = settings.providers.openrouter.model || settings.openrouterModel || "";
        settings.providers.openrouter.apiKey = settings.providers.openrouter.apiKey || settings.openrouterApiKey || "";
    }

    if (settings.completionChainRegex === undefined && settings.completionOutlineRegex !== undefined) {
        settings.completionChainRegex = settings.completionOutlineRegex || "";
    }
}

function saveSettings() {
    saveSettingsDebounced();
}

function getProviderKey() {
    const settings = getSettings();
    return PROVIDERS[settings.connectionType] ? settings.connectionType : "cloudflare";
}

function getProviderSettings(providerKey = getProviderKey()) {
    const settings = getSettings();
    settings.providers[providerKey] = settings.providers[providerKey] || deepClone(DEFAULT_PROVIDER_SETTINGS[providerKey]);
    return settings.providers[providerKey];
}

function normalizeEndpoint(endpoint) {
    return String(endpoint || "").trim().replace(/\/+$/, "");
}

function getCurrentModel() {
    return String(getProviderSettings().model || "").trim();
}

function getCurrentApiKey() {
    return String(getProviderSettings().apiKey || "").trim();
}

function getApiEndpoint() {
    const providerKey = getProviderKey();
    const provider = PROVIDERS[providerKey];
    const providerSettings = getProviderSettings(providerKey);
    return normalizeEndpoint(providerSettings.endpoint || provider.defaultEndpoint || "");
}

function setText($element, text) {
    $element.text(String(text ?? ""));
}

function splitForbiddenPhrases(settings) {
    return String(settings.forbiddenPhrases || "")
        .split(/[，,]/)
        .map(item => item.trim())
        .filter(Boolean);
}

function buildForbiddenInstruction(settings) {
    const phrases = splitForbiddenPhrases(settings);
    if (!phrases.length) {
        return "";
    }

    return [
        "额外要求：润色过程中必须处理下列不允许出现的八股文/模板化字符或短语。",
        "如果原文中存在这些内容，需要在保持语义自然的前提下替换、改写或移除。",
        "最终输出中不得出现这些内容：",
        phrases.map(item => `- ${item}`).join("\n"),
    ].join("\n");
}

function stripCodeFence(text) {
    const value = String(text || "").trim();
    const match = value.match(/^```(?:\w+)?\s*([\s\S]*?)\s*```$/);
    return match ? match[1].trim() : value;
}

function isAssistantMessage(message) {
    return Boolean(message && !message.is_user && !message.is_system);
}

function isUserMessage(message) {
    return Boolean(message && message.is_user && !message.is_system);
}

function isRefinableMessage(message) {
    return isAssistantMessage(message) || isUserMessage(message);
}

function getMessageText(message) {
    return message?.mes || "";
}

function getMessageById(messageId) {
    const message = /** @type {RefinerChatMessage | undefined} */ (chat[messageId]);
    if (!isRefinableMessage(message)) {
        return { message: null, messageId: -1 };
    }
    return { message, messageId };
}

function getLatestUserMessage(beforeMessageId = chat.length) {
    for (let i = Math.min(beforeMessageId - 1, chat.length - 1); i >= 0; i--) {
        const message = /** @type {RefinerChatMessage | undefined} */ (chat[i]);
        if (isUserMessage(message)) {
            return { message, messageId: i };
        }
    }
    return { message: null, messageId: -1 };
}

function getRecentContext(messageId, limit) {
    const start = Math.max(0, messageId - Math.max(0, Number(limit) || 0));
    return chat
        .slice(start, messageId)
        .map((message, offset) => {
            const item = /** @type {RefinerChatMessage} */ (message);
            if (!item || item.is_system) {
                return null;
            }
            const role = item.is_user ? "用户" : "AI";
            return `${role}[${start + offset}]:\n${getMessageText(item)}`;
        })
        .filter(Boolean)
        .join("\n\n---\n\n");
}

function extractTextToRefine(text, settings) {
    if (!settings.filterEnabled || !settings.filterRegex) {
        return text;
    }

    try {
        const regex = new RegExp(settings.filterRegex, "g");
        const matches = [];
        let match;
        while ((match = regex.exec(text)) !== null) {
            matches.push(match[1] !== undefined ? match[1] : match[0]);
            if (match[0] === "") {
                regex.lastIndex++;
            }
        }
        return matches.length ? matches[matches.length - 1] : text;
    } catch (error) {
        console.error("[Response Refiner] 正则表达式错误:", error);
        toastr.error("正则表达式错误，已改为处理全文", "Response Refiner");
        return text;
    }
}

function replaceRefinedText(originalText, refinedText, settings) {
    if (!settings.filterEnabled || !settings.filterRegex) {
        return refinedText;
    }

    try {
        const regex = new RegExp(settings.filterRegex, "g");
        let lastMatch = null;
        let match;
        while ((match = regex.exec(originalText)) !== null) {
            lastMatch = match;
            if (match[0] === "") {
                regex.lastIndex++;
            }
        }

        if (!lastMatch) {
            return refinedText;
        }

        const lastIndex = lastMatch.index;
        const lastMatchText = lastMatch[0];
        const captured = lastMatch[1];
        const before = originalText.substring(0, lastIndex);
        const after = originalText.substring(lastIndex + lastMatchText.length);

        if (captured !== undefined) {
            return before + lastMatchText.replace(captured, refinedText) + after;
        }

        return before + refinedText + after;
    } catch (error) {
        console.error("[Response Refiner] 文本替换错误:", error);
        return refinedText;
    }
}

function extractChainContext(text, regexText) {
    const pattern = String(regexText || "").trim();
    if (!pattern) {
        return "";
    }

    try {
        const regex = new RegExp(pattern, "g");
        const matches = [];
        let match;
        while ((match = regex.exec(text)) !== null) {
            matches.push(match[1] !== undefined ? match[1] : match[0]);
            if (match[0] === "") {
                regex.lastIndex++;
            }
        }
        return matches.length ? String(matches[matches.length - 1]).trim() : "";
    } catch (error) {
        console.error("[Response Refiner] 思维链捕获正则错误:", error);
        toastr.warning("思维链捕获正则错误，已仅使用上一条用户消息", "回复补完");
        return "";
    }
}

function inferBodyTagsFromFilterRegex(regexText) {
    const text = String(regexText || "").trim();
    const match = text.match(/^<([a-zA-Z][\w:-]*)\b[^>]*>\s*\(\[\\s\\S\]\*\?\)\s*<\/\1>$/)
        || text.match(/^<([a-zA-Z][\w:-]*)\b[^>]*>\(\[\\s\\S\]\*\?\)<\/\1>$/)
        || text.match(/^<([a-zA-Z][\w:-]*)\b[^>]*>\(\.\*\?\)<\/\1>$/);
    if (!match) return null;
    return { tag: match[1], startTag: `<${match[1]}>`, endTag: `</${match[1]}>` };
}

function getBodyRuleInfo(settings, notify = false) {
    const inferred = inferBodyTagsFromFilterRegex(settings.filterRegex);
    if (!inferred) {
        if (notify && settings.filterEnabled) {
            toastr.warning("润色捕获正则过复杂，无法自动识别正文标签；请确保正文规则标签与润色捕获正则一致。", "正文规则");
        }
        return { rule: null, key: "", startTag: "", endTag: "", inferred: null };
    }

    let rule = (settings.formatRules || []).find(item => item && item.startTag === inferred.startTag && item.endTag === inferred.endTag);
    if (!rule) {
        rule = (settings.formatRules || []).find(item => item && (item.id === "content" || item.name === "正文"));
    }
    if (!rule) {
        rule = {
            id: "content",
            enabled: true,
            name: "正文",
            startTag: inferred.startTag,
            endTag: inferred.endTag,
            prompt: "正文内容需要完整、连贯、符合上下文语气。",
            template: `${inferred.startTag}
这里填写正文内容
${inferred.endTag}`,
        };
        settings.formatRules = settings.formatRules || [];
        settings.formatRules.unshift(rule);
    }
    rule.id = rule.id || "content";
    rule.enabled = true;
    rule.name = "正文";
    rule.startTag = inferred.startTag;
    rule.endTag = inferred.endTag;
    if (!rule.prompt) rule.prompt = "正文内容需要完整、连贯、符合上下文语气。";
    if (!rule.template) rule.template = `${inferred.startTag}
这里填写正文内容
${inferred.endTag}`;
    return { rule, key: getRuleKey(rule), startTag: inferred.startTag, endTag: inferred.endTag, inferred };
}

function ensureBodyRule(settings, notify = false) {
    return getBodyRuleInfo(settings, notify);
}

function getEnabledFormatRules(settings) {
    return (settings.formatRules || [])
        .filter(rule => rule && rule.enabled && rule.startTag && rule.endTag)
        .map((rule, index) => ({ ...rule, order: index + 1 }));
}

function buildFormatRulesText(settings) {
    const rules = getEnabledFormatRules(settings);
    if (!rules.length) {
        return "未配置启用的格式标签规则。";
    }

    return rules.map(rule => [
        `${rule.order}. ${rule.name || "未命名标签"}`,
        `规则ID：${getRuleKey(rule)}`,
        `开始标签：${rule.startTag}`,
        `结束标签：${rule.endTag}`,
        `标签内提示词：${rule.prompt || "无"}`,
        `标签内容模板：\n${rule.template || `${rule.startTag}\n\n${rule.endTag}`}`,
    ].join("\n")).join("\n\n");
}

function getRuleKey(rule) {
    return String(rule.id || rule.name || rule.startTag || "rule").replace(/[^a-zA-Z0-9_\-\u4e00-\u9fa5]/g, "_");
}

function escapeRegExp(value) {
    return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractTaggedSegments(text, rules) {
    const source = String(text || "");
    return rules.map(rule => {
        const start = String(rule.startTag || "");
        const end = String(rule.endTag || "");
        if (!start || !end) {
            return { rule, key: getRuleKey(rule), content: "", found: false };
        }
        const regex = new RegExp(`${escapeRegExp(start)}([\\s\\S]*?)${escapeRegExp(end)}`, "g");
        const matches = [];
        let match;
        while ((match = regex.exec(source)) !== null) {
            matches.push(match[1] || "");
            if (match[0] === "") regex.lastIndex++;
        }
        return { rule, key: getRuleKey(rule), content: matches.length ? matches[matches.length - 1] : "", found: matches.length > 0 };
    });
}

function getNonBodyFormatRules(settings) {
    const body = getBodyRuleInfo(settings, false);
    return getEnabledFormatRules(settings).filter(rule => getRuleKey(rule) !== body.key);
}

function buildFormatRulesTextForRules(rules) {
    if (!rules.length) return "未配置启用的非正文格式标签规则。";
    return rules.map((rule, index) => [
        `${index + 1}. ${rule.name || "未命名标签"}`,
        `规则ID：${getRuleKey(rule)}`,
        `开始标签：${rule.startTag}`,
        `结束标签：${rule.endTag}`,
        `标签内提示词：${rule.prompt || "无"}`,
        `标签内容模板：\n${rule.template || `${rule.startTag}\n\n${rule.endTag}`}`,
    ].join("\n")).join("\n\n");
}

function buildFormatReplacementPromptText(sourceText, settings) {
    const rules = getNonBodyFormatRules(settings);
    const segments = extractTaggedSegments(sourceText, rules);
    return [
        "你是 AI 回复格式检查和补全修正助手。",
        "你只处理非正文标签；正文标签由润色功能处理，禁止输出或修改正文规则对应的内容。",
        "你的任务是修正每个非正文规则对应标签内部的文本，使其满足标签提示词和模板要求。",
        "为了减少输出 token，你绝对不要输出完整回复、开始标签、结束标签、解释或 Markdown 代码块。",
        "你必须只输出一个 JSON 对象，键为规则ID，值为该规则标签内修正后的纯文本。",
        "如果某个非正文标签不存在但规则要求补全，请仍在对应规则ID里输出应插入的标签内文本。",
        "格式标签规则：",
        buildFormatRulesTextForRules(rules),
        "当前标签内内容：",
        JSON.stringify(Object.fromEntries(segments.map(item => [item.key, item.content])), null, 2),
        "完整 AI 回复文本仍作为上下文提供，但不要完整复述：",
        sourceText,
    ].join("\n");
}

function buildRefineMessages(sourceText, settings, isUser) {
    const basePrompt = isUser ? settings.userPrompt : settings.prompt;
    const forbiddenInstruction = buildForbiddenInstruction(settings);
    const system = [basePrompt, forbiddenInstruction]
        .filter(Boolean)
        .join("\n\n");

    return [
        { role: "system", content: system },
        { role: "user", content: sourceText },
    ];
}

function buildFormatMessages(sourceText, settings, replacementOnly = false) {
    if (replacementOnly) {
        return [
            { role: "system", content: "你是严格的格式标签内容修正器。你只输出符合要求的 JSON 对象，不输出解释、完整回复、标签或 Markdown。" },
            { role: "user", content: buildFormatReplacementPromptText(sourceText, settings) },
        ];
    }

    const system = [
        "你是 AI 回复格式检查和补全修正助手。",
        "你只能根据用户提供的当前 AI 回复文本、格式标签规则、标签提示词和模板进行修正。",
        "需要补全缺失标签、修复错误顺序、闭合不完整标签，并按模板与提示词修正标签内内容。",
        "不得引入无关设定，不得解释修改过程。最终只输出完整修正后的 AI 回复。",
        "格式标签规则：",
        buildFormatRulesText(settings),
    ].join("\n");

    return [
        { role: "system", content: system },
        { role: "user", content: sourceText },
    ];
}

function buildCompletionMessages(sourceText, settings, messageId, missingRules = []) {
    const chain = extractChainContext(sourceText, settings.completionChainRegex || settings.completionOutlineRegex);
    const previousUser = getLatestUserMessage(messageId);
    const previousUserText = previousUser.message ? getMessageText(previousUser.message) : "";
    const missingText = missingRules.length ? buildFormatRulesTextForRules(missingRules) : "无明确缺失标签，仅按截断位置补完。";

    const system = [
        "你是 AI 回复补完助手。当前 AI 回复可能被截断或缺少部分标签。",
        "你必须配合格式检查规则继续生成剩余部分，续写内容需要满足标签顺序、开始标签、结束标签、标签内提示词和模板要求。",
        "补完依据：需要补完回复的上一条用户消息始终是主上下文；如果提供思维链，则结合思维链和上一条用户消息；没有思维链时只依靠上一条用户消息。",
        "如果正文标签已开始但未完成，允许从正文开头或截断处继续补完，并遵守正文规则；格式修正阶段之后不会处理正文标签。",
        "只输出续写或缺失标签内容，不要重复已经存在的完整内容，不要输出说明、标题或分析。",
        settings.completionPrompt || "",
        "全部格式标签规则：",
        buildFormatRulesText(settings),
    ].filter(Boolean).join("\n");

    const user = [
        `上一条用户消息：\n${previousUserText || "未找到上一条用户消息"}`,
        chain ? `参考思维链：\n${chain}` : "参考思维链：未匹配到，禁止自行编造思维链。",
        `当前缺失或需要关注的标签规则：\n${missingText}`,
        `已经生成但可能被截断的 AI 回复：\n${sourceText}`,
        "请只输出需要追加到当前回复末尾的内容。",
    ].join("\n\n---\n\n");

    return [
        { role: "system", content: system },
        { role: "user", content: user },
    ];
}

function getOpenAICompatibleHeaders(providerKey, apiKey) {
    const headers = {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`,
    };

    if (providerKey === "openrouter") {
        headers["HTTP-Referer"] = window.location.origin;
        headers["X-Title"] = "SillyTavern Response Refiner";
    }

    return headers;
}

async function callAI(messages, options = {}) {
    const settings = getSettings();
    const providerKey = getProviderKey();
    const provider = PROVIDERS[providerKey];
    const endpoint = getApiEndpoint();
    const model = getCurrentModel();
    const apiKey = getCurrentApiKey();
    const maxTokens = Number(options.maxTokens || settings.maxTokens) || DEFAULT_SETTINGS.maxTokens;
    const temperature = Number(options.temperature ?? settings.temperature) || 0;
    const signal = options.signal;
    const onToken = typeof options.onToken === "function" ? options.onToken : null;
    const stream = Boolean(options.stream && onToken);

    if (!endpoint || !model) {
        throw new Error("请先配置接口地址和模型");
    }
    if (!apiKey) {
        throw new Error("请先配置 API Key");
    }

    if (provider.apiStyle === "gemini") {
        if (onToken) onToken("[Gemini 暂使用非流式请求，正在等待完整响应...]\n");
        return callGemini(endpoint, model, apiKey, messages, temperature, maxTokens, signal);
    }

    if (provider.apiStyle === "claude") {
        if (onToken) onToken("[Claude 暂使用非流式请求，正在等待完整响应...]\n");
        return callClaude(endpoint, model, apiKey, messages, temperature, maxTokens, signal);
    }

    return callOpenAICompatible(providerKey, endpoint, model, apiKey, messages, temperature, maxTokens, signal, stream, onToken);
}

async function callOpenAICompatible(providerKey, endpoint, model, apiKey, messages, temperature, maxTokens, signal, stream = false, onToken = null) {
    const response = await fetch(`${endpoint}/chat/completions`, {
        method: "POST",
        headers: getOpenAICompatibleHeaders(providerKey, apiKey),
        signal,
        body: JSON.stringify({
            model,
            messages,
            temperature,
            max_tokens: maxTokens,
            stream,
        }),
    });

    if (stream && response.ok && response.body) {
        return readOpenAIStream(response, onToken);
    }

    const data = await safeJson(response);
    if (!response.ok) {
        throw new Error(data?.error?.message || data?.message || `HTTP ${response.status}`);
    }

    const content = data?.choices?.[0]?.message?.content;
    if (typeof content !== "string") {
        throw new Error("响应格式不正确：未找到 choices[0].message.content");
    }

    return stripCodeFence(content);
}

async function readOpenAIStream(response, onToken) {
    const reader = response.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";
    let result = "";
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split(/\r?\n/);
        buffer = lines.pop() || "";
        for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed || !trimmed.startsWith("data:")) continue;
            const payload = trimmed.slice(5).trim();
            if (payload === "[DONE]") continue;
            try {
                const data = JSON.parse(payload);
                const token = data?.choices?.[0]?.delta?.content || data?.choices?.[0]?.text || "";
                if (token) {
                    result += token;
                    onToken?.(token);
                }
            } catch (_error) {
                // 忽略单行非 JSON 流片段。
            }
        }
    }
    return stripCodeFence(result);
}

async function callGemini(endpoint, model, apiKey, messages, temperature, maxTokens, signal) {
    const modelName = model.startsWith("models/") ? model : `models/${model}`;
    const systemText = messages.filter(item => item.role === "system").map(item => item.content).join("\n\n");
    const userText = messages.filter(item => item.role !== "system").map(item => item.content).join("\n\n");
    const response = await fetch(`${endpoint}/${modelName}:generateContent?key=${encodeURIComponent(apiKey)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal,
        body: JSON.stringify({
            systemInstruction: systemText ? { parts: [{ text: systemText }] } : undefined,
            contents: [{ role: "user", parts: [{ text: userText }] }],
            generationConfig: {
                temperature,
                maxOutputTokens: maxTokens,
            },
        }),
    });

    const data = await safeJson(response);
    if (!response.ok) {
        throw new Error(data?.error?.message || data?.message || `HTTP ${response.status}`);
    }

    const content = data?.candidates?.[0]?.content?.parts?.map(part => part.text || "").join("");
    if (typeof content !== "string") {
        throw new Error("响应格式不正确：未找到 Gemini candidates 内容");
    }

    return stripCodeFence(content);
}

async function callClaude(endpoint, model, apiKey, messages, temperature, maxTokens, signal) {
    const system = messages.filter(item => item.role === "system").map(item => item.content).join("\n\n");
    const user = messages.filter(item => item.role !== "system").map(item => item.content).join("\n\n");
    const response = await fetch(`${endpoint}/v1/messages`, {
        method: "POST",
        signal,
        headers: {
            "Content-Type": "application/json",
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
            model,
            max_tokens: maxTokens,
            temperature,
            system,
            messages: [{ role: "user", content: user }],
        }),
    });

    const data = await safeJson(response);
    if (!response.ok) {
        throw new Error(data?.error?.message || data?.message || `HTTP ${response.status}`);
    }

    const content = data?.content?.map(part => part.text || "").join("");
    if (typeof content !== "string") {
        throw new Error("响应格式不正确：未找到 Claude content 内容");
    }

    return stripCodeFence(content);
}

async function safeJson(response) {
    try {
        return await response.json();
    } catch (_error) {
        return {};
    }
}

function normalizeModelList(data, providerKey) {
    if (providerKey === "gemini") {
        const models = Array.isArray(data?.models) ? data.models : [];
        return models
            .filter(model => String(model.name || "").includes("models/"))
            .map(model => ({
                id: String(model.name).replace(/^models\//, ""),
                name: model.displayName || String(model.name).replace(/^models\//, ""),
            }));
    }

    const models = Array.isArray(data?.data) ? data.data : [];
    return models.map(model => ({
        id: model.id || model.name,
        name: model.name || model.display_name || model.id,
    })).filter(model => model.id);
}

async function fetchProviderModels(providerKey = getProviderKey()) {
    const provider = PROVIDERS[providerKey];
    const providerSettings = getProviderSettings(providerKey);
    const endpoint = normalizeEndpoint(providerSettings.endpoint || provider.defaultEndpoint || "");
    const apiKey = String(providerSettings.apiKey || "").trim();

    if (!endpoint) {
        toastr.warning("请先填写接口地址", "模型列表");
        return [];
    }
    if (!apiKey) {
        toastr.warning("请先保存 API Key", "模型列表");
        return [];
    }

    let response;
    if (provider.modelFetch === "gemini") {
        response = await fetch(`${endpoint}/models?key=${encodeURIComponent(apiKey)}`);
    } else if (provider.modelFetch === "claude") {
        response = await fetch(`${endpoint}/v1/models`, {
            headers: {
                "x-api-key": apiKey,
                "anthropic-version": "2023-06-01",
            },
        });
    } else {
        response = await fetch(`${endpoint}/models`, {
            headers: getOpenAICompatibleHeaders(providerKey, apiKey),
        });
    }

    const data = await safeJson(response);
    if (!response.ok) {
        throw new Error(data?.error?.message || data?.message || `HTTP ${response.status}`);
    }

    return normalizeModelList(data, providerKey).sort((a, b) => String(a.name).localeCompare(String(b.name)));
}

function groupModels(models) {
    const grouped = {};
    for (const model of models) {
        const vendor = String(model.id || "other").includes("/") ? String(model.id).split("/")[0] : "models";
        grouped[vendor] = grouped[vendor] || [];
        grouped[vendor].push(model);
    }
    return grouped;
}

function updateModelSelect(providerKey = getProviderKey()) {
    const settings = getSettings();
    const providerSettings = getProviderSettings(providerKey);
    const models = Array.isArray(providerSettings.models) ? providerSettings.models : [];
    const $select = $("#response_refiner_model_select");
    const $manual = $("#response_refiner_model_manual");

    $select.empty();
    if (!models.length) {
        $select.append($("<option>", { value: "", text: "未加载模型列表，可手动输入" }));
    } else {
        const grouped = groupModels(models);
        for (const [group, items] of Object.entries(grouped)) {
            const $group = $("<optgroup>", { label: group.toUpperCase() });
            for (const model of items) {
                $group.append($("<option>", { value: model.id, text: model.name || model.id }));
            }
            $select.append($group);
        }
    }

    if (providerSettings.model && models.some(model => model.id === providerSettings.model)) {
        $select.val(providerSettings.model);
    }
    $manual.val(providerSettings.model || "");
    settings.connectionType = providerKey;
}

async function refreshProviderModels() {
    const providerKey = getProviderKey();
    const providerSettings = getProviderSettings(providerKey);
    const $button = $("#response_refiner_refresh_models");
    $button.prop("disabled", true).find("i").addClass("fa-spin");

    try {
        const models = await fetchProviderModels(providerKey);
        providerSettings.models = models;
        if (models.length && !models.some(model => model.id === providerSettings.model)) {
            providerSettings.model = models[0].id;
        }
        saveSettings();
        updateModelSelect(providerKey);
        toastr.success(`已获取 ${models.length} 个模型`, "模型列表");
    } catch (error) {
        console.error("[Response Refiner] 获取模型列表失败:", error);
        toastr.error(`获取模型列表失败: ${error instanceof Error ? error.message : String(error)}`, "模型列表");
        updateModelSelect(providerKey);
    } finally {
        $button.prop("disabled", false).find("i").removeClass("fa-spin");
    }
}

function updateConnectionTypeUI() {
    const providerKey = getProviderKey();
    const provider = PROVIDERS[providerKey];
    const providerSettings = getProviderSettings(providerKey);

    setText($("#response_refiner_endpoint_label"), provider.endpointLabel);
    $("#response_refiner_endpoint").attr("placeholder", provider.endpointPlaceholder).val(providerSettings.endpoint || provider.defaultEndpoint || "");
    $("#response_refiner_api_key_input").val(providerSettings.apiKey || "");
    updateModelSelect(providerKey);
}

async function testConnection() {
    try {
        toastr.info("正在测试连接...", "测试连接");
        const text = await callAI([
            { role: "system", content: "你是连接测试助手。" },
            { role: "user", content: "请只回复：连接成功" },
        ], { temperature: 0, maxTokens: 32 });
        if (!text) {
            throw new Error("API 未返回文本");
        }
        toastr.success("连接测试成功", "测试连接");
    } catch (error) {
        console.error("[Response Refiner] 连接测试失败:", error);
        toastr.error(`连接测试失败: ${error instanceof Error ? error.message : String(error)}`, "测试连接");
    }
}

async function runRefine(originalText, settings, isUser, signal, onToken = null) {
    const textToRefine = isUser ? originalText : extractTextToRefine(originalText, settings);
    const refinedText = await callAI(buildRefineMessages(textToRefine, settings, isUser), { signal, stream: settings.streamStatusEnabled, onToken });
    const finalText = isUser ? refinedText : replaceRefinedText(originalText, refinedText, settings);
    return {
        stage: "refine",
        original_text: textToRefine,
        refined_text: refinedText,
        candidate_text: finalText,
    };
}

function parseFormatReplacementOutput(text) {
    const cleaned = stripCodeFence(text);
    try {
        return JSON.parse(cleaned);
    } catch (_error) {
        const match = cleaned.match(/\{[\s\S]*\}/);
        if (match) {
            return JSON.parse(match[0]);
        }
        throw new Error("格式检查返回内容不是可解析 JSON");
    }
}

function replaceTaggedSegmentContent(sourceText, rule, replacement) {
    const start = String(rule.startTag || "");
    const end = String(rule.endTag || "");
    const value = String(replacement ?? "").trim();
    if (!start || !end || !value) {
        return sourceText;
    }

    const regex = new RegExp(`${escapeRegExp(start)}([\\s\\S]*?)${escapeRegExp(end)}`, "g");
    let lastMatch = null;
    let match;
    while ((match = regex.exec(sourceText)) !== null) {
        lastMatch = match;
        if (match[0] === "") regex.lastIndex++;
    }

    if (!lastMatch) {
        return `${sourceText}\n${start}\n${value}\n${end}`;
    }

    const before = sourceText.slice(0, lastMatch.index);
    const after = sourceText.slice(lastMatch.index + lastMatch[0].length);
    return `${before}${start}${value.startsWith("\n") ? "" : "\n"}${value}${value.endsWith("\n") ? "" : "\n"}${end}${after}`;
}

function applyFormatReplacements(sourceText, settings, replacements, rules = getNonBodyFormatRules(settings)) {
    let result = sourceText;
    for (const rule of rules) {
        const key = getRuleKey(rule);
        if (Object.prototype.hasOwnProperty.call(replacements, key)) {
            result = replaceTaggedSegmentContent(result, rule, replacements[key]);
        }
    }
    return result;
}

function getMissingFormatRules(sourceText, settings) {
    return getEnabledFormatRules(settings).filter(rule => {
        const segment = extractTaggedSegments(sourceText, [rule])[0];
        return !segment?.found;
    });
}

async function runFormat(sourceText, settings, signal, replacementOnly = true, onToken = null) {
    if (!replacementOnly) {
        const formattedText = await callAI(buildFormatMessages(sourceText, settings), { signal, stream: settings.streamStatusEnabled, onToken });
        return {
            stage: "format",
            original_text: sourceText,
            refined_text: formattedText,
            candidate_text: formattedText,
        };
    }

    const rules = getNonBodyFormatRules(settings);
    if (!rules.length) {
        return {
            stage: "format",
            original_text: sourceText,
            refined_text: "{}",
            candidate_text: sourceText,
        };
    }
    const replacementText = await callAI(buildFormatMessages(sourceText, settings, true), { signal, stream: settings.streamStatusEnabled, onToken });
    const replacements = parseFormatReplacementOutput(replacementText);
    const formattedText = applyFormatReplacements(sourceText, settings, replacements, rules);
    return {
        stage: "format",
        original_text: sourceText,
        refined_text: JSON.stringify(replacements, null, 2),
        candidate_text: formattedText,
    };
}

async function runCompletion(sourceText, settings, messageId, signal, onToken = null, missingRules = []) {
    const continuation = await callAI(buildCompletionMessages(sourceText, settings, messageId, missingRules), { signal, stream: settings.streamStatusEnabled, onToken });
    const joined = sourceText + continuation;
    const formatted = joined;
    return {
        stage: "completion",
        original_text: sourceText,
        refined_text: continuation,
        candidate_text: formatted,
    };
}

function getSelectedStages(settings, isUser) {
    if (isUser) {
        return settings.features.refineEnabled ? ["refine"] : [];
    }

    const stages = [];
    if (settings.features.refineEnabled) stages.push("refine");
    if (settings.features.formatEnabled) stages.push("format");
    if (settings.features.completionEnabled) stages.push("completion");
    return stages;
}

function appendStatus(messageId, stage, text, reset = false) {
    if (!getSettings().streamStatusEnabled) return;
    const key = `${messageId}:${stage}`;
    if (reset) state.statusBuffers.set(key, "");
    const next = (state.statusBuffers.get(key) || "") + String(text || "");
    state.statusBuffers.set(key, next);
    const $panel = $(`#chat .mes[mesid="${messageId}"] .response-refiner-status-panel`);
    if (!$panel.length) return;
    $panel.find(".response-refiner-status-stage").text(stage);
    $panel.find(".response-refiner-status-text").text(next || "等待模型返回...");
    const node = $panel.find(".response-refiner-status-text").get(0);
    if (node) node.scrollTop = node.scrollHeight;
}

function runStageWithStatus(messageId, stage, fn) {
    appendStatus(messageId, stage, `开始执行：${stage}\n`, true);
    return fn(token => appendStatus(messageId, stage, token));
}

async function runSelectedPipeline(fullOriginalText, settings, isUser, messageId, signal) {
    let workingText = fullOriginalText;
    const stageResults = [];
    if (isUser) {
        if (settings.features.refineEnabled) {
            const result = await runStageWithStatus(messageId, "功能1 润色", onToken => runRefine(workingText, settings, true, signal, onToken));
            stageResults.push(result);
            workingText = result.candidate_text;
        }
        return { workingText, stageResults, stages: stageResults.map(item => item.stage) };
    }

    toastr.info("执行已勾选功能会按需分步处理，最多消耗三次请求。", "Response Refiner");
    const missingRules = getMissingFormatRules(workingText, settings);
    if (settings.features.completionEnabled && missingRules.length) {
        const result = await runStageWithStatus(messageId, "功能3 回复补完", onToken => runCompletion(workingText, settings, messageId, signal, onToken, missingRules));
        stageResults.push(result);
        workingText = result.candidate_text;
    }
    if (settings.features.formatEnabled) {
        const result = await runStageWithStatus(messageId, "功能2 格式检查和补全修正", onToken => runFormat(workingText, settings, signal, true, onToken));
        stageResults.push(result);
        workingText = result.candidate_text;
    }
    if (settings.features.refineEnabled) {
        const result = await runStageWithStatus(messageId, "功能1 润色", onToken => runRefine(workingText, settings, false, signal, onToken));
        stageResults.push(result);
        workingText = result.candidate_text;
    }
    return { workingText, stageResults, stages: ["selected"] };
}

async function requestFeature(messageId, feature) {
    const { message, messageId: resolvedId } = getMessageById(messageId);
    if (!message || resolvedId < 0) {
        toastr.error("未找到有效消息", "Response Refiner");
        return;
    }

    const isUser = isUserMessage(message);
    if (isUser && feature !== "refine" && feature !== "selected") {
        toastr.warning("该功能仅对 AI 回复生效", "Response Refiner");
        return;
    }

    if (state.busyMessageIds.has(resolvedId)) {
        toastr.warning("该消息正在处理中，请稍候", "Response Refiner");
        return;
    }

    const settings = getSettings();
    ensureBodyRule(settings, true);
    let stages = feature === "selected" ? getSelectedStages(settings, isUser) : [feature];
    stages = stages.filter(stage => stage !== "completion" || !isUser);
    stages = stages.filter(stage => stage !== "format" || !isUser);

    if (!stages.length) {
        toastr.warning("没有已启用的可执行功能", "Response Refiner");
        return;
    }

    const controller = new AbortController();
    state.busyMessageIds.add(resolvedId);
    state.requestControllers.set(resolvedId, controller);
    updateMessageButtons(resolvedId);

    try {
        const fullOriginalText = getMessageText(message);
        let workingText = fullOriginalText;
        let stageResults = [];

        if (feature === "selected") {
            const selected = await runSelectedPipeline(fullOriginalText, settings, isUser, resolvedId, controller.signal);
            workingText = selected.workingText;
            stageResults = selected.stageResults;
            stages = selected.stages;
        } else {
            for (const stage of stages) {
                let result;
                if (stage === "refine") {
                    result = await runStageWithStatus(resolvedId, "功能1 润色", onToken => runRefine(workingText, settings, isUser, controller.signal, onToken));
                } else if (stage === "format") {
                    result = await runStageWithStatus(resolvedId, "功能2 格式检查和补全修正", onToken => runFormat(workingText, settings, controller.signal, true, onToken));
                } else if (stage === "completion") {
                    result = await runStageWithStatus(resolvedId, "功能3 回复补完", onToken => runCompletion(workingText, settings, resolvedId, controller.signal, onToken, getMissingFormatRules(workingText, settings)));
                } else {
                    continue;
                }
                stageResults.push(result);
                workingText = result.candidate_text;
            }
        }

        const last = stageResults[stageResults.length - 1];
        message.extra = message.extra || {};
        message.extra.response_refiner = {
            feature,
            stages,
            stage_results: stageResults,
            original_text: stageResults[0]?.original_text || fullOriginalText,
            refined_text: last?.refined_text || workingText,
            full_original_text: fullOriginalText,
            candidate_text: workingText,
            applied: false,
        };

        await saveChatConditional();
        updateMessageButtons(resolvedId);
        updateComparisonPanel(resolvedId);
        toastr.success("处理完成，请查看预览并决定是否替换", "Response Refiner");
    } catch (error) {
        if (error?.name === "AbortError") {
            toastr.warning("请求已停止", "Response Refiner");
        } else {
            console.error("[Response Refiner] 处理失败:", error);
            toastr.error(String(error instanceof Error ? error.message : error), "Response Refiner");
        }
    } finally {
        state.busyMessageIds.delete(resolvedId);
        state.requestControllers.delete(resolvedId);
        updateMessageButtons(resolvedId);
    }
}

function stopRequest(messageId) {
    const controller = state.requestControllers.get(messageId);
    if (controller) {
        controller.abort();
    }
}

async function applyCandidate(messageId) {
    const { message, messageId: resolvedId } = getMessageById(messageId);
    if (!message || resolvedId < 0) return;

    const refinerData = message.extra?.response_refiner;
    if (!refinerData?.candidate_text) {
        toastr.warning("没有可替换的候选结果", "替换");
        return;
    }

    message.mes = refinerData.candidate_text;
    refinerData.applied = true;
    await saveChatConditional();
    updateMessageBlockCompat(resolvedId, message);
    updateMessageButtons(resolvedId);
    updateComparisonPanel(resolvedId);
    toastr.success("已替换为候选结果", "替换");
}

async function restoreOriginal(messageId) {
    const { message, messageId: resolvedId } = getMessageById(messageId);
    if (!message || resolvedId < 0) return;

    const refinerData = message.extra?.response_refiner;
    if (!refinerData?.full_original_text) {
        toastr.warning("没有可恢复的原文", "恢复");
        return;
    }

    message.mes = refinerData.full_original_text;
    refinerData.applied = false;
    await saveChatConditional();
    updateMessageBlockCompat(resolvedId, message);
    updateMessageButtons(resolvedId);
    updateComparisonPanel(resolvedId);
    toastr.success("已恢复原文", "恢复");
}

function makeActionButton(messageId, feature, icon, title, disabled = false) {
    const $button = $("<div>", {
        class: `response-refiner-btn response-refiner-run ${disabled ? "disabled" : ""}`,
        title,
        "data-message-id": messageId,
        "data-feature": feature,
    });
    $button.append($("<i>", { class: `fa-solid ${icon}` }));
    return $button;
}

function updateMessageButtons(messageId) {
    const message = /** @type {RefinerChatMessage | undefined} */ (chat[messageId]);
    if (!isRefinableMessage(message)) return;

    const $messageBlock = $(`#chat .mes[mesid="${messageId}"]`);
    if (!$messageBlock.length) return;

    let $container = $messageBlock.find(".response-refiner-actions");
    if (!$container.length) {
        $container = $('<div class="response-refiner-actions"></div>');
        $messageBlock.find(".mes_text").after($container);
    }
    $container.empty();
    $messageBlock.find(".response-refiner-preview").remove();

    const settings = getSettings();
    const isUser = isUserMessage(message);
    const isBusy = state.busyMessageIds.has(messageId);
    const refinerData = message.extra?.response_refiner;
    const hasCandidate = Boolean(refinerData?.candidate_text);
    const isApplied = Boolean(refinerData?.applied);

    if (isBusy) {
        const $stopBtn = makeActionButton(messageId, "stop", "fa-stop", "停止当前请求");
        $stopBtn.removeClass("response-refiner-run").addClass("response-refiner-stop");
        $container.append(makeActionButton(messageId, "busy", "fa-spinner fa-spin", "处理中", true));
        $container.append($stopBtn);
        if (settings.streamStatusEnabled) {
            renderStatusPanel($container, messageId);
        }
    } else {
        if (settings.features.refineEnabled) {
            $container.append(makeActionButton(messageId, "refine", "fa-wand-magic-sparkles", "润色（含八股文替换/移除提示词）"));
        }
        if (!isUser && settings.features.formatEnabled) {
            $container.append(makeActionButton(messageId, "format", "fa-code", "格式检查和补全修正"));
        }
        if (!isUser && settings.features.completionEnabled) {
            $container.append(makeActionButton(messageId, "completion", "fa-forward", "回复补完"));
        }
        if (!isUser) {
            $container.append(makeActionButton(messageId, "selected", "fa-list-check", "执行设置中已勾选的功能（最多三次请求）"));
        }
    }

    if (hasCandidate && !isApplied) {
        const $applyBtn = $("<div>", {
            class: "response-refiner-btn response-refiner-apply",
            title: "替换为候选结果",
            "data-message-id": messageId,
        }).append($("<i>", { class: "fa-solid fa-check" }));
        $container.append($applyBtn);
    }

    if (hasCandidate && isApplied) {
        const $restoreBtn = $("<div>", {
            class: "response-refiner-btn response-refiner-restore",
            title: "恢复原文",
            "data-message-id": messageId,
        }).append($("<i>", { class: "fa-solid fa-rotate-left" }));
        $container.append($restoreBtn);
    }

    if (hasCandidate && !isApplied) {
        renderInlinePreview($container, refinerData);
    }
}

function renderStatusPanel($container, messageId) {
    const $panel = $("<div>", { class: "response-refiner-status-panel" });
    const $header = $("<div>", { class: "response-refiner-status-header" });
    $header.append($("<strong>", { text: "生成状态" }));
    $header.append($("<span>", { class: "response-refiner-status-stage", text: "等待开始" }));
    const $text = $("<div>", { class: "response-refiner-status-text", text: "等待模型返回..." });
    $panel.append($header, $text);
    $container.after($panel);
    appendStatus(messageId, "等待开始", "");
}

function renderInlinePreview($container, refinerData) {
    const $preview = $("<div>", { class: "response-refiner-preview" });
    const $label = $("<div>", { class: "response-refiner-preview-label" });
    $label.append($("<i>", { class: "fa-solid fa-chevron-right" }));
    $label.append(document.createTextNode(" 处理预览（点击展开）"));

    const $content = $("<div>", { class: "response-refiner-preview-content" }).hide();
    const $grid = $("<div>", { class: "response-refiner-preview-grid" });
    const $left = $("<div>");
    const $right = $("<div>");
    $left.append($("<div>", { class: "response-refiner-preview-title", text: "原文/输入" }));
    $left.append($("<div>", { class: "response-refiner-preview-text" }).text(refinerData.original_text || refinerData.full_original_text || ""));
    $right.append($("<div>", { class: "response-refiner-preview-title", text: "候选结果" }));
    $right.append($("<div>", { class: "response-refiner-preview-text" }).text(refinerData.candidate_text || refinerData.refined_text || ""));
    $grid.append($left, $right);
    $content.append($grid);
    $preview.append($label, $content);
    $container.after($preview);

    $label.on("click", function () {
        const visible = $content.is(":visible");
        $content.slideToggle(200);
        $(this).find("i").toggleClass("fa-chevron-right", visible).toggleClass("fa-chevron-down", !visible);
        $(this).contents().filter(function () { return this.nodeType === 3; }).remove();
        $(this).append(document.createTextNode(visible ? " 处理预览（点击展开）" : " 处理预览（点击折叠）"));
    });
}

function updateComparisonPanel(messageId) {
    if (!state.comparisonPanelVisible) return;

    const message = /** @type {RefinerChatMessage | undefined} */ (chat[messageId]);
    const refinerData = message?.extra?.response_refiner;
    const $panel = $("#response_refiner_comparison_panel");
    const $content = $panel.find(".response-refiner-comparison-content");
    if (!$panel.length || !$content.length) return;

    $content.empty();
    if (!refinerData?.candidate_text) {
        $content.append($("<p>", { text: "暂无处理结果" }));
        return;
    }

    const $left = $("<div>", { class: "response-refiner-comparison-column" });
    $left.append($("<div>", { class: "response-refiner-comparison-title", text: "原文/输入" }));
    $left.append($("<div>", { class: "response-refiner-comparison-text" }).text(refinerData.original_text || refinerData.full_original_text || ""));
    const $right = $("<div>", { class: "response-refiner-comparison-column" });
    $right.append($("<div>", { class: "response-refiner-comparison-title", text: "候选结果" }));
    $right.append($("<div>", { class: "response-refiner-comparison-text" }).text(refinerData.candidate_text || ""));
    $content.append($left, $right);
}

function toggleComparisonPanel() {
    state.comparisonPanelVisible = !state.comparisonPanelVisible;
    const $panel = $("#response_refiner_comparison_panel");
    if (state.comparisonPanelVisible) {
        $panel.slideDown(200);
    } else {
        $panel.slideUp(200);
    }
}

function extractTagPairsFromSample(text) {
    const source = String(text || "");
    const tagRegex = /<([a-zA-Z][\w:-]*)\b[^>]*>([\s\S]*?)<\/\1>/g;
    const pairs = [];
    const seen = new Set();
    let match;
    while ((match = tagRegex.exec(source)) !== null) {
        const tag = match[1];
        if (seen.has(tag)) continue;
        seen.add(tag);
        pairs.push({ tag, startTag: `<${tag}>`, endTag: `</${tag}>`, content: match[2] || "" });
    }
    return pairs;
}

function buildExtractRulesMessages(sampleText, pairs) {
    return [
        { role: "system", content: "你是格式规则提取助手。你只输出 JSON 数组，不输出解释、Markdown 或额外文本。" },
        { role: "user", content: [
            "请根据完整回复样例和已提取的标签，为每个标签生成格式检查规则建议。",
            "输出 JSON 数组，每项包含 name、startTag、endTag、prompt、template。",
            "prompt 要描述该标签内部内容应该满足的规则；template 要包含开始标签、占位内容和结束标签。",
            "已提取标签：",
            JSON.stringify(pairs, null, 2),
            "完整回复样例：",
            sampleText,
        ].join("\n") },
    ];
}

function normalizeExtractedRules(text) {
    const cleaned = stripCodeFence(text);
    const match = cleaned.match(/\[[\s\S]*\]/);
    const parsed = JSON.parse(match ? match[0] : cleaned);
    if (!Array.isArray(parsed)) throw new Error("AI 未返回规则数组");
    return parsed.map(item => ({
        id: crypto?.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random()),
        enabled: true,
        name: String(item.name || "自动提取规则"),
        startTag: String(item.startTag || ""),
        endTag: String(item.endTag || ""),
        prompt: String(item.prompt || ""),
        template: String(item.template || `${item.startTag || ""}\n\n${item.endTag || ""}`),
    })).filter(rule => rule.startTag && rule.endTag);
}

async function runExtractRules() {
    const sample = String($("#response_refiner_extract_source").val() || "");
    const pairs = extractTagPairsFromSample(sample);
    if (!pairs.length) {
        toastr.warning("未在样例中找到成对标签", "自动提取格式规则");
        return;
    }
    const $button = $("#response_refiner_extract_run");
    $button.prop("disabled", true).find("i").addClass("fa-spin");
    try {
        const text = await callAI(buildExtractRulesMessages(sample, pairs));
        state.extractedRules = normalizeExtractedRules(text);
        $("#response_refiner_extract_output").text(JSON.stringify(state.extractedRules, null, 2));
        $("#response_refiner_extract_apply").prop("disabled", !state.extractedRules.length);
        toastr.success(`已生成 ${state.extractedRules.length} 条规则建议`, "自动提取格式规则");
    } catch (error) {
        console.error("[Response Refiner] 自动提取格式规则失败:", error);
        toastr.error(String(error instanceof Error ? error.message : error), "自动提取格式规则");
    } finally {
        $button.prop("disabled", false).find("i").removeClass("fa-spin");
    }
}

function applyExtractedRules() {
    if (!state.extractedRules.length) return;
    const settings = getSettings();
    syncFormatRulesFromDom();
    settings.formatRules.push(...state.extractedRules);
    state.extractedRules = [];
    saveSettings();
    renderFormatRules();
    $("#response_refiner_extract_apply").prop("disabled", true);
    $("#response_refiner_extract_modal").hide();
    toastr.success("已追加自动提取的格式规则", "自动提取格式规则");
}

function formatPromptMessagesForPreview(title, messages) {
    return [
        `## ${title}`,
        ...messages.map((message, index) => `### ${index + 1}. ${message.role}\n${message.content}`),
    ].join("\n\n");
}

function buildSelectedPromptPreview(sourceText, isUser) {
    const settings = getSettings();
    const source = sourceText || "【这里是待处理文本】";
    const blocks = [];
    const textToRefine = isUser ? source : extractTextToRefine(source, settings);
    blocks.push(formatPromptMessagesForPreview("功能1 润色", buildRefineMessages(textToRefine, settings, isUser)));
    if (!isUser) {
        blocks.push(formatPromptMessagesForPreview("功能2 格式检查和补全修正（仅非正文标签，返回替换 JSON）", buildFormatMessages(source, settings, true)));
        blocks.push(formatPromptMessagesForPreview("功能3 回复补完", buildCompletionMessages(source, settings, chat.length, getMissingFormatRules(source, settings))));
        blocks.push([
            "## 执行设置中已勾选的功能（三步组合管线）",
            "脚本先解析当前回复标签状态，然后按需执行：",
            "1. 若启用回复补完且发现缺失标签，则调用功能3补完；否则跳过补完请求。",
            "2. 若启用格式检查，则调用功能2，只修正和补全非正文标签。",
            "3. 若启用润色，则调用功能1，只润色润色捕获正则对应的正文标签内容。",
            "组合执行最多会消耗三次请求，最终由脚本把每个阶段生成的标签内文本回填到原文。",
        ].join("\n"));
    }
    return blocks.join("\n\n====================\n\n");
}

function updatePromptPreview() {
    const isUser = String($("#response_refiner_prompt_preview_type").val() || "assistant") === "user";
    const source = String($("#response_refiner_prompt_preview_source").val() || "");
    $("#response_refiner_prompt_preview_output").text(buildSelectedPromptPreview(source, isUser));
}

function initCollapsibleSections() {
    const settings = getSettings();
    settings.uiCollapsedSections = settings.uiCollapsedSections || {};
    $("#response_refiner_container .response-refiner-section").each(function () {
        const $section = $(this);
        const id = String($section.data("section-id") || "");
        const collapsed = Boolean(settings.uiCollapsedSections[id]);
        $section.toggleClass("collapsed", collapsed);
        $section.find("> .response-refiner-section-body").toggle(!collapsed);
    });
}

function toggleSettingsSection($section) {
    const settings = getSettings();
    const id = String($section.data("section-id") || "");
    const collapsed = !$section.hasClass("collapsed");
    $section.toggleClass("collapsed", collapsed);
    $section.find("> .response-refiner-section-body").slideToggle(150);
    settings.uiCollapsedSections = settings.uiCollapsedSections || {};
    settings.uiCollapsedSections[id] = collapsed;
    saveSettings();
}

function renderProviderOptions() {
    const $select = $("#response_refiner_connection_type");
    $select.empty();
    for (const [key, provider] of Object.entries(PROVIDERS)) {
        $select.append($("<option>", { value: key, text: provider.label }));
    }
}

function renderFormatRules() {
    const settings = getSettings();
    ensureBodyRule(settings, false);
    const body = getBodyRuleInfo(settings, false);
    const $list = $("#response_refiner_format_rules");
    $list.empty();

    (settings.formatRules || []).forEach((rule, index) => {
        const ruleId = rule.id || (rule.id = crypto?.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random()));
        const isBodyRule = body.key && getRuleKey(rule) === body.key;
        const collapsed = Boolean(settings.collapsedFormatRules?.[ruleId]);
        const $card = $("<div>", { class: `response-refiner-rule-card ${collapsed ? "collapsed" : ""}`, "data-index": index, "data-rule-id": ruleId });
        const $header = $("<div>", { class: "response-refiner-rule-header" });
        const $title = $("<div>", { class: "response-refiner-rule-title" });
        const $toggle = $("<span>", { class: "response-refiner-rule-toggle", title: "展开/关闭规则" }).append($("<i>", { class: "fa-solid fa-chevron-down response-refiner-rule-icon" }));
        const $enabled = $("<input>", { type: "checkbox", class: "response-refiner-rule-enabled", title: "启用规则" }).prop("checked", Boolean(rule.enabled));
        const $name = $("<input>", { class: "text_pole response-refiner-rule-name", type: "text", title: "规则名" }).val(rule.name || `规则 ${index + 1}`).prop("disabled", isBodyRule);
        $title.append($toggle, $enabled, $name);
        const $actions = $("<div>", { class: "response-refiner-rule-actions" });
        $actions.append($("<button>", { type: "button", class: "menu_button response-refiner-rule-up", text: "上移" }));
        $actions.append($("<button>", { type: "button", class: "menu_button response-refiner-rule-down", text: "下移" }));
        if (!isBodyRule) {
            $actions.append($("<button>", { type: "button", class: "menu_button response-refiner-rule-delete", text: "删除" }));
        }
        $header.append($title, $actions);

        const $body = $("<div>", { class: "response-refiner-rule-body" }).toggle(!collapsed);
        $body.append($("<label>", { text: "开始标签" }));
        $body.append($("<input>", { class: "text_pole response-refiner-rule-start", type: "text", placeholder: "例如: <content>" }).val(rule.startTag || "").prop("disabled", isBodyRule));
        $body.append($("<label>", { text: "结束标签" }));
        $body.append($("<input>", { class: "text_pole response-refiner-rule-end", type: "text", placeholder: "例如: </content>" }).val(rule.endTag || "").prop("disabled", isBodyRule));
        $body.append($("<label>", { text: "标签内提示词" }));
        $body.append($("<textarea>", { class: "text_pole response-refiner-rule-prompt", rows: 3 }).val(rule.prompt || ""));
        $body.append($("<label>", { text: "标签内容模板" }));
        $body.append($("<textarea>", { class: "text_pole response-refiner-rule-template", rows: 4 }).val(rule.template || ""));
        $card.append($header, $body);
        $list.append($card);
    });
}

function syncFormatRulesFromDom() {
    const settings = getSettings();
    settings.formatRules = [];
    $("#response_refiner_format_rules .response-refiner-rule-card").each(function () {
        const $card = $(this);
        settings.formatRules.push({
            id: String($card.data("rule-id") || (crypto?.randomUUID ? crypto.randomUUID() : Date.now() + Math.random())),
            enabled: $card.find(".response-refiner-rule-enabled").prop("checked"),
            name: String($card.find(".response-refiner-rule-name").val() || ""),
            startTag: String($card.find(".response-refiner-rule-start").val() || ""),
            endTag: String($card.find(".response-refiner-rule-end").val() || ""),
            prompt: String($card.find(".response-refiner-rule-prompt").val() || ""),
            template: String($card.find(".response-refiner-rule-template").val() || ""),
        });
    });
    ensureBodyRule(settings, false);
    saveSettings();
}

function bindSettings() {
    const settings = getSettings();
    renderProviderOptions();
    initCollapsibleSections();

    $(document).on("click keydown", "#response_refiner_container .response-refiner-section-header", function (event) {
        if (event.type === "keydown" && event.key !== "Enter" && event.key !== " ") return;
        if ($(event.target).is("input, textarea, select, button, .menu_button, option")) return;
        event.preventDefault();
        toggleSettingsSection($(this).closest(".response-refiner-section"));
    });

    $("#response_refiner_connection_type")
        .val(getProviderKey())
        .on("change", function () {
            settings.connectionType = String($(this).val());
            saveSettings();
            updateConnectionTypeUI();
        });

    $("#response_refiner_endpoint").on("input", function () {
        getProviderSettings().endpoint = String($(this).val());
        saveSettings();
    });

    $("#response_refiner_model_select").on("change", function () {
        const value = String($(this).val() || "");
        if (value) {
            getProviderSettings().model = value;
            $("#response_refiner_model_manual").val(value);
            saveSettings();
        }
    });

    $("#response_refiner_model_manual").on("input", function () {
        getProviderSettings().model = String($(this).val());
        saveSettings();
    });

    $("#response_refiner_refresh_models").on("click", refreshProviderModels);

    $("#response_refiner_api_key_input").on("input", function () {
        getProviderSettings().apiKey = String($(this).val()).trim();
        saveSettings();
    });

    $("#response_refiner_api_key_toggle").on("click", function () {
        const $input = $("#response_refiner_api_key_input");
        const isPassword = $input.attr("type") === "password";
        $input.attr("type", isPassword ? "text" : "password");
        $(this).find("i").toggleClass("fa-eye", !isPassword).toggleClass("fa-eye-slash", isPassword);
    });

    $("#response_refiner_test_connection").on("click", testConnection);

    $("#response_refiner_refine_enabled").prop("checked", settings.features.refineEnabled).on("change", function () {
        settings.features.refineEnabled = $(this).prop("checked");
        saveSettings();
        refreshAllMessageButtons();
        updatePromptPreview();
    });
    $("#response_refiner_format_enabled").prop("checked", settings.features.formatEnabled).on("change", function () {
        settings.features.formatEnabled = $(this).prop("checked");
        saveSettings();
        refreshAllMessageButtons();
        updatePromptPreview();
    });
    $("#response_refiner_completion_enabled").prop("checked", settings.features.completionEnabled).on("change", function () {
        settings.features.completionEnabled = $(this).prop("checked");
        saveSettings();
        refreshAllMessageButtons();
        updatePromptPreview();
    });

    $("#response_refiner_prompt").val(settings.prompt).on("input", function () {
        settings.prompt = String($(this).val());
        saveSettings();
        updatePromptPreview();
    });
    $("#response_refiner_user_prompt").val(settings.userPrompt).on("input", function () {
        settings.userPrompt = String($(this).val());
        saveSettings();
        updatePromptPreview();
    });
    $("#response_refiner_forbidden_phrases").val(settings.forbiddenPhrases).on("input", function () {
        settings.forbiddenPhrases = String($(this).val());
        saveSettings();
        updatePromptPreview();
    });
    $("#response_refiner_temperature").val(settings.temperature).on("input", function () {
        settings.temperature = Number($(this).val());
        $("#response_refiner_temperature_value").text(settings.temperature.toFixed(2));
        saveSettings();
    });
    $("#response_refiner_temperature_value").text(Number(settings.temperature).toFixed(2));
    $("#response_refiner_max_tokens").val(settings.maxTokens).on("input", function () {
        settings.maxTokens = Number($(this).val());
        saveSettings();
    });
    $("#response_refiner_filter_enabled").prop("checked", settings.filterEnabled).on("change", function () {
        settings.filterEnabled = $(this).prop("checked");
        ensureBodyRule(settings, true);
        saveSettings();
        renderFormatRules();
    });
    $("#response_refiner_filter_regex").val(settings.filterRegex).on("input", function () {
        settings.filterRegex = String($(this).val());
        ensureBodyRule(settings, true);
        saveSettings();
        renderFormatRules();
        updatePromptPreview();
    });
    $("#response_refiner_stream_status_enabled").prop("checked", settings.streamStatusEnabled).on("change", function () {
        settings.streamStatusEnabled = $(this).prop("checked");
        saveSettings();
    });

    $("#response_refiner_completion_chain_regex").val(settings.completionChainRegex || settings.completionOutlineRegex || "").on("input", function () {
        settings.completionChainRegex = String($(this).val());
        settings.completionOutlineRegex = settings.completionChainRegex;
        saveSettings();
    });
    $("#response_refiner_completion_prompt").val(settings.completionPrompt).on("input", function () {
        settings.completionPrompt = String($(this).val());
        saveSettings();
    });
    $("#response_refiner_prompt_preview_type, #response_refiner_prompt_preview_source").on("input change", updatePromptPreview);
    $("#response_refiner_refresh_prompt_preview").on("click", updatePromptPreview);

    $("#response_refiner_open_extract_rules").on("click", function () {
        state.extractedRules = [];
        $("#response_refiner_extract_output").text("暂无提取结果");
        $("#response_refiner_extract_apply").prop("disabled", true);
        $("#response_refiner_extract_modal").show();
    });
    $("#response_refiner_extract_close, #response_refiner_extract_modal .response-refiner-modal-backdrop").on("click", function () {
        $("#response_refiner_extract_modal").hide();
    });
    $("#response_refiner_extract_run").on("click", runExtractRules);
    $("#response_refiner_extract_apply").on("click", applyExtractedRules);

    renderFormatRules();
    $("#response_refiner_add_format_rule").on("click", function () {
        syncFormatRulesFromDom();
        settings.formatRules.push({
            id: crypto?.randomUUID ? crypto.randomUUID() : String(Date.now()),
            enabled: true,
            name: "新标签",
            startTag: "",
            endTag: "",
            prompt: "",
            template: "",
        });
        saveSettings();
        renderFormatRules();
    });

    $(document).on("click", ".response-refiner-rule-toggle", function () {
        const settings = getSettings();
        const $card = $(this).closest(".response-refiner-rule-card");
        const ruleId = String($card.data("rule-id") || "");
        const collapsed = !$card.hasClass("collapsed");
        $card.toggleClass("collapsed", collapsed);
        $card.find(".response-refiner-rule-body").slideToggle(150);
        settings.collapsedFormatRules = settings.collapsedFormatRules || {};
        settings.collapsedFormatRules[ruleId] = collapsed;
        saveSettings();
    });

    $(document).on("input change", "#response_refiner_format_rules input, #response_refiner_format_rules textarea", syncFormatRulesFromDom);
    $(document).on("click", ".response-refiner-rule-delete", function () {
        const index = Number($(this).closest(".response-refiner-rule-card").data("index"));
        syncFormatRulesFromDom();
        settings.formatRules.splice(index, 1);
        saveSettings();
        renderFormatRules();
    });
    $(document).on("click", ".response-refiner-rule-up", function () {
        const index = Number($(this).closest(".response-refiner-rule-card").data("index"));
        syncFormatRulesFromDom();
        if (index > 0) {
            [settings.formatRules[index - 1], settings.formatRules[index]] = [settings.formatRules[index], settings.formatRules[index - 1]];
            saveSettings();
            renderFormatRules();
        }
    });
    $(document).on("click", ".response-refiner-rule-down", function () {
        const index = Number($(this).closest(".response-refiner-rule-card").data("index"));
        syncFormatRulesFromDom();
        if (index < settings.formatRules.length - 1) {
            [settings.formatRules[index + 1], settings.formatRules[index]] = [settings.formatRules[index], settings.formatRules[index + 1]];
            saveSettings();
            renderFormatRules();
        }
    });

    $("#response_refiner_toggle_comparison").on("click", toggleComparisonPanel);
    updateConnectionTypeUI();
    updatePromptPreview();
}

function refreshAllMessageButtons() {
    for (let i = 0; i < chat.length; i++) {
        const message = /** @type {RefinerChatMessage | undefined} */ (chat[i]);
        if (isRefinableMessage(message)) {
            updateMessageButtons(i);
        }
    }
}

async function loadSettingsHtml() {
    const urls = [
        `/scripts/extensions/third-party/${MODULE_NAME}/settings.html`,
        `/scripts/extensions/${MODULE_NAME}/settings.html`,
    ];

    for (const url of urls) {
        try {
            const response = await fetch(url);
            if (response.ok) {
                return response.text();
            }
        } catch (error) {
            console.warn(`[Response Refiner] settings.html 加载失败: ${url}`, error);
        }
    }

    throw new Error(`无法加载 settings.html，请确认插件目录为 third-party/${MODULE_NAME}`);
}

async function addUi() {
    const settingsHtml = await loadSettingsHtml();
    $("#extensions_settings2").append(settingsHtml);

    const $panel = $("<div>", { id: "response_refiner_comparison_panel" }).hide();
    const $header = $("<div>", { class: "response-refiner-comparison-panel-header" });
    $header.append($("<h3>", { text: "Response Refiner 预览" }));
    $header.append($("<button>", { id: "response_refiner_toggle_comparison", class: "menu_button", type: "button" }).append($("<i>", { class: "fa-solid fa-chevron-down" })));
    $panel.append($header, $("<div>", { class: "response-refiner-comparison-content" }).append($("<p>", { text: "暂无处理结果" })));
    $("#chat").before($panel);
}

function onCharacterMessageRendered(messageId) {
    if (typeof messageId !== "number" || messageId < 0) return;
    const message = /** @type {RefinerChatMessage | undefined} */ (chat[messageId]);
    if (isRefinableMessage(message)) {
        updateMessageButtons(messageId);
    }
}

$(document).on("click", ".response-refiner-run", async function () {
    if ($(this).hasClass("disabled")) return;
    const messageId = Number($(this).data("message-id"));
    const feature = String($(this).data("feature"));
    if (feature === "busy") return;
    await requestFeature(messageId, feature);
});

$(document).on("click", ".response-refiner-stop", function () {
    stopRequest(Number($(this).data("message-id")));
});

$(document).on("click", ".response-refiner-apply", async function () {
    await applyCandidate(Number($(this).data("message-id")));
});

$(document).on("click", ".response-refiner-restore", async function () {
    await restoreOriginal(Number($(this).data("message-id")));
});

$(async () => {
    try {
        if (state.initialized) return;
        state.initialized = true;
        getSettings();
        await addUi();
        bindSettings();

        eventSource.on(event_types.CHARACTER_MESSAGE_RENDERED, onCharacterMessageRendered);
        if (event_types.MESSAGE_UPDATED) {
            eventSource.on(event_types.MESSAGE_UPDATED, onCharacterMessageRendered);
        }
        eventSource.on(event_types.CHAT_CHANGED, refreshAllMessageButtons);
        refreshAllMessageButtons();
        console.info("[Response Refiner] 扩展加载完成");
    } catch (error) {
        state.initialized = false;
        console.error("[Response Refiner] 扩展初始化失败:", error);
        toastr?.error?.(error instanceof Error ? error.message : String(error), "Response Refiner 初始化失败");
    }
});

