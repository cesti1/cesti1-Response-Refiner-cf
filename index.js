import * as SillyTavern from "../../../script.js";
import {
    extension_settings,
    renderExtensionTemplateAsync,
} from "../../extensions.js";

const {
    chat,
    event_types,
    eventSource,
    saveChatConditional,
    saveSettingsDebounced,
} = SillyTavern;

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
    userPrompt: "你是一个文本润色助手。请润色用户提供的输入文本，提升表达的清晰度、准确性和流畅度。保持原意不变，直接输出润色后的文本，不要添加前后缀说明。",
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
    completionOutlineRegex: "",
    completionPrompt: "如果当前 AI 回复被截断，请根据格式规则、可用大纲或用户最后一次输入继续补完。只输出续写部分，不要重复已经存在的内容，不要解释。",
    completionContextMessages: 6,
};

/** @typedef {{ mes?: string, is_user?: boolean, is_system?: boolean, extra?: Record<string, any> }} RefinerChatMessage */

const state = {
    initialized: false,
    busyMessageIds: new Set(),
    comparisonPanelVisible: false,
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

function extractOutlineContext(text, regexText) {
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
        console.error("[Response Refiner] 大纲捕获正则错误:", error);
        toastr.warning("大纲捕获正则错误，已改用最后一次用户输入", "回复补完");
        return "";
    }
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
        `开始标签：${rule.startTag}`,
        `结束标签：${rule.endTag}`,
        `标签内提示词：${rule.prompt || "无"}`,
        `标签内容模板：\n${rule.template || `${rule.startTag}\n\n${rule.endTag}`}`,
    ].join("\n")).join("\n\n");
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

function buildFormatMessages(sourceText, settings) {
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

function buildCompletionMessages(sourceText, settings, messageId) {
    const outline = extractOutlineContext(sourceText, settings.completionOutlineRegex);
    const latestUser = getLatestUserMessage(messageId);
    const latestUserText = latestUser.message ? getMessageText(latestUser.message) : "";
    const recentContext = getRecentContext(messageId, settings.completionContextMessages);
    const contextKind = outline ? "大纲捕获正则匹配到的大纲上下文" : "用户最后一次输入上下文";
    const contextText = outline || latestUserText || "未找到可用上下文";

    const system = [
        "你是 AI 回复补完助手。当前 AI 回复可能被截断。",
        "你必须配合格式检查规则继续生成剩余部分，续写内容需要满足标签顺序、开始标签、结束标签、标签内提示词和模板要求。",
        "优先依据大纲上下文补完；如果没有大纲，则依据用户最后一次输入补完。",
        "只输出续写部分，不要重复已经存在的内容，不要输出说明、标题或分析。",
        settings.completionPrompt || "",
        "格式标签规则：",
        buildFormatRulesText(settings),
    ].filter(Boolean).join("\n");

    const user = [
        `上下文类型：${contextKind}`,
        `上下文内容：\n${contextText}`,
        recentContext ? `最近对话摘录：\n${recentContext}` : "",
        `已经生成但可能被截断的 AI 回复：\n${sourceText}`,
        "请从截断处继续补完，只输出续写部分。",
    ].filter(Boolean).join("\n\n---\n\n");

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

    if (!endpoint || !model) {
        throw new Error("请先配置接口地址和模型");
    }
    if (!apiKey) {
        throw new Error("请先配置 API Key");
    }

    if (provider.apiStyle === "gemini") {
        return callGemini(endpoint, model, apiKey, messages, temperature, maxTokens);
    }

    if (provider.apiStyle === "claude") {
        return callClaude(endpoint, model, apiKey, messages, temperature, maxTokens);
    }

    return callOpenAICompatible(providerKey, endpoint, model, apiKey, messages, temperature, maxTokens);
}

async function callOpenAICompatible(providerKey, endpoint, model, apiKey, messages, temperature, maxTokens) {
    const response = await fetch(`${endpoint}/chat/completions`, {
        method: "POST",
        headers: getOpenAICompatibleHeaders(providerKey, apiKey),
        body: JSON.stringify({
            model,
            messages,
            temperature,
            max_tokens: maxTokens,
        }),
    });

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

async function callGemini(endpoint, model, apiKey, messages, temperature, maxTokens) {
    const modelName = model.startsWith("models/") ? model : `models/${model}`;
    const systemText = messages.filter(item => item.role === "system").map(item => item.content).join("\n\n");
    const userText = messages.filter(item => item.role !== "system").map(item => item.content).join("\n\n");
    const response = await fetch(`${endpoint}/${modelName}:generateContent?key=${encodeURIComponent(apiKey)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
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

async function callClaude(endpoint, model, apiKey, messages, temperature, maxTokens) {
    const system = messages.filter(item => item.role === "system").map(item => item.content).join("\n\n");
    const user = messages.filter(item => item.role !== "system").map(item => item.content).join("\n\n");
    const response = await fetch(`${endpoint}/v1/messages`, {
        method: "POST",
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

async function runRefine(originalText, settings, isUser) {
    const textToRefine = isUser ? originalText : extractTextToRefine(originalText, settings);
    const refinedText = await callAI(buildRefineMessages(textToRefine, settings, isUser));
    const finalText = isUser ? refinedText : replaceRefinedText(originalText, refinedText, settings);
    return {
        stage: "refine",
        original_text: textToRefine,
        refined_text: refinedText,
        candidate_text: finalText,
    };
}

async function runFormat(sourceText, settings) {
    const formattedText = await callAI(buildFormatMessages(sourceText, settings));
    return {
        stage: "format",
        original_text: sourceText,
        refined_text: formattedText,
        candidate_text: formattedText,
    };
}

async function runCompletion(sourceText, settings, messageId) {
    const continuation = await callAI(buildCompletionMessages(sourceText, settings, messageId));
    const joined = sourceText + continuation;
    const formatted = getEnabledFormatRules(settings).length
        ? await callAI(buildFormatMessages(joined, settings))
        : joined;
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
    let stages = feature === "selected" ? getSelectedStages(settings, isUser) : [feature];
    stages = stages.filter(stage => stage !== "completion" || !isUser);
    stages = stages.filter(stage => stage !== "format" || !isUser);

    if (!stages.length) {
        toastr.warning("没有已启用的可执行功能", "Response Refiner");
        return;
    }

    state.busyMessageIds.add(resolvedId);
    updateMessageButtons(resolvedId);

    try {
        const fullOriginalText = getMessageText(message);
        let workingText = fullOriginalText;
        const stageResults = [];

        for (const stage of stages) {
            let result;
            if (stage === "refine") {
                result = await runRefine(workingText, settings, isUser);
            } else if (stage === "format") {
                result = await runFormat(workingText, settings);
            } else if (stage === "completion") {
                result = await runCompletion(workingText, settings, resolvedId);
            } else {
                continue;
            }
            stageResults.push(result);
            workingText = result.candidate_text;
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
        console.error("[Response Refiner] 处理失败:", error);
        toastr.error(String(error instanceof Error ? error.message : error), "Response Refiner");
    } finally {
        state.busyMessageIds.delete(resolvedId);
        updateMessageButtons(resolvedId);
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
        $container.append(makeActionButton(messageId, "busy", "fa-spinner fa-spin", "处理中", true));
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
        $container.append(makeActionButton(messageId, "selected", "fa-list-check", "执行设置中已勾选的功能"));
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

function renderProviderOptions() {
    const $select = $("#response_refiner_connection_type");
    $select.empty();
    for (const [key, provider] of Object.entries(PROVIDERS)) {
        $select.append($("<option>", { value: key, text: provider.label }));
    }
}

function renderFormatRules() {
    const settings = getSettings();
    const $list = $("#response_refiner_format_rules");
    $list.empty();

    (settings.formatRules || []).forEach((rule, index) => {
        const $card = $("<div>", { class: "response-refiner-rule-card", "data-index": index });
        const $header = $("<div>", { class: "response-refiner-rule-header" });
        const $enabled = $("<label>", { class: "checkbox_label" }).append(
            $("<input>", { type: "checkbox", class: "response-refiner-rule-enabled" }).prop("checked", Boolean(rule.enabled)),
            $("<span>", { text: `规则 ${index + 1}` }),
        );
        const $actions = $("<div>", { class: "response-refiner-rule-actions" });
        $actions.append($("<button>", { type: "button", class: "menu_button response-refiner-rule-up", text: "上移" }));
        $actions.append($("<button>", { type: "button", class: "menu_button response-refiner-rule-down", text: "下移" }));
        $actions.append($("<button>", { type: "button", class: "menu_button response-refiner-rule-delete", text: "删除" }));
        $header.append($enabled, $actions);

        $card.append($header);
        $card.append($("<label>", { text: "名称" }));
        $card.append($("<input>", { class: "text_pole response-refiner-rule-name", type: "text" }).val(rule.name || ""));
        $card.append($("<label>", { text: "开始标签" }));
        $card.append($("<input>", { class: "text_pole response-refiner-rule-start", type: "text", placeholder: "例如: <content>" }).val(rule.startTag || ""));
        $card.append($("<label>", { text: "结束标签" }));
        $card.append($("<input>", { class: "text_pole response-refiner-rule-end", type: "text", placeholder: "例如: </content>" }).val(rule.endTag || ""));
        $card.append($("<label>", { text: "标签内提示词" }));
        $card.append($("<textarea>", { class: "text_pole response-refiner-rule-prompt", rows: 3 }).val(rule.prompt || ""));
        $card.append($("<label>", { text: "标签内容模板" }));
        $card.append($("<textarea>", { class: "text_pole response-refiner-rule-template", rows: 4 }).val(rule.template || ""));
        $list.append($card);
    });
}

function syncFormatRulesFromDom() {
    const settings = getSettings();
    settings.formatRules = [];
    $("#response_refiner_format_rules .response-refiner-rule-card").each(function () {
        const $card = $(this);
        settings.formatRules.push({
            id: crypto?.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random()),
            enabled: $card.find(".response-refiner-rule-enabled").prop("checked"),
            name: String($card.find(".response-refiner-rule-name").val() || ""),
            startTag: String($card.find(".response-refiner-rule-start").val() || ""),
            endTag: String($card.find(".response-refiner-rule-end").val() || ""),
            prompt: String($card.find(".response-refiner-rule-prompt").val() || ""),
            template: String($card.find(".response-refiner-rule-template").val() || ""),
        });
    });
    saveSettings();
}

function bindSettings() {
    const settings = getSettings();
    renderProviderOptions();

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
    });
    $("#response_refiner_format_enabled").prop("checked", settings.features.formatEnabled).on("change", function () {
        settings.features.formatEnabled = $(this).prop("checked");
        saveSettings();
        refreshAllMessageButtons();
    });
    $("#response_refiner_completion_enabled").prop("checked", settings.features.completionEnabled).on("change", function () {
        settings.features.completionEnabled = $(this).prop("checked");
        saveSettings();
        refreshAllMessageButtons();
    });

    $("#response_refiner_prompt").val(settings.prompt).on("input", function () {
        settings.prompt = String($(this).val());
        saveSettings();
    });
    $("#response_refiner_user_prompt").val(settings.userPrompt).on("input", function () {
        settings.userPrompt = String($(this).val());
        saveSettings();
    });
    $("#response_refiner_forbidden_phrases").val(settings.forbiddenPhrases).on("input", function () {
        settings.forbiddenPhrases = String($(this).val());
        saveSettings();
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
        saveSettings();
    });
    $("#response_refiner_filter_regex").val(settings.filterRegex).on("input", function () {
        settings.filterRegex = String($(this).val());
        saveSettings();
    });

    $("#response_refiner_completion_outline_regex").val(settings.completionOutlineRegex).on("input", function () {
        settings.completionOutlineRegex = String($(this).val());
        saveSettings();
    });
    $("#response_refiner_completion_prompt").val(settings.completionPrompt).on("input", function () {
        settings.completionPrompt = String($(this).val());
        saveSettings();
    });
    $("#response_refiner_completion_context_messages").val(settings.completionContextMessages).on("input", function () {
        settings.completionContextMessages = Number($(this).val());
        saveSettings();
    });

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
}

function refreshAllMessageButtons() {
    for (let i = 0; i < chat.length; i++) {
        const message = /** @type {RefinerChatMessage | undefined} */ (chat[i]);
        if (isRefinableMessage(message)) {
            updateMessageButtons(i);
        }
    }
}

async function addUi() {
    const settingsHtml = await renderExtensionTemplateAsync(MODULE_NAME, "settings");
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

$(document).on("click", ".response-refiner-apply", async function () {
    await applyCandidate(Number($(this).data("message-id")));
});

$(document).on("click", ".response-refiner-restore", async function () {
    await restoreOriginal(Number($(this).data("message-id")));
});

$(async () => {
    if (state.initialized) return;
    state.initialized = true;
    getSettings();
    await addUi();
    bindSettings();

    eventSource.on(event_types.CHARACTER_MESSAGE_RENDERED, onCharacterMessageRendered);
    eventSource.on(event_types.MESSAGE_UPDATED, onCharacterMessageRendered);
    eventSource.on(event_types.CHAT_CHANGED, refreshAllMessageButtons);
    refreshAllMessageButtons();
});

