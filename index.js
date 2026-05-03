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

const MODULE_NAME = "response-refiner";
const SETTINGS_KEY = /** @type {const} */ ("response_refiner");
const VERSION = "0.1.0-beta.1";
const START_OF_TEXT_TAG = "__RESPONSE_REFINER_START_OF_TEXT__";
const START_OF_TEXT_LABEL = "文本开头";

const DEFAULT_EXTRACT_RULES_PROMPT = [
  "请根据完整回复样例和已提取的标签，为每个标签生成格式检查规则建议。",
  "输出 JSON 数组，每项包含 name、startTag、endTag、prompt、template；如果输入项带 startAnchored:true，必须原样保留 startTag、startAnchored 和 forceTop。",
  "prompt 要描述该标签内部内容应该满足的规则；template 要包含开始标签、占位内容和结束标签；若 startAnchored:true，则 template 从占位内容开始并以结束标签结束，不要编造开始标签。",
].join("\n");

const DEFAULT_EXTRACT_CHAIN_RULE_PROMPT = [
  "请分析样例中 {{startTag}} 到 {{endTag}} 的已有思维链或推理结构，并输出严格 JSON。",
  "JSON 必须包含 name、steps、boundary_features、output_rule。steps 要保留样例中已经存在的步骤名称、编号顺序、层级和边界特征；不要照抄具体私密思考内容。",
  "boundary_features 要说明开始边界、结束边界、是否文本起点块、段落/编号/内部标签结构。",
  "output_rule.prompt 要整理成可复用格式规则，要求后续生成继续遵守相同步骤顺序、内部标签块、段落边界和闭合关系；output_rule.template 要给出可复用模板。",
  "如果是文本起点块，不得编造不存在的开始标签；最终必须确保以 {{endTag}} 正确闭合。",
].join("\n");

const DEFAULT_REFINE_SYSTEM_TEMPLATE =
  "{{basePrompt}}\n\n{{forbiddenInstruction}}";
const DEFAULT_REFINE_USER_TEMPLATE = "{{textToRefine}}";
const DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE = [
  "你是 AI 回复格式检查和补全修正助手，也是严格的格式标签内容修正器。",
  "你只处理非正文标签；正文标签由润色功能处理，禁止输出或修改正文规则对应的内容。",
  "你的任务是修正每个非正文规则对应标签内部的文本，使其满足标签提示词和模板要求。",
  "为了减少输出 token，你绝对不要输出完整回复、开始标签、结束标签、解释或 Markdown 代码块。",
  "你必须只输出一个 JSON 对象，键为规则ID，值为该规则标签内修正后的纯文本。",
  "JSON 值必须是最终要写入该标签内部的实际内容：不要照搬标签内容模板，不要输出模板占位符，不要包含开始标签或结束标签；文本起点块也不要输出“文本开头”字样和结束标签。",
  "如果某个非正文标签不存在但规则要求补全，请仍在对应规则ID里输出应插入的标签内文本；如果现有内容已经正确，可原样返回现有标签内内容。",
  "脚本会按照格式标签规则顺序重排标签块；若开始标签为“文本开头”，该块必须位于回复最开头并由脚本补上结束标签。",
].join("\n");
const DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE = [
  "需要判断、检查、补完的格式标签规则：",
  "{{formatRules}}",
  "完整原文：",
  "{{sourceText}}",
  "当前已提取标签内容 JSON：",
  "{{currentSegmentsJson}}",
  "任务说明：请仅修正或补全上述非正文规则对应的标签内部文本；不要输出完整原文，不要输出开始标签或结束标签。",
  "严格 JSON 输出格式：输出一个对象，键必须是规则ID，值必须是对应标签内部最终文本；值不能为空字符串，已有内容正确时原样返回。",
  "最后约束：只输出 JSON，不要输出解释、Markdown 代码块或任何额外文本。",
].join("\n");
const DEFAULT_FORMAT_FULL_SYSTEM_TEMPLATE = [
  "你是 AI 回复格式检查和补全修正助手。",
  "你只能根据用户提供的当前 AI 回复文本、格式标签规则、标签提示词和模板进行修正。",
  "需要补全缺失标签、修复错误顺序、闭合不完整标签，并按模板与提示词修正标签内内容。",
  "不得引入无关设定，不得解释修改过程。最终只输出完整修正后的 AI 回复。",
  "格式标签规则：",
  "{{formatRules}}",
].join("\n");
const DEFAULT_FORMAT_FULL_USER_TEMPLATE = "{{sourceText}}";
const DEFAULT_COMPLETION_SYSTEM_TEMPLATE = [
  "你是 AI 回复补完助手。当前 AI 回复可能被截断或缺少部分标签。",
  "任务边界：只根据用户消息、当前回复、捕获到的思维链和格式标签规则生成续写或缺失片段。",
  "安全约束：不得引入无关设定，不得解释过程，不得输出标题、分析或 Markdown 代码块，不得重复已经存在的完整内容。",
  "{{completionRequirement}}",
  "如果需要补全缺失标签，必须生成该标签的实际内容；禁止照搬标签内容模板、示例占位符或 [时间]/[地点]/[选项内容] 这类占位符。",
  "如果当前回复已经包含完整思维链结束标签，不要再次输出 </think> 或思维链模板；只有当前最后位置确实处在思维链内部时才补完思维链结束标签。",
  "{{completionPrompt}}",
].join("\n");
const DEFAULT_COMPLETION_USER_TEMPLATE = [
  "全部格式标签规则：\n{{formatRules}}",
  "完整原文 / 当前已经生成但可能被截断的 AI 回复：\n{{sourceText}}",
  "上一条用户回复：\n{{previousUserText}}",
  "捕获到的思维链：\n{{chainContext}}",
  "未闭合/缺失标签判断：\n{{unclosedInfo}}\n\n当前缺失或需要关注的标签规则：\n{{missingRules}}",
  "补完位置说明：请从完整原文末尾继续追加，不要改写或重复原文已有部分。\n{{completionUserInstruction}}",
  "最后约束：只输出需要追加到当前回复末尾的内容，不要输出完整回复、解释、标题或 Markdown 代码块。",
].join("\n\n---\n\n");

const DEFAULT_ENDPOINTS = {
  openrouter: "https://openrouter.ai/api/v1",
  gemini: "https://generativelanguage.googleapis.com/v1beta",
  claude: "https://api.anthropic.com",
  openai: "https://api.openai.com/v1",
  deepseek: "https://api.deepseek.com/v1",
};

const PROVIDERS = {
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
  connectionType: "direct",
  providers: DEFAULT_PROVIDER_SETTINGS,
  prompt:
    "你是一个文本润色助手。你只能处理用户提供的文本，不引入对话历史、设定扩写、旁白说明或解释。保持原意不变，只提升措辞、流畅度、节奏、文笔与可读性。直接输出处理后的正文，不要添加前后缀说明。",
  userPrompt:
    "你是一个用户输入润色助手。用户输入通常只是一小段简单描述、指令或台词。你只能在原始输入范围内改善措辞、错别字、语序和清晰度，不得扩写成章节、正文、剧情段落、完整回复或新增设定。保持长度与信息量基本接近原文，直接输出润色后的用户输入，不要添加解释、标题、前后缀或引号。",
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
  completionPrompt:
    "如果当前 AI 回复被截断，请根据格式规则、可用思维链或需要补完回复的上一条用户消息继续补完。只输出续写或缺失标签内容，不要解释。",
  refineSystemTemplate: DEFAULT_REFINE_SYSTEM_TEMPLATE,
  refineUserTemplate: DEFAULT_REFINE_USER_TEMPLATE,
  formatReplacementSystemTemplate: DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE,
  formatReplacementUserTemplate: DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE,
  formatFullSystemTemplate: DEFAULT_FORMAT_FULL_SYSTEM_TEMPLATE,
  formatFullUserTemplate: DEFAULT_FORMAT_FULL_USER_TEMPLATE,
  completionSystemTemplate: DEFAULT_COMPLETION_SYSTEM_TEMPLATE,
  completionUserTemplate: DEFAULT_COMPLETION_USER_TEMPLATE,
  extractRulesPrompt: DEFAULT_EXTRACT_RULES_PROMPT,
  extractChainRulePrompt: DEFAULT_EXTRACT_CHAIN_RULE_PROMPT,
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
  statusHeartbeats: new Map(),
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
  migratePromptTemplates(settings);
  settings.version = VERSION;
  return settings;
}

function migrateLegacySettings(settings) {
  settings.providers =
    settings.providers || deepClone(DEFAULT_PROVIDER_SETTINGS);
  settings.providers.direct =
    settings.providers.direct || deepClone(DEFAULT_PROVIDER_SETTINGS.direct);

  const legacyProviderKey = "cloud" + "flare";
  const legacyModelKey = legacyProviderKey + "Model";
  const legacyApiKeyKey = legacyProviderKey + "ApiKey";
  const legacyProvider = settings.providers?.[legacyProviderKey] || {};
  if (
    settings.endpoint !== undefined ||
    settings[legacyModelKey] !== undefined ||
    settings[legacyApiKeyKey] !== undefined ||
    settings.connectionType === legacyProviderKey ||
    legacyProvider.endpoint ||
    legacyProvider.model ||
    legacyProvider.apiKey
  ) {
    settings.providers.direct.endpoint =
      settings.providers.direct.endpoint ||
      legacyProvider.endpoint ||
      settings.endpoint ||
      "";
    settings.providers.direct.model =
      settings.providers.direct.model ||
      legacyProvider.model ||
      settings[legacyModelKey] ||
      settings.model ||
      "";
    settings.providers.direct.apiKey =
      settings.providers.direct.apiKey ||
      legacyProvider.apiKey ||
      settings[legacyApiKeyKey] ||
      settings.apiKey ||
      "";
    settings.connectionType = "direct";
  }
  delete settings.providers[legacyProviderKey];

  if (
    settings.directEndpoint !== undefined ||
    settings.directModel !== undefined ||
    settings.directApiKey !== undefined
  ) {
    settings.providers.direct.endpoint =
      settings.providers.direct.endpoint || settings.directEndpoint || "";
    settings.providers.direct.model =
      settings.providers.direct.model || settings.directModel || "";
    settings.providers.direct.apiKey =
      settings.providers.direct.apiKey || settings.directApiKey || "";
  }

  if (
    settings.openrouterEndpoint !== undefined ||
    settings.openrouterModel !== undefined ||
    settings.openrouterApiKey !== undefined
  ) {
    settings.providers.openrouter =
      settings.providers.openrouter ||
      deepClone(DEFAULT_PROVIDER_SETTINGS.openrouter);
    settings.providers.openrouter.endpoint =
      settings.providers.openrouter.endpoint ||
      settings.openrouterEndpoint ||
      PROVIDERS.openrouter.defaultEndpoint;
    settings.providers.openrouter.model =
      settings.providers.openrouter.model || settings.openrouterModel || "";
    settings.providers.openrouter.apiKey =
      settings.providers.openrouter.apiKey || settings.openrouterApiKey || "";
  }

  if (!PROVIDERS[settings.connectionType]) {
    settings.connectionType = "direct";
  }

  if (
    settings.completionChainRegex === undefined &&
    settings.completionOutlineRegex !== undefined
  ) {
    settings.completionChainRegex = settings.completionOutlineRegex || "";
  }
}

function migratePromptTemplates(settings) {
  settings.refineSystemTemplate =
    String(settings.refineSystemTemplate || "").trim() ||
    DEFAULT_REFINE_SYSTEM_TEMPLATE;
  settings.refineUserTemplate =
    String(settings.refineUserTemplate || "") || DEFAULT_REFINE_USER_TEMPLATE;
  settings.formatReplacementSystemTemplate =
    String(settings.formatReplacementSystemTemplate || "").trim() ||
    DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE;
  settings.formatReplacementUserTemplate =
    String(settings.formatReplacementUserTemplate || "") ||
    DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE;
  settings.formatFullSystemTemplate =
    String(settings.formatFullSystemTemplate || "").trim() ||
    DEFAULT_FORMAT_FULL_SYSTEM_TEMPLATE;
  settings.formatFullUserTemplate =
    String(settings.formatFullUserTemplate || "") ||
    DEFAULT_FORMAT_FULL_USER_TEMPLATE;
  settings.completionSystemTemplate =
    String(settings.completionSystemTemplate || "").trim() ||
    DEFAULT_COMPLETION_SYSTEM_TEMPLATE;
  settings.completionUserTemplate =
    String(settings.completionUserTemplate || "") ||
    DEFAULT_COMPLETION_USER_TEMPLATE;

  const formatUser = String(settings.formatReplacementUserTemplate || "");
  if (
    formatUser.includes("你是 AI 回复格式检查和补全修正助手。") ||
    formatUser.includes("JSON 值必须是最终要写入该标签内部的实际内容") ||
    !formatUser.includes("需要判断、检查、补完的格式标签规则")
  ) {
    settings.formatReplacementUserTemplate =
      DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE;
  }

  const completionSystem = String(settings.completionSystemTemplate || "");
  if (
    completionSystem.includes("全部格式标签规则：") ||
    completionSystem.includes("补完依据：需要补完回复的上一条用户消息")
  ) {
    settings.completionSystemTemplate = DEFAULT_COMPLETION_SYSTEM_TEMPLATE;
  }

  const completionUser = String(settings.completionUserTemplate || "");
  if (
    !completionUser.includes("全部格式标签规则") ||
    !completionUser.includes("最后约束：只输出需要追加到当前回复末尾的内容")
  ) {
    settings.completionUserTemplate = DEFAULT_COMPLETION_USER_TEMPLATE;
  }
}

function saveSettings() {
  saveSettingsDebounced();
}

function getProviderKey() {
  const settings = getSettings();
  if (!PROVIDERS[settings.connectionType]) {
    settings.connectionType = "direct";
  }
  return settings.connectionType;
}

function getProviderSettings(providerKey = getProviderKey()) {
  const settings = getSettings();
  settings.providers[providerKey] =
    settings.providers[providerKey] ||
    deepClone(DEFAULT_PROVIDER_SETTINGS[providerKey]);
  mergeDefaults(
    settings.providers[providerKey],
    DEFAULT_PROVIDER_SETTINGS[providerKey],
  );
  return settings.providers[providerKey];
}

function normalizeEndpoint(endpoint) {
  return String(endpoint || "")
    .trim()
    .replace(/\/+$/, "");
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
  const endpoint = normalizeEndpoint(providerSettings.endpoint);
  if (endpoint) {
    return endpoint;
  }
  if (providerKey === "direct") {
    return "";
  }
  return normalizeEndpoint(provider.defaultEndpoint || "");
}

function getSettingValue(value, fallback = "") {
  const normalized = String(value ?? "");
  return normalized.trim() ? normalized : fallback;
}

function syncStructureTemplateInputs(settings = getSettings()) {
  $("#response_refiner_refine_user_template").val(
    getSettingValue(settings.refineUserTemplate, DEFAULT_REFINE_USER_TEMPLATE),
  );
  $("#response_refiner_format_replacement_user_template").val(
    getSettingValue(
      settings.formatReplacementUserTemplate,
      DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE,
    ),
  );
  $("#response_refiner_format_full_user_template").val(
    getSettingValue(
      settings.formatFullUserTemplate,
      DEFAULT_FORMAT_FULL_USER_TEMPLATE,
    ),
  );
  $("#response_refiner_completion_user_template").val(
    getSettingValue(
      settings.completionUserTemplate,
      DEFAULT_COMPLETION_USER_TEMPLATE,
    ),
  );
}

function setText($element, text) {
  $element.text(String(text ?? ""));
}

function splitForbiddenPhrases(settings) {
  return String(settings.forbiddenPhrases || "")
    .split(/[，,]/)
    .map((item) => item.trim())
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
    phrases.map((item) => `- ${item}`).join("\n"),
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
  const message = /** @type {RefinerChatMessage | undefined} */ (
    chat[messageId]
  );
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

function getLatestAssistantMessageText() {
  for (let i = chat.length - 1; i >= 0; i--) {
    const message = /** @type {RefinerChatMessage | undefined} */ (chat[i]);
    if (isAssistantMessage(message)) {
      return getMessageText(message);
    }
  }
  return "";
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
  } catch (_error) {
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
  } catch (_error) {
    return refinedText;
  }
}

function normalizeRegexFlags(flags, defaultFlags = "g") {
  const seen = new Set();
  return String(flags || defaultFlags || "")
    .split("")
    .concat(String(defaultFlags || "").split(""))
    .filter(
      (flag) => /[dgimsuvy]/.test(flag) && !seen.has(flag) && seen.add(flag),
    )
    .join("");
}

function unescapeHtmlEntities(text) {
  return String(text || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}
function normalizeRegexBody(body) {
  return String(body || "")
    .replace(/\\\\([sSdDwWbB])/g, "\\$1")
    .replace(/\\\\\//g, "\\/");
}

function parseUserRegex(regexText, defaultFlags = "g") {
  const raw = unescapeHtmlEntities(regexText).trim();
  if (!raw) return null;
  const literal = raw.match(/^\/(.*)\/([a-z]*)$/i);
  if (literal) {
    return new RegExp(
      normalizeRegexBody(literal[1]),
      normalizeRegexFlags(literal[2], defaultFlags),
    );
  }
  return new RegExp(
    normalizeRegexBody(raw),
    normalizeRegexFlags(defaultFlags, defaultFlags),
  );
}

function isOnlyXmlLikeTag(value) {
  return /^<\/?[a-zA-Z][\w:-]*\b[^>]*>$/.test(String(value || "").trim());
}

function getPreferredRegexCapture(match) {
  if (!match) return "";
  const captures = match
    .slice(1)
    .filter((value) => value !== undefined && String(value).trim());
  const contentCapture = captures.find((value) => !isOnlyXmlLikeTag(value));
  if (contentCapture !== undefined) {
    return contentCapture;
  }
  return captures.length ? captures[0] : match[0];
}

function getRegexMatches(text, regexText) {
  const pattern = String(regexText || "").trim();
  if (!pattern) {
    return [];
  }

  const regex = parseUserRegex(pattern, "g");
  if (!regex) return [];
  const matches = [];
  let match;
  while ((match = regex.exec(text)) !== null) {
    const captured = getPreferredRegexCapture(match);
    if (String(captured || "").trim()) {
      matches.push(captured);
    }
    if (match[0] === "") {
      regex.lastIndex++;
    }
  }
  return matches;
}

function findStartAnchoredChainByEndTag(text, regexText) {
  const pattern = unescapeHtmlEntities(regexText).trim();
  const endTagMatch =
    pattern.match(/<\\?\/([a-zA-Z][\w:-]*)>/) ||
    pattern.match(/<\/([a-zA-Z][\w:-]*)>/);
  if (!endTagMatch) return "";
  const endTag = `</${endTagMatch[1]}>`;
  const source = String(text || "");
  const index = source.indexOf(endTag);
  if (index <= 0) return "";
  return source.slice(0, index).trim();
}

function getTextMatchCandidates(text) {
  const raw = String(text || "");
  const htmlDecoded = unescapeHtmlEntities(raw);
  return [...new Set([raw, htmlDecoded])];
}

function extractChainContext(text, regexText) {
  try {
    const matches = getTextMatchCandidates(text).flatMap((candidate) =>
      getRegexMatches(candidate, regexText),
    );
    if (matches.length) {
      return String(matches[matches.length - 1]).trim();
    }
    const fallback = getTextMatchCandidates(text)
      .map((candidate) => findStartAnchoredChainByEndTag(candidate, regexText))
      .find(Boolean);
    return fallback || "";
  } catch (_error) {
    toastr.warning("思维链捕获正则错误，已仅使用上一条用户消息", "回复补完");
    return "";
  }
}

function inferBodyTagsFromFilterRegex(regexText) {
  const text = String(regexText || "").trim();
  const match =
    text.match(/^<([a-zA-Z][\w:-]*)\b[^>]*>\s*\(\[\\s\\S\]\*\?\)\s*<\/\1>$/) ||
    text.match(/^<([a-zA-Z][\w:-]*)\b[^>]*>\(\[\\s\\S\]\*\?\)<\/\1>$/) ||
    text.match(/^<([a-zA-Z][\w:-]*)\b[^>]*>\(\.\*\?\)<\/\1>$/);
  if (!match) return null;
  return { tag: match[1], startTag: `<${match[1]}>`, endTag: `</${match[1]}>` };
}

function getBodyRuleInfo(settings, notify = false) {
  const inferred = inferBodyTagsFromFilterRegex(settings.filterRegex);
  if (!inferred) {
    if (notify && settings.filterEnabled) {
      toastr.warning(
        "润色捕获正则过复杂，无法自动识别正文标签；请确保正文规则标签与润色捕获正则一致。",
        "正文规则",
      );
    }
    return { rule: null, key: "", startTag: "", endTag: "", inferred: null };
  }

  let rule = (settings.formatRules || []).find(
    (item) =>
      item &&
      item.startTag === inferred.startTag &&
      item.endTag === inferred.endTag,
  );
  if (!rule) {
    rule = (settings.formatRules || []).find(
      (item) => item && (item.id === "content" || item.name === "正文"),
    );
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
  if (!rule.template)
    rule.template = `${inferred.startTag}
这里填写正文内容
${inferred.endTag}`;
  return {
    rule,
    key: getRuleKey(rule),
    startTag: inferred.startTag,
    endTag: inferred.endTag,
    inferred,
  };
}

function ensureBodyRule(settings, notify = false) {
  return getBodyRuleInfo(settings, notify);
}

function isStartAnchoredRule(rule) {
  const start = String(rule?.startTag || "");
  const realStart = String(rule?.realStartTag || "");
  return (
    start === START_OF_TEXT_TAG ||
    start === START_OF_TEXT_LABEL ||
    realStart === START_OF_TEXT_TAG ||
    Boolean(rule?.startAnchored || rule?.forceTop)
  );
}

function isStartOnlyAnchoredRule(rule) {
  const realStart = String(rule?.realStartTag || "");
  const start = String(rule?.startTag || "");
  return (
    isStartAnchoredRule(rule) &&
    (!realStart ||
      realStart === START_OF_TEXT_TAG ||
      realStart === START_OF_TEXT_LABEL ||
      start === START_OF_TEXT_TAG ||
      start === START_OF_TEXT_LABEL)
  );
}

function getRuleStartLabel(rule) {
  if (isStartOnlyAnchoredRule(rule)) return START_OF_TEXT_LABEL;
  if (isStartAnchoredRule(rule) && rule.realStartTag)
    return String(rule.realStartTag);
  return String(rule?.startTag || "");
}

function getEnabledFormatRules(settings) {
  return (settings.formatRules || [])
    .filter(
      (rule) =>
        rule &&
        rule.enabled &&
        (rule.startTag || rule.realStartTag || isStartAnchoredRule(rule)) &&
        rule.endTag,
    )
    .map((rule, index) => ({ ...rule, order: index + 1 }));
}

function buildFormatRulesText(settings) {
  const rules = getEnabledFormatRules(settings);
  if (!rules.length) {
    return "未配置启用的格式标签规则。";
  }

  return rules
    .map((rule) =>
      [
        `${rule.order}. ${rule.name || "未命名标签"}`,
        `规则ID：${getRuleKey(rule)}`,
        `开始标签：${getRuleStartLabel(rule)}`,
        `结束标签：${rule.endTag}`,
        `标签内提示词：${rule.prompt || "无"}`,
        `标签内容模板：\n${rule.template || `${rule.startTag}\n\n${rule.endTag}`}`,
      ].join("\n"),
    )
    .join("\n\n");
}

function getRuleKey(rule) {
  return String(
    rule.id || rule.name || rule.endTag || rule.startTag || "rule",
  ).replace(/[^a-zA-Z0-9_\-\u4e00-\u9fa5]/g, "_");
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && bStart < aEnd;
}

function isRangeOccupied(start, end, occupiedRanges) {
  return occupiedRanges.some((range) =>
    rangesOverlap(start, end, range.start, range.end),
  );
}

function findRuleMatchesOutsideOccupied(source, rule, occupiedRanges) {
  const start = String(rule.realStartTag || rule.startTag || "");
  const end = String(rule.endTag || "");
  if (!end) return [];

  const matches = [];
  if (isStartOnlyAnchoredRule(rule)) {
    const endIndex = source.indexOf(end);
    if (endIndex >= 0) {
      const blockEnd = endIndex + end.length;
      if (!isRangeOccupied(0, blockEnd, occupiedRanges)) {
        matches.push({
          contentStart: 0,
          contentEnd: endIndex,
          blockStart: 0,
          blockEnd,
        });
      }
    }
    return matches;
  }

  const actualStart = isStartAnchoredRule(rule)
    ? String(rule.realStartTag || "")
    : start;
  if (!actualStart) return [];

  let searchFrom = 0;
  let guard = 0;
  while (searchFrom < source.length && guard++ < 10000) {
    const startIndex = source.indexOf(actualStart, searchFrom);
    if (startIndex < 0) break;
    const contentStart = startIndex + actualStart.length;
    const endIndex = source.indexOf(end, contentStart);
    if (endIndex < 0) break;
    const blockEnd = endIndex + end.length;
    if (!isRangeOccupied(startIndex, blockEnd, occupiedRanges)) {
      matches.push({
        contentStart,
        contentEnd: endIndex,
        blockStart: startIndex,
        blockEnd,
      });
    }
    searchFrom = Math.max(startIndex + 1, blockEnd);
  }
  return matches;
}

function extractTaggedSegments(text, rules) {
  const source = String(text || "");
  const occupiedRanges = [];
  return rules.map((rule) => {
    const matches = findRuleMatchesOutsideOccupied(
      source,
      rule,
      occupiedRanges,
    );
    const last = matches.length ? matches[matches.length - 1] : null;
    if (last) {
      occupiedRanges.push({
        start: last.blockStart,
        end: last.blockEnd,
        key: getRuleKey(rule),
      });
    }
    return {
      rule,
      key: getRuleKey(rule),
      content: last ? source.slice(last.contentStart, last.contentEnd) : "",
      found: Boolean(last),
      range: last ? { start: last.blockStart, end: last.blockEnd } : null,
    };
  });
}

function renderTemplate(template, values) {
  const source = String(template || "");
  return source.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_match, key) =>
    String(values[key] ?? ""),
  );
}

function compactPromptText(text) {
  return String(text || "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function getNonBodyFormatRules(settings) {
  const body = getBodyRuleInfo(settings, false);
  return getEnabledFormatRules(settings).filter(
    (rule) => getRuleKey(rule) !== body.key,
  );
}

function buildFormatRulesTextForRules(rules) {
  if (!rules.length) return "未配置启用的非正文格式标签规则。";
  return rules
    .map((rule, index) =>
      [
        `${index + 1}. ${rule.name || "未命名标签"}`,
        `规则ID：${getRuleKey(rule)}`,
        `开始标签：${getRuleStartLabel(rule)}`,
        `结束标签：${rule.endTag}`,
        `标签内提示词：${rule.prompt || "无"}`,
        `标签内容模板：\n${rule.template || `${rule.startTag}\n\n${rule.endTag}`}`,
      ].join("\n"),
    )
    .join("\n\n");
}

function buildFormatReplacementValues(sourceText, settings) {
  const rules = getNonBodyFormatRules(settings);
  const segments = extractTaggedSegments(sourceText, rules);
  return {
    sourceText,
    formatRules: buildFormatRulesTextForRules(rules),
    currentSegmentsJson: JSON.stringify(
      Object.fromEntries(segments.map((item) => [item.key, item.content])),
      null,
      2,
    ),
  };
}

function formatMessagesForStatus(messages) {
  return messages
    .map(
      (message, index) =>
        `### ${index + 1}. ${message.role}\n${message.content}`,
    )
    .join("\n\n---\n\n");
}

function buildRefineMessages(sourceText, settings, isUser) {
  const basePrompt = isUser ? settings.userPrompt : settings.prompt;
  const forbiddenInstruction = buildForbiddenInstruction(settings);
  const values = {
    basePrompt,
    forbiddenInstruction,
    textToRefine: sourceText,
    sourceText,
    messageType: isUser ? "user" : "assistant",
  };
  return [
    {
      role: "system",
      content: compactPromptText(
        renderTemplate(
          settings.refineSystemTemplate || DEFAULT_REFINE_SYSTEM_TEMPLATE,
          values,
        ),
      ),
    },
    {
      role: "user",
      content: renderTemplate(
        settings.refineUserTemplate || DEFAULT_REFINE_USER_TEMPLATE,
        values,
      ),
    },
  ];
}

function buildFormatMessages(sourceText, settings, replacementOnly = false) {
  if (replacementOnly) {
    const values = buildFormatReplacementValues(sourceText, settings);
    return [
      {
        role: "system",
        content: compactPromptText(
          renderTemplate(
            settings.formatReplacementSystemTemplate ||
              DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE,
            values,
          ),
        ),
      },
      {
        role: "user",
        content: compactPromptText(
          renderTemplate(
            settings.formatReplacementUserTemplate ||
              DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE,
            values,
          ),
        ),
      },
    ];
  }

  const values = { sourceText, formatRules: buildFormatRulesText(settings) };
  return [
    {
      role: "system",
      content: compactPromptText(
        renderTemplate(
          settings.formatFullSystemTemplate ||
            DEFAULT_FORMAT_FULL_SYSTEM_TEMPLATE,
          values,
        ),
      ),
    },
    {
      role: "user",
      content: renderTemplate(
        settings.formatFullUserTemplate || DEFAULT_FORMAT_FULL_USER_TEMPLATE,
        values,
      ),
    },
  ];
}

function findLastUnclosedTag(sourceText, rules) {
  const source = String(sourceText || "");
  const events = [];
  for (const rule of rules) {
    const start = String(rule.startTag || "");
    const end = String(rule.endTag || "");
    if (!end || isStartAnchoredRule(rule) || !start) continue;
    const regex = new RegExp(
      `${escapeRegExp(start)}|${escapeRegExp(end)}`,
      "g",
    );
    let match;
    while ((match = regex.exec(source)) !== null) {
      events.push({
        index: match.index,
        text: match[0],
        type: match[0] === start ? "start" : "end",
        rule,
      });
      if (match[0] === "") regex.lastIndex++;
    }
  }
  events.sort((a, b) => a.index - b.index);
  const stack = [];
  for (const event of events) {
    const key = getRuleKey(event.rule);
    if (event.type === "start") {
      stack.push(event);
    } else {
      const lastIndex = stack
        .map((item) => getRuleKey(item.rule))
        .lastIndexOf(key);
      if (lastIndex >= 0) stack.splice(lastIndex, 1);
    }
  }
  return stack.length ? stack[stack.length - 1] : null;
}

function buildCompletionPlan(sourceText, settings, explicitMissingRules = []) {
  const rules = getEnabledFormatRules(settings);
  const body = getBodyRuleInfo(settings, false);
  const unclosed = findLastUnclosedTag(sourceText, rules);
  const missingRules = explicitMissingRules.length
    ? explicitMissingRules
    : getMissingFormatRules(sourceText, settings);
  const isBodyUnclosed = Boolean(
    unclosed && body.key && getRuleKey(unclosed.rule) === body.key,
  );
  return {
    rules,
    body,
    unclosed,
    isBodyUnclosed,
    missingRules,
    missingNonBodyRules: missingRules.filter(
      (rule) => !body.key || getRuleKey(rule) !== body.key,
    ),
  };
}

function buildCompletionMessages(
  sourceText,
  settings,
  messageId,
  missingRules = [],
) {
  const plan = buildCompletionPlan(sourceText, settings, missingRules);
  const chain = extractChainContext(
    sourceText,
    settings.completionChainRegex || settings.completionOutlineRegex,
  );
  const previousUser = getLatestUserMessage(messageId);
  const previousUserText = previousUser.message
    ? getMessageText(previousUser.message)
    : "";
  const missingText = plan.missingRules.length
    ? buildFormatRulesTextForRules(plan.missingRules)
    : "无明确缺失标签，仅按截断位置补完。";
  const unclosedText = plan.unclosed
    ? `最后一个未闭合标签：${plan.unclosed.rule.name || getRuleKey(plan.unclosed.rule)} ${plan.unclosed.rule.startTag} ... ${plan.unclosed.rule.endTag}`
    : "未检测到未闭合标签。";

  const completionRequirement = plan.isBodyUnclosed
    ? "关键要求：当前最后一个未闭合标签是正文标签。你必须优先从截断处继续写正文内容，不能只输出正文结束标签，也不能立刻跳到下一个标签。正文续写完整后，可以继续输出正文结束标签与后续缺失标签内容。"
    : "你必须配合格式检查规则继续生成剩余部分，续写内容需要满足标签顺序、开始标签、结束标签、标签内提示词和模板要求。";

  const chainRegexText = String(
    settings.completionChainRegex || settings.completionOutlineRegex || "",
  ).trim();
  const chainMissText = chainRegexText
    ? `参考思维链：未匹配到，禁止自行编造思维链。\n调试信息：已使用正则 ${chainRegexText} 匹配当前 AI 回复文本，文本长度 ${String(sourceText || "").length}。`
    : "参考思维链：未配置捕获正则，禁止自行编造思维链。";
  const completionUserInstruction = plan.isBodyUnclosed
    ? "请从正文截断处继续写，先补完整正文内容，再输出必要的正文结束标签和后续缺失标签。"
    : "请只输出需要追加到当前回复末尾的内容。";
  const values = {
    sourceText,
    previousUserText: previousUserText || "未找到上一条用户消息",
    chainContext: chain || chainMissText,
    unclosedInfo: unclosedText,
    missingRules: missingText,
    formatRules: buildFormatRulesText(settings),
    completionPrompt: settings.completionPrompt || "",
    completionRequirement,
    completionUserInstruction,
  };

  return [
    {
      role: "system",
      content: compactPromptText(
        renderTemplate(
          settings.completionSystemTemplate ||
            DEFAULT_COMPLETION_SYSTEM_TEMPLATE,
          values,
        ),
      ),
    },
    {
      role: "user",
      content: compactPromptText(
        renderTemplate(
          settings.completionUserTemplate || DEFAULT_COMPLETION_USER_TEMPLATE,
          values,
        ),
      ),
    },
  ];
}

function getOpenAICompatibleHeaders(providerKey, apiKey) {
  const headers = {
    "Content-Type": "application/json",
    Authorization: `Bearer ${apiKey}`,
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
  const maxTokens =
    Number(options.maxTokens || settings.maxTokens) ||
    DEFAULT_SETTINGS.maxTokens;
  const temperature = Number(options.temperature ?? settings.temperature) || 0;
  const signal = options.signal;
  const onToken =
    typeof options.onToken === "function" ? options.onToken : null;
  const stream = Boolean(options.stream && onToken);

  if (!endpoint || !model) {
    throw new Error("请先配置接口地址和模型");
  }
  if (!apiKey) {
    throw new Error("请先配置 API Key");
  }

  if (provider.apiStyle === "gemini") {
    if (onToken) onToken("[Gemini 暂使用非流式请求，正在等待完整响应...]\n");
    return callGemini(
      endpoint,
      model,
      apiKey,
      messages,
      temperature,
      maxTokens,
      signal,
    );
  }

  if (provider.apiStyle === "claude") {
    if (onToken) onToken("[Claude 暂使用非流式请求，正在等待完整响应...]\n");
    return callClaude(
      endpoint,
      model,
      apiKey,
      messages,
      temperature,
      maxTokens,
      signal,
    );
  }

  return callOpenAICompatible(
    providerKey,
    endpoint,
    model,
    apiKey,
    messages,
    temperature,
    maxTokens,
    signal,
    stream,
    onToken,
  );
}

async function callOpenAICompatible(
  providerKey,
  endpoint,
  model,
  apiKey,
  messages,
  temperature,
  maxTokens,
  signal,
  stream = false,
  onToken = null,
) {
  const payload = {
    model,
    messages,
    temperature,
    max_tokens: maxTokens,
    stream,
  };

  if (providerKey === "deepseek") {
    payload.reasoning_effort = "max";
    payload.thinking = { type: "enabled" };
  }

  const response = await fetch(`${endpoint}/chat/completions`, {
    method: "POST",
    headers: getOpenAICompatibleHeaders(providerKey, apiKey),
    signal,
    body: JSON.stringify(payload),
  });

  if (stream && response.ok && response.body) {
    return readOpenAIStream(response, onToken);
  }

  const data = await safeJson(response);
  if (!response.ok) {
    throw new Error(
      data?.error?.message || data?.message || `HTTP ${response.status}`,
    );
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
        const token =
          data?.choices?.[0]?.delta?.content || data?.choices?.[0]?.text || "";
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

async function callGemini(
  endpoint,
  model,
  apiKey,
  messages,
  temperature,
  maxTokens,
  signal,
) {
  const modelName = model.startsWith("models/") ? model : `models/${model}`;
  const systemText = messages
    .filter((item) => item.role === "system")
    .map((item) => item.content)
    .join("\n\n");
  const userText = messages
    .filter((item) => item.role !== "system")
    .map((item) => item.content)
    .join("\n\n");
  const response = await fetch(
    `${endpoint}/${modelName}:generateContent?key=${encodeURIComponent(apiKey)}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        systemInstruction: systemText
          ? { parts: [{ text: systemText }] }
          : undefined,
        contents: [{ role: "user", parts: [{ text: userText }] }],
        generationConfig: {
          temperature,
          maxOutputTokens: maxTokens,
        },
      }),
    },
  );

  const data = await safeJson(response);
  if (!response.ok) {
    throw new Error(
      data?.error?.message || data?.message || `HTTP ${response.status}`,
    );
  }

  const content = data?.candidates?.[0]?.content?.parts
    ?.map((part) => part.text || "")
    .join("");
  if (typeof content !== "string") {
    throw new Error("响应格式不正确：未找到 Gemini candidates 内容");
  }

  return stripCodeFence(content);
}

async function callClaude(
  endpoint,
  model,
  apiKey,
  messages,
  temperature,
  maxTokens,
  signal,
) {
  const system = messages
    .filter((item) => item.role === "system")
    .map((item) => item.content)
    .join("\n\n");
  const user = messages
    .filter((item) => item.role !== "system")
    .map((item) => item.content)
    .join("\n\n");
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
    throw new Error(
      data?.error?.message || data?.message || `HTTP ${response.status}`,
    );
  }

  const content = data?.content?.map((part) => part.text || "").join("");
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
      .filter((model) => String(model.name || "").includes("models/"))
      .map((model) => ({
        id: String(model.name).replace(/^models\//, ""),
        name: model.displayName || String(model.name).replace(/^models\//, ""),
      }));
  }

  const models = Array.isArray(data?.data) ? data.data : [];
  return models
    .map((model) => ({
      id: model.id || model.name,
      name: model.name || model.display_name || model.id,
    }))
    .filter((model) => model.id);
}

async function fetchProviderModels(providerKey = getProviderKey()) {
  const provider = PROVIDERS[providerKey];
  const providerSettings = getProviderSettings(providerKey);
  const endpoint = normalizeEndpoint(
    providerSettings.endpoint || provider.defaultEndpoint || "",
  );
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
    response = await fetch(
      `${endpoint}/models?key=${encodeURIComponent(apiKey)}`,
    );
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
    throw new Error(
      data?.error?.message || data?.message || `HTTP ${response.status}`,
    );
  }

  return normalizeModelList(data, providerKey).sort((a, b) =>
    String(a.name).localeCompare(String(b.name)),
  );
}

function groupModels(models) {
  const grouped = {};
  for (const model of models) {
    const vendor = String(model.id || "other").includes("/")
      ? String(model.id).split("/")[0]
      : "models";
    grouped[vendor] = grouped[vendor] || [];
    grouped[vendor].push(model);
  }
  return grouped;
}

function updateModelSelect(providerKey = getProviderKey()) {
  const settings = getSettings();
  const providerSettings = getProviderSettings(providerKey);
  const models = Array.isArray(providerSettings.models)
    ? providerSettings.models
    : [];
  const $select = $("#response_refiner_model_select");
  const $manual = $("#response_refiner_model_manual");

  $select.empty();
  if (!models.length) {
    $select.append(
      $("<option>", { value: "", text: "未加载模型列表，可手动输入" }),
    );
  } else {
    const grouped = groupModels(models);
    for (const [group, items] of Object.entries(grouped)) {
      const $group = $("<optgroup>", { label: group.toUpperCase() });
      for (const model of items) {
        $group.append(
          $("<option>", { value: model.id, text: model.name || model.id }),
        );
      }
      $select.append($group);
    }
  }

  if (
    providerSettings.model &&
    models.some((model) => model.id === providerSettings.model)
  ) {
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
    if (
      models.length &&
      !models.some((model) => model.id === providerSettings.model)
    ) {
      providerSettings.model = models[0].id;
    }
    saveSettings();
    updateModelSelect(providerKey);
    toastr.success(`已获取 ${models.length} 个模型`, "模型列表");
  } catch (error) {
    toastr.error(
      `获取模型列表失败: ${error instanceof Error ? error.message : String(error)}`,
      "模型列表",
    );
    updateModelSelect(providerKey);
  } finally {
    $button.prop("disabled", false).find("i").removeClass("fa-spin");
  }
}

function updateConnectionTypeUI() {
  const providerKey = getProviderKey();
  const provider = PROVIDERS[providerKey];
  const providerSettings = getProviderSettings(providerKey);
  const endpointValue = normalizeEndpoint(providerSettings.endpoint);
  const displayEndpoint =
    endpointValue ||
    (providerKey === "direct"
      ? ""
      : normalizeEndpoint(provider.defaultEndpoint || ""));

  setText($("#response_refiner_endpoint_label"), provider.endpointLabel);
  $("#response_refiner_endpoint")
    .attr("placeholder", provider.endpointPlaceholder)
    .val(displayEndpoint);
  $("#response_refiner_api_key_input").val(providerSettings.apiKey || "");
  updateModelSelect(providerKey);
}

async function testConnection() {
  try {
    toastr.info("正在测试连接...", "测试连接");
    const text = await callAI(
      [
        { role: "system", content: "你是连接测试助手。" },
        { role: "user", content: "请只回复：连接成功" },
      ],
      { temperature: 0, maxTokens: 32 },
    );
    if (!text) {
      throw new Error("API 未返回文本");
    }
    toastr.success("连接测试成功", "测试连接");
  } catch (error) {
    toastr.error(
      `连接测试失败: ${error instanceof Error ? error.message : String(error)}`,
      "测试连接",
    );
  }
}

async function runRefine(
  originalText,
  settings,
  isUser,
  signal,
  onToken = null,
  onStatus = null,
) {
  const textToRefine = isUser
    ? originalText
    : extractTextToRefine(originalText, settings);
  const messages = buildRefineMessages(textToRefine, settings, isUser);
  onStatus?.(`完整的实际发送提示词：\n${formatMessagesForStatus(messages)}\n`);
  const refinedText = await callAI(messages, {
    signal,
    stream: settings.streamStatusEnabled,
    onToken,
  });
  const finalText = isUser
    ? refinedText
    : replaceRefinedText(originalText, refinedText, settings);
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

function normalizeReplacementContent(rule, replacement) {
  const end = String(rule.endTag || "");
  const actualStart = isStartAnchoredRule(rule)
    ? String(rule.realStartTag || "")
    : String(rule.startTag || "");
  let value = String(replacement ?? "").trim();
  if (isStartAnchoredRule(rule)) {
    value = value
      .replace(
        new RegExp(`^\\s*${escapeRegExp(START_OF_TEXT_LABEL)}\\s*`, "i"),
        "",
      )
      .trim();
  }
  if (actualStart && actualStart !== START_OF_TEXT_TAG) {
    value = value
      .replace(new RegExp(`^\\s*${escapeRegExp(actualStart)}\\s*`), "")
      .trim();
  }
  if (end) {
    value = value
      .replace(new RegExp(`\\s*${escapeRegExp(end)}\\s*$`), "")
      .trim();
  }
  return value;
}

function replaceTaggedSegmentContent(sourceText, rule, replacement) {
  const start = String(rule.startTag || "");
  const end = String(rule.endTag || "");
  const value = normalizeReplacementContent(rule, replacement);
  if (!end || !value || (!start && !isStartAnchoredRule(rule))) {
    return sourceText;
  }

  if (isStartAnchoredRule(rule)) {
    const block = `${value}${value.endsWith("\n") ? "" : "\n"}${end}`;
    const regex = new RegExp(`^([\\s\\S]*?)${escapeRegExp(end)}`);
    if (regex.test(sourceText)) {
      return sourceText.replace(regex, block);
    }
    return `${block}${sourceText.startsWith("\n") ? "" : "\n"}${sourceText}`;
  }

  const regex = new RegExp(
    `${escapeRegExp(start)}([\\s\\S]*?)${escapeRegExp(end)}`,
    "g",
  );
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

function getRuleFullBlock(text, rule) {
  const source = String(text || "");
  const start = String(rule.startTag || "");
  const end = String(rule.endTag || "");
  if (!end || (!start && !isStartAnchoredRule(rule))) return "";
  if (isStartAnchoredRule(rule)) {
    const match = source.match(new RegExp(`^([\\s\\S]*?${escapeRegExp(end)})`));
    return match ? match[1] : "";
  }
  const regex = new RegExp(
    `${escapeRegExp(start)}[\\s\\S]*?${escapeRegExp(end)}`,
    "g",
  );
  const matches = source.match(regex);
  return matches?.length ? matches[matches.length - 1] : "";
}

function removeRuleBlocks(text, rules) {
  let result = String(text || "");
  for (const rule of rules) {
    const start = String(rule.startTag || "");
    const end = String(rule.endTag || "");
    if (!end || (!start && !isStartAnchoredRule(rule))) continue;
    const regex = isStartAnchoredRule(rule)
      ? new RegExp(`^([\\s\\S]*?${escapeRegExp(end)})(?:\\s*)`)
      : new RegExp(
          `${escapeRegExp(start)}[\\s\\S]*?${escapeRegExp(end)}\\s*`,
          "g",
        );
    result = result.replace(regex, "");
  }
  return result.trim();
}

function reorderFormatBlocks(sourceText, rules) {
  const orderedBlocks = [];
  for (const rule of rules) {
    const block = getRuleFullBlock(sourceText, rule);
    if (block) orderedBlocks.push(block.trim());
  }
  const remainder = removeRuleBlocks(sourceText, rules);
  return [...orderedBlocks, remainder].filter(Boolean).join("\n");
}

function applyFormatReplacements(
  sourceText,
  settings,
  replacements,
  rules = getNonBodyFormatRules(settings),
) {
  let result = sourceText;
  for (const rule of rules) {
    const key = getRuleKey(rule);
    if (Object.prototype.hasOwnProperty.call(replacements, key)) {
      const value = normalizeReplacementContent(rule, replacements[key]);
      if (!value) {
        throw new Error(`格式检查返回的 ${key} 为空字符串，已拒绝清空标签内容`);
      }
      result = replaceTaggedSegmentContent(result, rule, value);
    }
  }
  return reorderFormatBlocks(result, getEnabledFormatRules(settings));
}

function getMissingFormatRules(sourceText, settings) {
  return getEnabledFormatRules(settings).filter((rule) => {
    const segment = extractTaggedSegments(sourceText, [rule])[0];
    return !segment?.found;
  });
}

async function runFormat(
  sourceText,
  settings,
  signal,
  replacementOnly = true,
  onToken = null,
  onStatus = null,
) {
  if (!replacementOnly) {
    const messages = buildFormatMessages(sourceText, settings);
    onStatus?.(
      `完整的实际发送提示词：\n${formatMessagesForStatus(messages)}\n`,
    );
    const formattedText = await callAI(messages, {
      signal,
      stream: settings.streamStatusEnabled,
      onToken,
    });
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
  const messages = buildFormatMessages(sourceText, settings, true);
  onStatus?.(`完整的实际发送提示词：\n${formatMessagesForStatus(messages)}\n`);
  onToken = onToken || null;
  const replacementText = await callAI(messages, {
    signal,
    stream: settings.streamStatusEnabled,
    onToken,
  });
  const replacements = parseFormatReplacementOutput(replacementText);
  const formattedText = applyFormatReplacements(
    sourceText,
    settings,
    replacements,
    rules,
  );
  return {
    stage: "format",
    original_text: sourceText,
    refined_text: JSON.stringify(replacements, null, 2),
    candidate_text: formattedText,
  };
}

function appendTemplateForMissingRules(sourceText, rules) {
  // 不再追加模板占位符。缺失标签应由模型生成实际内容；脚本只负责必要的正文闭合兜底，避免候选结果出现 [时间]/[选项内容] 等模板文本。
  return sourceText;
}

function finalizeCompletionText(sourceText, continuation, settings, plan) {
  let joined = sourceText + String(continuation || "");
  if (
    plan?.isBodyUnclosed &&
    plan.body?.endTag &&
    !extractTaggedSegments(joined, [plan.body.rule]).some((item) => item.found)
  ) {
    joined += `${joined.endsWith("\n") ? "" : "\n"}${plan.body.endTag}`;
  }
  if (plan?.isBodyUnclosed && plan.missingNonBodyRules?.length) {
    joined = appendTemplateForMissingRules(joined, plan.missingNonBodyRules);
  }
  return joined;
}

async function runCompletion(
  sourceText,
  settings,
  messageId,
  signal,
  onToken = null,
  missingRules = [],
  onStatus = null,
) {
  const plan = buildCompletionPlan(sourceText, settings, missingRules);
  onStatus?.(
    plan.isBodyUnclosed
      ? "检测到正文标签未闭合：将要求模型先从截断处续写正文，再由脚本兜底闭合正文标签。\n"
      : "未检测到正文标签未闭合：按缺失标签/截断位置补完。\n",
  );
  const messages = buildCompletionMessages(
    sourceText,
    settings,
    messageId,
    plan.missingRules,
  );
  onStatus?.(`完整的实际发送提示词：\n${formatMessagesForStatus(messages)}\n`);
  const continuation = await callAI(messages, {
    signal,
    stream: settings.streamStatusEnabled,
    onToken,
  });
  onStatus?.(
    `模型返回补完片段，长度 ${String(continuation || "").length} 字符。\n`,
  );
  const formatted = finalizeCompletionText(
    sourceText,
    continuation,
    settings,
    plan,
  );
  if (formatted !== sourceText + continuation) {
    onStatus?.("脚本已执行补完兜底：闭合未闭合的正文标签。\n");
  }
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

function getStatusKey(messageId) {
  return `${messageId}:active`;
}

function appendStatus(
  messageId,
  stage,
  text,
  reset = false,
  stateClass = "",
  options = {},
) {
  if (!getSettings().streamStatusEnabled) return;
  const key = getStatusKey(messageId);
  if (reset) state.statusBuffers.set(key, "");
  const stamp = new Date().toLocaleTimeString();
  const value = String(text || "");
  const chunk = options.timestamp === false ? value : `[${stamp}] ${value}`;
  const next = (state.statusBuffers.get(key) || "") + (value ? chunk : "");
  state.statusBuffers.set(key, next);
  const $panel = $(
    `#chat .mes[mesid="${messageId}"] .response-refiner-status-panel`,
  ).last();
  if (!$panel.length) return;
  $panel.removeClass(
    "response-refiner-status-running response-refiner-status-done response-refiner-status-error response-refiner-status-stopped",
  );
  if (stateClass) $panel.addClass(stateClass);
  $panel.find(".response-refiner-status-stage").text(stage);
  $panel.find(".response-refiner-status-text").text(next || "等待开始...");
  const node = $panel.find(".response-refiner-status-text").get(0);
  if (node) node.scrollTop = node.scrollHeight;
}

function setStatusFinal(messageId, stage, text, stateClass) {
  stopStatusHeartbeat(messageId);
  appendStatus(messageId, stage, text, false, stateClass);
}

const STATUS_HEARTBEAT_IDLE_MS = 15000;
const STATUS_HEARTBEAT_INTERVAL_MS = 15000;

function formatElapsedTime(ms) {
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return minutes ? `${minutes}分${seconds}秒` : `${seconds}秒`;
}

function stopStatusHeartbeat(messageId) {
  const key = getStatusKey(messageId);
  const heartbeat = state.statusHeartbeats.get(key);
  if (heartbeat?.timer) {
    clearInterval(heartbeat.timer);
  }
  state.statusHeartbeats.delete(key);
}

function startStatusHeartbeat(messageId, stage, startedAt, getLastActivityAt) {
  if (!getSettings().streamStatusEnabled) return null;
  stopStatusHeartbeat(messageId);
  const key = getStatusKey(messageId);
  const heartbeat = {
    timer: setInterval(() => {
      const lastActivityAt = Number(getLastActivityAt()) || startedAt;
      if (Date.now() - lastActivityAt < STATUS_HEARTBEAT_IDLE_MS) return;
      const waited = formatElapsedTime(Date.now() - startedAt);
      appendStatus(
        messageId,
        stage,
        `等待提示：已等待 ${waited}｜当前阶段：${stage}｜暂未收到新状态或新输出｜可能原因：接口排队、网络较慢、非流式接口仍在生成或模型尚未产生新 token｜可继续等待或点击停止。\n`,
      );
    }, STATUS_HEARTBEAT_INTERVAL_MS),
  };
  state.statusHeartbeats.set(key, heartbeat);
  return heartbeat;
}

function runStageWithStatus(messageId, stage, fn) {
  const startedAt = Date.now();
  let lastActivityAt = startedAt;
  const markActivity = () => {
    lastActivityAt = Date.now();
  };
  appendStatus(
    messageId,
    stage,
    `开始执行：${stage}\n`,
    true,
    "response-refiner-status-running",
  );
  appendStatus(messageId, stage, "正在构建提示词并发送请求，等待模型返回...\n");
  startStatusHeartbeat(messageId, stage, startedAt, () => lastActivityAt);
  let outputStarted = false;
  const onToken = (token) => {
    markActivity();
    if (!outputStarted) {
      outputStarted = true;
      appendStatus(messageId, stage, "收到模型输出：\n");
    }
    appendStatus(messageId, stage, token, false, "", { timestamp: false });
  };
  const onStatus = (text) => {
    markActivity();
    appendStatus(messageId, stage, text);
  };
  return fn(onToken, onStatus)
    .then((result) => {
      stopStatusHeartbeat(messageId);
      appendStatus(
        messageId,
        stage,
        `阶段完成，候选文本长度 ${String(result?.candidate_text || "").length} 字符。\n`,
      );
      return result;
    })
    .catch((error) => {
      stopStatusHeartbeat(messageId);
      throw error;
    });
}

async function runSelectedPipeline(
  fullOriginalText,
  settings,
  isUser,
  messageId,
  signal,
) {
  let workingText = fullOriginalText;
  const stageResults = [];
  if (isUser) {
    if (settings.features.refineEnabled) {
      const result = await runStageWithStatus(
        messageId,
        "功能1 润色",
        (onToken, onStatus) =>
          runRefine(workingText, settings, true, signal, onToken, onStatus),
      );
      stageResults.push(result);
      workingText = result.candidate_text;
    }
    return {
      workingText,
      stageResults,
      stages: stageResults.map((item) => item.stage),
    };
  }

  toastr.info(
    "执行已勾选功能会按需分步处理，最多消耗三次请求。",
    "Response Refiner",
  );
  appendStatus(
    messageId,
    "执行已勾选功能",
    "开始解析勾选功能与当前标签状态。\n",
    true,
    "response-refiner-status-running",
  );
  const missingRules = getMissingFormatRules(workingText, settings);
  const completionPlan = buildCompletionPlan(
    workingText,
    settings,
    missingRules,
  );
  appendStatus(
    messageId,
    "执行已勾选功能",
    `检测到缺失标签 ${missingRules.length} 个；正文未闭合：${completionPlan.isBodyUnclosed ? "是" : "否"}。\n`,
  );
  if (
    settings.features.completionEnabled &&
    (missingRules.length || completionPlan.isBodyUnclosed)
  ) {
    const result = await runStageWithStatus(
      messageId,
      "功能3 回复补完",
      (onToken, onStatus) =>
        runCompletion(
          workingText,
          settings,
          messageId,
          signal,
          onToken,
          missingRules,
          onStatus,
        ),
    );
    stageResults.push(result);
    workingText = result.candidate_text;
  } else {
    appendStatus(
      messageId,
      "执行已勾选功能",
      settings.features.completionEnabled
        ? "跳过回复补完：未检测到缺失标签或正文未闭合。\n"
        : "跳过回复补完：设置中未启用。\n",
    );
  }
  if (settings.features.formatEnabled) {
    const result = await runStageWithStatus(
      messageId,
      "功能2 格式检查和补全修正",
      (onToken, onStatus) =>
        runFormat(workingText, settings, signal, true, onToken, onStatus),
    );
    stageResults.push(result);
    workingText = result.candidate_text;
  } else {
    appendStatus(messageId, "执行已勾选功能", "跳过格式检查：设置中未启用。\n");
  }
  if (settings.features.refineEnabled) {
    const result = await runStageWithStatus(
      messageId,
      "功能1 润色",
      (onToken, onStatus) =>
        runRefine(workingText, settings, false, signal, onToken, onStatus),
    );
    stageResults.push(result);
    workingText = result.candidate_text;
  } else {
    appendStatus(messageId, "执行已勾选功能", "跳过润色：设置中未启用。\n");
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
  let stages =
    feature === "selected" ? getSelectedStages(settings, isUser) : [feature];
  stages = stages.filter((stage) => stage !== "completion" || !isUser);
  stages = stages.filter((stage) => stage !== "format" || !isUser);

  if (!stages.length) {
    toastr.warning("没有已启用的可执行功能", "Response Refiner");
    return;
  }

  const controller = new AbortController();
  state.busyMessageIds.add(resolvedId);
  state.requestControllers.set(resolvedId, controller);
  state.statusBuffers.delete(getStatusKey(resolvedId));
  stopStatusHeartbeat(resolvedId);
  if (message.extra?.response_refiner) {
    delete message.extra.response_refiner.current_candidate;
  }
  updateMessageButtons(resolvedId);
  appendStatus(
    resolvedId,
    "请求已创建",
    `准备执行：${feature === "selected" ? "执行设置中已勾选的功能" : stages.join("、")}。\n`,
    true,
    "response-refiner-status-running",
  );

  try {
    const fullOriginalText = getMessageText(message);
    let workingText = fullOriginalText;
    let stageResults = [];

    if (feature === "selected") {
      const selected = await runSelectedPipeline(
        fullOriginalText,
        settings,
        isUser,
        resolvedId,
        controller.signal,
      );
      workingText = selected.workingText;
      stageResults = selected.stageResults;
      stages = selected.stages;
    } else {
      for (const stage of stages) {
        let result;
        if (stage === "refine") {
          result = await runStageWithStatus(
            resolvedId,
            "功能1 润色",
            (onToken, onStatus) =>
              runRefine(
                workingText,
                settings,
                isUser,
                controller.signal,
                onToken,
                onStatus,
              ),
          );
        } else if (stage === "format") {
          result = await runStageWithStatus(
            resolvedId,
            "功能2 格式检查和补全修正",
            (onToken, onStatus) =>
              runFormat(
                workingText,
                settings,
                controller.signal,
                true,
                onToken,
                onStatus,
              ),
          );
        } else if (stage === "completion") {
          result = await runStageWithStatus(
            resolvedId,
            "功能3 回复补完",
            (onToken, onStatus) =>
              runCompletion(
                workingText,
                settings,
                resolvedId,
                controller.signal,
                onToken,
                getMissingFormatRules(workingText, settings),
                onStatus,
              ),
          );
        } else {
          continue;
        }
        stageResults.push(result);
        workingText = result.candidate_text;
      }
    }

    const last = stageResults[stageResults.length - 1];
    const previousRestore = message.extra?.response_refiner?.restore_snapshot;
    message.extra = message.extra || {};
    message.extra.response_refiner = {
      feature,
      stages,
      stage_results: stageResults,
      original_text: stageResults[0]?.original_text || fullOriginalText,
      refined_text: last?.refined_text || workingText,
      full_original_text: fullOriginalText,
      restore_snapshot: previousRestore?.text ? previousRestore : null,
      current_candidate: {
        feature,
        stages,
        stage_results: stageResults,
        original_text: stageResults[0]?.original_text || fullOriginalText,
        refined_text: last?.refined_text || workingText,
        full_original_text: fullOriginalText,
        candidate_text: workingText,
        applied: false,
        created_at: Date.now(),
      },
      candidate_text: workingText,
      applied: false,
    };

    appendStatus(
      resolvedId,
      "保存结果",
      "候选结果已生成，正在保存聊天数据。\n",
    );
    await saveChatConditional();
    updateMessageButtons(resolvedId);
    updateComparisonPanel(resolvedId);
    setStatusFinal(
      resolvedId,
      "处理完成",
      "处理完成，请在预览中确认候选结果。\n",
      "response-refiner-status-done",
    );
    toastr.success("处理完成，请查看预览并决定是否替换", "Response Refiner");
  } catch (error) {
    if (error?.name === "AbortError") {
      setStatusFinal(
        resolvedId,
        "已停止",
        "请求已停止。\n",
        "response-refiner-status-stopped",
      );
      toastr.warning("请求已停止", "Response Refiner");
    } else {
      setStatusFinal(
        resolvedId,
        "处理失败",
        `${String(error instanceof Error ? error.message : error)}\n`,
        "response-refiner-status-error",
      );
      toastr.error(
        String(error instanceof Error ? error.message : error),
        "Response Refiner",
      );
    }
  } finally {
    state.busyMessageIds.delete(resolvedId);
    state.requestControllers.delete(resolvedId);
    stopStatusHeartbeat(resolvedId);
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
  const candidate = refinerData?.current_candidate || refinerData;
  if (!candidate?.candidate_text) {
    toastr.warning("没有可替换的候选结果", "替换");
    return;
  }

  const restoreText =
    candidate.full_original_text ||
    refinerData.full_original_text ||
    getMessageText(message);
  message.mes = candidate.candidate_text;
  refinerData.restore_snapshot = {
    text: restoreText,
    applied_text: candidate.candidate_text,
    applied_at: Date.now(),
  };
  refinerData.applied = true;
  refinerData.candidate_text = "";
  refinerData.current_candidate = null;
  await saveChatConditional();
  updateMessageBlockCompat(resolvedId, message);
  updateMessageButtons(resolvedId);
  updateComparisonPanel(resolvedId);
  toastr.success("已替换为候选结果，可继续再次执行或恢复原文", "替换");
}

async function restoreOriginal(messageId) {
  const { message, messageId: resolvedId } = getMessageById(messageId);
  if (!message || resolvedId < 0) return;

  const refinerData = message.extra?.response_refiner;
  const restoreText =
    refinerData?.restore_snapshot?.text || refinerData?.full_original_text;
  if (!restoreText) {
    toastr.warning("没有可恢复的原文", "恢复");
    return;
  }

  message.mes = restoreText;
  refinerData.applied = false;
  refinerData.candidate_text = "";
  refinerData.current_candidate = null;
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
  const message = /** @type {RefinerChatMessage | undefined} */ (
    chat[messageId]
  );
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
  if (!state.busyMessageIds.has(messageId)) {
    $messageBlock
      .find(".response-refiner-status-panel.response-refiner-status-running")
      .removeClass("response-refiner-status-running");
  }

  const settings = getSettings();
  const isUser = isUserMessage(message);
  const isBusy = state.busyMessageIds.has(messageId);
  const refinerData = message.extra?.response_refiner;
  const currentCandidate =
    refinerData?.current_candidate ||
    (!refinerData?.applied && refinerData?.candidate_text ? refinerData : null);
  const hasCandidate = Boolean(currentCandidate?.candidate_text);
  const canRestore = Boolean(
    refinerData?.restore_snapshot?.text ||
    (refinerData?.applied && refinerData?.full_original_text),
  );
  const isApplied = Boolean(refinerData?.applied);

  if (isBusy) {
    const $stopBtn = makeActionButton(
      messageId,
      "stop",
      "fa-stop",
      "停止当前请求",
    );
    $stopBtn
      .removeClass("response-refiner-run")
      .addClass("response-refiner-stop");
    $container.append(
      makeActionButton(messageId, "busy", "fa-spinner fa-spin", "处理中", true),
    );
    $container.append($stopBtn);
    if (settings.streamStatusEnabled) {
      renderStatusPanel($container, messageId);
    }
  } else {
    $container.append(
      makeActionButton(
        messageId,
        "refine",
        "fa-wand-magic-sparkles",
        "润色（含八股文替换/移除提示词）",
      ),
    );
    if (!isUser) {
      $container.append(
        makeActionButton(messageId, "format", "fa-code", "格式检查和补全修正"),
      );
      $container.append(
        makeActionButton(messageId, "completion", "fa-forward", "回复补完"),
      );
      $container.append(
        makeActionButton(
          messageId,
          "selected",
          "fa-list-check",
          "执行设置中已勾选的功能（最多三次请求）",
        ),
      );
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

  if (canRestore && !hasCandidate) {
    const $restoreBtn = $("<div>", {
      class: "response-refiner-btn response-refiner-restore",
      title: "恢复原文",
      "data-message-id": messageId,
    }).append($("<i>", { class: "fa-solid fa-rotate-left" }));
    $container.append($restoreBtn);
  }

  if (hasCandidate && !isApplied) {
    renderInlinePreview($container, currentCandidate);
  }
}

function renderStatusPanel($container, messageId) {
  const $messageBlock = $container.closest(".mes");
  let $panel = $messageBlock.find(".response-refiner-status-panel").last();
  if (!$panel.length) {
    $panel = $("<div>", {
      class: "response-refiner-status-panel response-refiner-status-running",
    });
    const $header = $("<div>", { class: "response-refiner-status-header" });
    $header.append($("<strong>", { text: "生成状态" }));
    $header.append(
      $("<span>", { class: "response-refiner-status-stage", text: "等待开始" }),
    );
    const $text = $("<div>", {
      class: "response-refiner-status-text",
      text: "等待开始...",
    });
    $panel.append($header, $text);
    $container.after($panel);
  } else if (!$panel.prev().is($container)) {
    $panel.detach().insertAfter($container);
  }
  $messageBlock.find(".response-refiner-status-panel").not($panel).remove();
  appendStatus(messageId, "等待开始", "");
}

function renderInlinePreview($container, refinerData) {
  const $preview = $("<div>", { class: "response-refiner-preview" });
  const $label = $("<div>", { class: "response-refiner-preview-label" });
  $label.append($("<i>", { class: "fa-solid fa-chevron-right" }));
  $label.append(document.createTextNode(" 处理预览（点击展开）"));

  const $content = $("<div>", {
    class: "response-refiner-preview-content",
  }).hide();
  const $grid = $("<div>", { class: "response-refiner-preview-grid" });
  const $left = $("<div>");
  const $right = $("<div>");
  $left.append(
    $("<div>", { class: "response-refiner-preview-title", text: "原文/输入" }),
  );
  $left.append(
    $("<div>", { class: "response-refiner-preview-text" }).text(
      refinerData.full_original_text || refinerData.original_text || "",
    ),
  );
  $right.append(
    $("<div>", { class: "response-refiner-preview-title", text: "候选结果" }),
  );
  $right.append(
    $("<div>", { class: "response-refiner-preview-text" }).text(
      refinerData.candidate_text || refinerData.refined_text || "",
    ),
  );
  $grid.append($left, $right);
  $content.append($grid);
  $preview.append($label, $content);
  $container.after($preview);

  $label.on("click", function () {
    const visible = $content.is(":visible");
    $content.slideToggle(200);
    $(this)
      .find("i")
      .toggleClass("fa-chevron-right", visible)
      .toggleClass("fa-chevron-down", !visible);
    $(this)
      .contents()
      .filter(function () {
        return this.nodeType === 3;
      })
      .remove();
    $(this).append(
      document.createTextNode(
        visible ? " 处理预览（点击展开）" : " 处理预览（点击折叠）",
      ),
    );
  });
}

function updateComparisonPanel(messageId) {
  if (!state.comparisonPanelVisible) return;

  const message = /** @type {RefinerChatMessage | undefined} */ (
    chat[messageId]
  );
  const refinerData = message?.extra?.response_refiner;
  const $panel = $("#response_refiner_comparison_panel");
  const $content = $panel.find(".response-refiner-comparison-content");
  if (!$panel.length || !$content.length) return;

  $content.empty();
  const candidate =
    refinerData?.current_candidate ||
    (!refinerData?.applied && refinerData?.candidate_text ? refinerData : null);
  if (!candidate?.candidate_text) {
    $content.append(
      $("<p>", {
        text: refinerData?.restore_snapshot?.text
          ? "当前无待替换候选；可使用消息按钮恢复原文。"
          : "暂无处理结果",
      }),
    );
    return;
  }

  const $left = $("<div>", { class: "response-refiner-comparison-column" });
  $left.append(
    $("<div>", {
      class: "response-refiner-comparison-title",
      text: "原文/输入",
    }),
  );
  $left.append(
    $("<div>", { class: "response-refiner-comparison-text" }).text(
      candidate.full_original_text || candidate.original_text || "",
    ),
  );
  const $right = $("<div>", { class: "response-refiner-comparison-column" });
  $right.append(
    $("<div>", {
      class: "response-refiner-comparison-title",
      text: "候选结果",
    }),
  );
  $right.append(
    $("<div>", { class: "response-refiner-comparison-text" }).text(
      candidate.candidate_text || "",
    ),
  );
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
  const pairs = [];
  const seen = new Set();

  const tagRegex = /<([a-zA-Z][\w:-]*)\b[^>]*>([\s\S]*?)<\/\1>/g;
  let match;
  while ((match = tagRegex.exec(source)) !== null) {
    const tag = match[1];
    if (seen.has(`paired:${tag}`)) continue;
    seen.add(`paired:${tag}`);
    pairs.push({
      tag,
      startTag: `<${tag}>`,
      endTag: `</${tag}>`,
      content: match[2] || "",
    });
  }
  return pairs;
}

function buildExtractRulesMessages(sampleText, pairs) {
  const settings = getSettings();
  const prompt =
    String(
      settings.extractRulesPrompt || DEFAULT_EXTRACT_RULES_PROMPT,
    ).trim() || DEFAULT_EXTRACT_RULES_PROMPT;
  return [
    {
      role: "system",
      content:
        "你是格式规则提取助手。你只输出 JSON 数组，不输出解释、Markdown 或额外文本。",
    },
    {
      role: "user",
      content: [
        prompt,
        "已提取标签：",
        JSON.stringify(pairs, null, 2),
        "完整回复样例：",
        sampleText,
      ].join("\n"),
    },
  ];
}

function makeRuleId() {
  return crypto?.randomUUID
    ? crypto.randomUUID()
    : String(Date.now() + Math.random());
}

function extractChainPairFromSample(text, regexText) {
  const pattern = String(regexText || "").trim();
  if (!pattern) return null;
  const source = String(text || "");
  const content = extractChainContext(source, pattern);
  if (!content) return null;
  const decodedPattern = unescapeHtmlEntities(pattern);
  const endTagMatch =
    decodedPattern.match(/<\\?\/([a-zA-Z][\w:-]*)>/) ||
    decodedPattern.match(/<\/([a-zA-Z][\w:-]*)>/);
  const startTagMatch = decodedPattern.match(/<([a-zA-Z][\w:-]*)\b[^>]*>/);
  const tag = endTagMatch?.[1] || startTagMatch?.[1] || "think";
  return {
    tag,
    startTag: startTagMatch ? `<${startTagMatch[1]}>` : START_OF_TEXT_TAG,
    realStartTag: startTagMatch ? `<${startTagMatch[1]}>` : START_OF_TEXT_TAG,
    endTag: endTagMatch ? `</${endTagMatch[1]}>` : `</${tag}>`,
    content,
    startAnchored: !startTagMatch,
    forceTop: !startTagMatch,
    sourceRegex: pattern,
  };
}

function buildExtractChainRuleMessages(pair) {
  const settings = getSettings();
  const endTag = String(pair.endTag || "");
  const realStartTag = String(
    pair.realStartTag || pair.startTag || START_OF_TEXT_TAG,
  );
  const startOnly = !realStartTag || realStartTag === START_OF_TEXT_TAG;
  const startLabel = startOnly ? START_OF_TEXT_LABEL : realStartTag;
  const tag = String(
    pair.tag || endTag.replace(/^<\//, "").replace(/>$/, "") || "特殊块",
  );
  const prompt =
    String(
      settings.extractChainRulePrompt || DEFAULT_EXTRACT_CHAIN_RULE_PROMPT,
    ).trim() || DEFAULT_EXTRACT_CHAIN_RULE_PROMPT;
  return [
    {
      role: "system",
      content:
        "你是思维链格式规则分析助手。你只输出严格 JSON，不输出解释、Markdown 或额外文本。",
    },
    {
      role: "user",
      content: [
        prompt
          .replace(/\{\{startTag\}\}/g, startLabel)
          .replace(/\{\{endTag\}\}/g, endTag)
          .replace(/\{\{tag\}\}/g, tag),
        "边界信息：",
        JSON.stringify(
          {
            tag,
            startTag: startLabel,
            endTag,
            startAnchored: startOnly,
            sourceRegex: pair.sourceRegex || "",
          },
          null,
          2,
        ),
        "捕获到的思维链原文：",
        String(pair.content || ""),
        "输出 JSON 结构示例：",
        JSON.stringify(
          {
            name: `${tag} 思维链`,
            steps: [
              {
                order: 1,
                label: "步骤名称",
                boundary: "边界特征",
                structure: "标签或段落结构",
                requirement: "生成要求",
              },
            ],
            boundary_features: {
              start_tag: startLabel,
              end_tag: endTag,
              start_anchored: startOnly,
              paragraph_pattern: "",
              numbering_pattern: "",
            },
            output_rule: {
              prompt: "可复用格式规则提示词",
              template: startOnly
                ? `这里填写${tag}内容\n${endTag}`
                : `${realStartTag}\n这里填写${tag}内容\n${endTag}`,
            },
          },
          null,
          2,
        ),
      ].join("\n"),
    },
  ];
}

function parseExtractChainRuleOutput(text) {
  const cleaned = stripCodeFence(text);
  const match = cleaned.match(/\{[\s\S]*\}/);
  return JSON.parse(match ? match[0] : cleaned);
}

function createFallbackAnchoredRuleFromPair(pair, reason = "结构化分析失败") {
  const settings = getSettings();
  const endTag = String(pair.endTag || "");
  const realStartTag = String(
    pair.realStartTag || pair.startTag || START_OF_TEXT_TAG,
  );
  const tag = String(
    pair.tag || endTag.replace(/^<\//, "").replace(/>$/, "") || "特殊块",
  );
  const startOnly = !realStartTag || realStartTag === START_OF_TEXT_TAG;
  const startLabel = startOnly ? START_OF_TEXT_LABEL : realStartTag;
  const promptTemplate =
    String(
      settings.extractChainRulePrompt || DEFAULT_EXTRACT_CHAIN_RULE_PROMPT,
    ).trim() || DEFAULT_EXTRACT_CHAIN_RULE_PROMPT;
  const prompt = promptTemplate
    .replace(/\{\{startTag\}\}/g, startLabel)
    .replace(/\{\{endTag\}\}/g, endTag)
    .replace(/\{\{tag\}\}/g, tag);
  return {
    id: makeRuleId(),
    enabled: true,
    name: startOnly ? `${tag} 文本起点块` : `${tag} 思维链`,
    startTag: startOnly ? START_OF_TEXT_TAG : realStartTag,
    realStartTag: startOnly ? START_OF_TEXT_TAG : realStartTag,
    endTag,
    prompt,
    template: startOnly
      ? `这里填写${tag}内容\n${endTag}`
      : `${realStartTag}\n这里填写${tag}内容\n${endTag}`,
    startAnchored: startOnly,
    forceTop: startOnly,
    sourceRegex: pair.sourceRegex || "",
  };
}

function normalizeExtractedRules(text, fixedRules = []) {
  const cleaned = stripCodeFence(text);
  const match = cleaned.match(/\[[\s\S]*\]/);
  const parsed = JSON.parse(match ? match[0] : cleaned);
  if (!Array.isArray(parsed)) throw new Error("AI 未返回规则数组");
  const generated = parsed
    .map((item) => {
      const startAnchored = Boolean(
        item.startAnchored ||
        item.forceTop ||
        item.startTag === START_OF_TEXT_TAG ||
        item.startTag === START_OF_TEXT_LABEL,
      );
      const endTag = String(item.endTag || "");
      return {
        id: makeRuleId(),
        enabled: true,
        name: String(
          item.name || (startAnchored ? "开头特殊块" : "自动提取规则"),
        ),
        startTag: startAnchored
          ? START_OF_TEXT_TAG
          : String(item.startTag || ""),
        endTag,
        prompt: String(item.prompt || ""),
        template: String(
          item.template ||
            (startAnchored
              ? `这里填写块内容\n${endTag}`
              : `${item.startTag || ""}\n\n${endTag}`),
        ),
        startAnchored,
        forceTop: Boolean(item.forceTop || startAnchored),
      };
    })
    .filter(
      (rule) => (rule.startTag || isStartAnchoredRule(rule)) && rule.endTag,
    );

  const fixedKeys = new Set(
    fixedRules.map((rule) => `${rule.startTag}|${rule.endTag}`),
  );
  return [
    ...fixedRules,
    ...generated.filter(
      (rule) => !fixedKeys.has(`${rule.startTag}|${rule.endTag}`),
    ),
  ].sort(
    (a, b) =>
      Number(Boolean(b.forceTop || isStartAnchoredRule(b))) -
      Number(Boolean(a.forceTop || isStartAnchoredRule(a))),
  );
}

function syncExtractPromptSettingsFromDom() {
  const settings = getSettings();
  const rulesPrompt = String(
    $("#response_refiner_extract_rules_prompt").val() || "",
  );
  const chainRulePrompt = String(
    $("#response_refiner_extract_chain_rule_prompt").val() || "",
  );
  settings.extractRulesPrompt = rulesPrompt || DEFAULT_EXTRACT_RULES_PROMPT;
  settings.extractChainRulePrompt =
    chainRulePrompt || DEFAULT_EXTRACT_CHAIN_RULE_PROMPT;
  saveSettings();
}

function populateExtractPromptInputs() {
  const settings = getSettings();
  $("#response_refiner_extract_rules_prompt").val(
    settings.extractRulesPrompt || DEFAULT_EXTRACT_RULES_PROMPT,
  );
  $("#response_refiner_extract_chain_rule_prompt").val(
    settings.extractChainRulePrompt || DEFAULT_EXTRACT_CHAIN_RULE_PROMPT,
  );
}

async function runExtractRules() {
  syncExtractPromptSettingsFromDom();
  const sample = String($("#response_refiner_extract_source").val() || "");
  const chainRegex = String(
    $("#response_refiner_extract_chain_regex").val() || "",
  ).trim();
  const normalPairs = extractTagPairsFromSample(sample);
  const chainPair = extractChainPairFromSample(sample, chainRegex);
  const pairs = chainPair ? [chainPair, ...normalPairs] : normalPairs;
  if (!pairs.length) {
    toastr.warning(
      "未在样例中找到可提取标签；如需提取无开始标签的思维链，请填写思维链捕获正则",
      "自动提取格式规则",
    );
    return;
  }
  const extractMessages = normalPairs.length
    ? buildExtractRulesMessages(sample, normalPairs)
    : [];
  const $button = $("#response_refiner_extract_run");
  $button.prop("disabled", true).find("i").addClass("fa-spin");
  try {
    let chainRules = [];
    if (chainPair) {
      try {
        const chainText = await callAI(
          buildExtractChainRuleMessages(chainPair),
          {
            temperature: 0.2,
            maxTokens: Math.max(800, Number(getSettings().maxTokens) || 1200),
          },
        );
        chainRules = [
          createAnchoredRuleFromAnalysis(
            chainPair,
            parseExtractChainRuleOutput(chainText),
          ),
        ];
      } catch (chainError) {
        chainRules = [
          createFallbackAnchoredRuleFromPair(
            chainPair,
            chainError instanceof Error
              ? chainError.message
              : String(chainError),
          ),
        ];
        toastr.warning(
          "思维链结构化分析失败，已生成带标记的兜底草案，请复核后再追加",
          "自动提取格式规则",
        );
      }
    }
    if (normalPairs.length) {
      const text = await callAI(extractMessages);
      state.extractedRules = normalizeExtractedRules(text, chainRules);
    } else {
      state.extractedRules = chainRules;
    }
    $("#response_refiner_extract_output").text(
      JSON.stringify(state.extractedRules, null, 2),
    );
    $("#response_refiner_extract_apply").prop(
      "disabled",
      !state.extractedRules.length,
    );
    toastr.success(
      `已生成 ${state.extractedRules.length} 条规则建议`,
      "自动提取格式规则",
    );
  } catch (error) {
    toastr.error(
      String(error instanceof Error ? error.message : error),
      "自动提取格式规则",
    );
  } finally {
    $button.prop("disabled", false).find("i").removeClass("fa-spin");
  }
}

function applyExtractedRules() {
  if (!state.extractedRules.length) return;
  const settings = getSettings();
  syncFormatRulesFromDom();
  settings.formatRules.unshift(
    ...state.extractedRules.filter(
      (rule) => rule.forceTop || isStartAnchoredRule(rule),
    ),
  );
  settings.formatRules.push(
    ...state.extractedRules.filter(
      (rule) => !(rule.forceTop || isStartAnchoredRule(rule)),
    ),
  );
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
    ...messages.map(
      (message, index) =>
        `### ${index + 1}. ${message.role}\n${message.content}`,
    ),
  ].join("\n\n");
}

function buildPromptPreviewBlocks(sourceText, isUser) {
  const settings = getSettings();
  const source = sourceText || "【这里是待处理文本】";
  const textToRefine = isUser ? source : extractTextToRefine(source, settings);
  return {
    refine: formatPromptMessagesForPreview(
      "功能1 润色",
      buildRefineMessages(textToRefine, settings, isUser),
    ),
    format: isUser
      ? "## 功能2 格式检查和补全修正\n用户输入不执行格式检查。"
      : formatPromptMessagesForPreview(
          "功能2 格式检查和补全修正（仅非正文标签，返回替换 JSON）",
          buildFormatMessages(source, settings, true),
        ),
    completion: isUser
      ? "## 功能3 回复补完\n用户输入不执行回复补完。"
      : formatPromptMessagesForPreview(
          "功能3 回复补完",
          buildCompletionMessages(
            source,
            settings,
            chat.length,
            getMissingFormatRules(source, settings),
          ),
        ),
    selected: [
      "## 执行设置中已勾选的功能（三步组合管线）",
      "脚本先解析当前回复标签状态，然后按需执行：",
      "1. 若启用回复补完且发现缺失标签或正文标签未闭合，则调用功能3补完；否则跳过补完请求。",
      "2. 若启用格式检查，则调用功能2，只修正和补全非正文标签。",
      "3. 若启用润色，则调用功能1，只润色润色捕获正则对应的正文标签内容。",
      "4. 状态窗口会记录每个阶段的跳过原因、等待响应、收到输出、保存结果和完成/错误状态。",
      "组合执行最多会消耗三次请求，最终由脚本把每个阶段生成的候选文本继续传给下一阶段。",
    ].join("\n"),
  };
}

function fillPromptPreviewSourceFromLatestAssistant(force = false) {
  const $source = $("#response_refiner_prompt_preview_source");
  const current = String($source.val() || "");
  if (!force && current.trim()) return current;
  const latest = getLatestAssistantMessageText();
  if (latest) {
    $source.val(latest);
    return latest;
  }
  return current;
}

function updatePromptPreview() {
  const isUser =
    String($("#response_refiner_prompt_preview_type").val() || "assistant") ===
    "user";
  const source = String(
    $("#response_refiner_prompt_preview_source").val() || "",
  );
  const selected = String(
    $("#response_refiner_prompt_preview_part").val() || "refine",
  );
  const blocks = buildPromptPreviewBlocks(source, isUser);
  $("#response_refiner_prompt_preview_output").text(
    blocks[selected] || blocks.refine,
  );
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
    const ruleId =
      rule.id ||
      (rule.id = crypto?.randomUUID
        ? crypto.randomUUID()
        : String(Date.now() + Math.random()));
    const isBodyRule = body.key && getRuleKey(rule) === body.key;
    const collapsed = Boolean(settings.collapsedFormatRules?.[ruleId]);
    const $card = $("<div>", {
      class: `response-refiner-rule-card ${collapsed ? "collapsed" : ""}`,
      "data-index": index,
      "data-rule-id": ruleId,
    });
    const $header = $("<div>", { class: "response-refiner-rule-header" });
    const $title = $("<div>", { class: "response-refiner-rule-title" });
    const $toggle = $("<span>", {
      class: "response-refiner-rule-toggle",
      title: "展开/关闭规则",
    }).append(
      $("<i>", {
        class: "fa-solid fa-chevron-down response-refiner-rule-icon",
      }),
    );
    const $enabled = $("<input>", {
      type: "checkbox",
      class: "response-refiner-rule-enabled",
      title: "启用规则",
    }).prop("checked", Boolean(rule.enabled));
    const $name = $("<input>", {
      class: "text_pole response-refiner-rule-name",
      type: "text",
      title: "规则名",
    })
      .val(rule.name || `规则 ${index + 1}`)
      .prop("disabled", isBodyRule);
    $title.append($toggle, $enabled, $name);
    const $actions = $("<div>", { class: "response-refiner-rule-actions" });
    const lockPosition = isStartOnlyAnchoredRule(rule);
    $actions.append(
      $("<button>", {
        type: "button",
        class: "menu_button response-refiner-rule-up",
        text: "上移",
        disabled: lockPosition,
      }),
    );
    $actions.append(
      $("<button>", {
        type: "button",
        class: "menu_button response-refiner-rule-down",
        text: "下移",
        disabled: lockPosition,
      }),
    );
    if (!isBodyRule) {
      $actions.append(
        $("<button>", {
          type: "button",
          class: "menu_button response-refiner-rule-delete",
          text: "删除",
        }),
      );
    }
    $header.append($title, $actions);

    const $body = $("<div>", { class: "response-refiner-rule-body" }).toggle(
      !collapsed,
    );
    $body.append(
      $("<label>", {
        text: isStartOnlyAnchoredRule(rule) ? "开始位置" : "开始标签",
      }),
    );
    $body.append(
      $("<input>", {
        class: "text_pole response-refiner-rule-start",
        type: "text",
        placeholder: "例如: <content>",
      })
        .val(getRuleStartLabel(rule))
        .prop("disabled", isBodyRule || isStartOnlyAnchoredRule(rule)),
    );
    $body.append($("<label>", { text: "结束标签" }));
    $body.append(
      $("<input>", {
        class: "text_pole response-refiner-rule-end",
        type: "text",
        placeholder: "例如: </content>",
      })
        .val(rule.endTag || "")
        .prop("disabled", isBodyRule),
    );
    $body.append($("<label>", { text: "标签内提示词" }));
    $body.append(
      $("<textarea>", {
        class: "text_pole response-refiner-rule-prompt",
        rows: 3,
      }).val(rule.prompt || ""),
    );
    $body.append($("<label>", { text: "标签内容模板" }));
    $body.append(
      $("<textarea>", {
        class: "text_pole response-refiner-rule-template",
        rows: 4,
      }).val(rule.template || ""),
    );
    $card.append($header, $body);
    $list.append($card);
  });
}

function syncFormatRulesFromDom() {
  const settings = getSettings();
  const oldRules = settings.formatRules || [];
  settings.formatRules = [];
  $("#response_refiner_format_rules .response-refiner-rule-card").each(
    function () {
      const $card = $(this);
      const existingRule = oldRules[Number($card.data("index"))] || {};
      const startOnly = isStartOnlyAnchoredRule(existingRule);
      const startValue = String(
        $card.find(".response-refiner-rule-start").val() || "",
      );
      settings.formatRules.push({
        id: String(
          $card.data("rule-id") ||
            (crypto?.randomUUID
              ? crypto.randomUUID()
              : Date.now() + Math.random()),
        ),
        enabled: $card.find(".response-refiner-rule-enabled").prop("checked"),
        name: String($card.find(".response-refiner-rule-name").val() || ""),
        startTag: startOnly ? START_OF_TEXT_TAG : startValue,
        realStartTag: startOnly ? START_OF_TEXT_TAG : startValue,
        endTag: String($card.find(".response-refiner-rule-end").val() || ""),
        prompt: String($card.find(".response-refiner-rule-prompt").val() || ""),
        template: String(
          $card.find(".response-refiner-rule-template").val() || "",
        ),
        startAnchored: startOnly || Boolean(existingRule.startAnchored),
        forceTop: startOnly || Boolean(existingRule.forceTop),
        sourceRegex: String(existingRule.sourceRegex || ""),
      });
    },
  );
  settings.formatRules.sort(
    (a, b) =>
      Number(isStartOnlyAnchoredRule(b)) - Number(isStartOnlyAnchoredRule(a)),
  );
  ensureBodyRule(settings, false);
  saveSettings();
}

function updateBasicSettingsInputs() {
  const settings = getSettings();
  $("#response_refiner_connection_type").val(getProviderKey());
  updateConnectionTypeUI();
  $("#response_refiner_refine_enabled").prop(
    "checked",
    settings.features.refineEnabled,
  );
  $("#response_refiner_format_enabled").prop(
    "checked",
    settings.features.formatEnabled,
  );
  $("#response_refiner_completion_enabled").prop(
    "checked",
    settings.features.completionEnabled,
  );
  $("#response_refiner_refine_system_template").val(
    settings.refineSystemTemplate || DEFAULT_REFINE_SYSTEM_TEMPLATE,
  );
  $("#response_refiner_prompt").val(settings.prompt);
  $("#response_refiner_user_prompt").val(settings.userPrompt);
  $("#response_refiner_forbidden_phrases").val(settings.forbiddenPhrases);
  $("#response_refiner_filter_enabled").prop("checked", settings.filterEnabled);
  $("#response_refiner_filter_regex").val(settings.filterRegex);
  $("#response_refiner_format_replacement_system_template").val(
    settings.formatReplacementSystemTemplate ||
      DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE,
  );
  $("#response_refiner_completion_chain_regex").val(
    settings.completionChainRegex || settings.completionOutlineRegex || "",
  );
  $("#response_refiner_completion_prompt").val(settings.completionPrompt);
  $("#response_refiner_completion_system_template").val(
    settings.completionSystemTemplate || DEFAULT_COMPLETION_SYSTEM_TEMPLATE,
  );
  $("#response_refiner_refine_user_template").val(
    settings.refineUserTemplate || DEFAULT_REFINE_USER_TEMPLATE,
  );
  $("#response_refiner_format_replacement_user_template").val(
    settings.formatReplacementUserTemplate ||
      DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE,
  );
  $("#response_refiner_format_full_user_template").val(
    settings.formatFullUserTemplate || DEFAULT_FORMAT_FULL_USER_TEMPLATE,
  );
  $("#response_refiner_completion_user_template").val(
    settings.completionUserTemplate || DEFAULT_COMPLETION_USER_TEMPLATE,
  );
  $("#response_refiner_temperature").val(settings.temperature);
  $("#response_refiner_temperature_value").text(
    Number(settings.temperature).toFixed(2),
  );
  $("#response_refiner_max_tokens").val(settings.maxTokens);
  $("#response_refiner_stream_status_enabled").prop(
    "checked",
    settings.streamStatusEnabled,
  );
}

function resetConnectionSettings() {
  const settings = getSettings();
  settings.connectionType = DEFAULT_SETTINGS.connectionType;
  settings.providers = deepClone(DEFAULT_PROVIDER_SETTINGS);
  saveSettings();
  updateBasicSettingsInputs();
  toastr.success("已恢复连接设置默认值", "Response Refiner");
}

function resetRefineSettings() {
  const settings = getSettings();
  Object.assign(settings.features, {
    refineEnabled: DEFAULT_SETTINGS.features.refineEnabled,
  });
  Object.assign(settings, {
    prompt: DEFAULT_SETTINGS.prompt,
    userPrompt: DEFAULT_SETTINGS.userPrompt,
    forbiddenPhrases: DEFAULT_SETTINGS.forbiddenPhrases,
    filterRegex: DEFAULT_SETTINGS.filterRegex,
    filterEnabled: DEFAULT_SETTINGS.filterEnabled,
    refineSystemTemplate: DEFAULT_REFINE_SYSTEM_TEMPLATE,
  });
  ensureBodyRule(settings, false);
  saveSettings();
  updateBasicSettingsInputs();
  renderFormatRules();
  refreshAllMessageButtons();
  updatePromptPreview();
  toastr.success("已恢复功能1默认设置", "Response Refiner");
}

function resetFormatSettings() {
  const settings = getSettings();
  Object.assign(settings.features, {
    formatEnabled: DEFAULT_SETTINGS.features.formatEnabled,
  });
  Object.assign(settings, {
    formatRules: deepClone(DEFAULT_SETTINGS.formatRules),
    collapsedFormatRules: {},
    formatReplacementSystemTemplate: DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE,
    formatFullSystemTemplate: DEFAULT_FORMAT_FULL_SYSTEM_TEMPLATE,
  });
  ensureBodyRule(settings, false);
  saveSettings();
  updateBasicSettingsInputs();
  renderFormatRules();
  refreshAllMessageButtons();
  updatePromptPreview();
  toastr.success("已恢复功能2默认设置", "Response Refiner");
}

function resetCompletionSettings() {
  const settings = getSettings();
  Object.assign(settings.features, {
    completionEnabled: DEFAULT_SETTINGS.features.completionEnabled,
  });
  Object.assign(settings, {
    completionChainRegex: DEFAULT_SETTINGS.completionChainRegex,
    completionOutlineRegex: DEFAULT_SETTINGS.completionOutlineRegex,
    completionPrompt: DEFAULT_SETTINGS.completionPrompt,
    completionSystemTemplate: DEFAULT_COMPLETION_SYSTEM_TEMPLATE,
  });
  saveSettings();
  updateBasicSettingsInputs();
  refreshAllMessageButtons();
  updatePromptPreview();
  toastr.success("已恢复功能3默认设置", "Response Refiner");
}

function resetStructureTemplateSettings() {
  const settings = getSettings();
  Object.assign(settings, {
    refineUserTemplate: DEFAULT_REFINE_USER_TEMPLATE,
    formatReplacementUserTemplate: DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE,
    formatFullUserTemplate: DEFAULT_FORMAT_FULL_USER_TEMPLATE,
    completionUserTemplate: DEFAULT_COMPLETION_USER_TEMPLATE,
  });
  saveSettings();
  updateBasicSettingsInputs();
  updatePromptPreview();
  toastr.success("已恢复高级结构模板默认值", "Response Refiner");
}

function resetGenerationSettings() {
  const settings = getSettings();
  Object.assign(settings, {
    temperature: DEFAULT_SETTINGS.temperature,
    maxTokens: DEFAULT_SETTINGS.maxTokens,
    streamStatusEnabled: DEFAULT_SETTINGS.streamStatusEnabled,
  });
  saveSettings();
  updateBasicSettingsInputs();
  toastr.success("已恢复生成参数默认值", "Response Refiner");
}

function resetPromptPreviewSettings() {
  $("#response_refiner_prompt_preview_type").val("assistant");
  $("#response_refiner_prompt_preview_part").val("refine");
  $("#response_refiner_prompt_preview_source").val("");
  updatePromptPreview();
  toastr.success("已恢复提示词预览默认状态", "Response Refiner");
}

function resetExtractSettings() {
  const settings = getSettings();
  Object.assign(settings, {
    extractRulesPrompt: DEFAULT_EXTRACT_RULES_PROMPT,
    extractChainRulePrompt: DEFAULT_EXTRACT_CHAIN_RULE_PROMPT,
  });
  state.extractedRules = [];
  saveSettings();
  populateExtractPromptInputs();
  $("#response_refiner_extract_source").val("");
  $("#response_refiner_extract_chain_regex").val("");
  $("#response_refiner_extract_output").text("暂无提取结果");
  $("#response_refiner_extract_apply").prop("disabled", true);
  toastr.success("已恢复自动提取默认设置", "Response Refiner");
}

function bindSettings() {
  const settings = getSettings();
  renderProviderOptions();
  initCollapsibleSections();

  $(document).on(
    "click keydown",
    "#response_refiner_container .response-refiner-section-header",
    function (event) {
      if (
        event.type === "keydown" &&
        event.key !== "Enter" &&
        event.key !== " "
      )
        return;
      if (
        $(event.target).is(
          "input, textarea, select, button, .menu_button, option",
        )
      )
        return;
      event.preventDefault();
      toggleSettingsSection($(this).closest(".response-refiner-section"));
    },
  );

  $("#response_refiner_connection_type")
    .val(getProviderKey())
    .on("change", function () {
      settings.connectionType = String($(this).val());
      saveSettings();
      updateConnectionTypeUI();
    });

  $("#response_refiner_endpoint").on("input", function () {
    const providerKey = getProviderKey();
    const provider = PROVIDERS[providerKey];
    const value = String($(this).val());
    getProviderSettings(providerKey).endpoint =
      providerKey === "direct" ? value : normalizeEndpoint(value);
    if (providerKey !== "direct" && !normalizeEndpoint(value)) {
      $(this).val(provider.defaultEndpoint || "");
    }
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
    $(this)
      .find("i")
      .toggleClass("fa-eye", !isPassword)
      .toggleClass("fa-eye-slash", isPassword);
  });

  $("#response_refiner_test_connection").on("click", testConnection);

  $("#response_refiner_refine_enabled")
    .prop("checked", settings.features.refineEnabled)
    .on("change", function () {
      settings.features.refineEnabled = $(this).prop("checked");
      saveSettings();
      refreshAllMessageButtons();
      updatePromptPreview();
    });
  $("#response_refiner_format_enabled")
    .prop("checked", settings.features.formatEnabled)
    .on("change", function () {
      settings.features.formatEnabled = $(this).prop("checked");
      saveSettings();
      refreshAllMessageButtons();
      updatePromptPreview();
    });
  $("#response_refiner_completion_enabled")
    .prop("checked", settings.features.completionEnabled)
    .on("change", function () {
      settings.features.completionEnabled = $(this).prop("checked");
      saveSettings();
      refreshAllMessageButtons();
      updatePromptPreview();
    });

  $("#response_refiner_refine_system_template")
    .val(settings.refineSystemTemplate || DEFAULT_REFINE_SYSTEM_TEMPLATE)
    .on("input", function () {
      settings.refineSystemTemplate =
        String($(this).val()) || DEFAULT_REFINE_SYSTEM_TEMPLATE;
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_prompt")
    .val(settings.prompt)
    .on("input", function () {
      settings.prompt = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_user_prompt")
    .val(settings.userPrompt)
    .on("input", function () {
      settings.userPrompt = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_forbidden_phrases")
    .val(settings.forbiddenPhrases)
    .on("input", function () {
      settings.forbiddenPhrases = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_temperature")
    .val(settings.temperature)
    .on("input", function () {
      settings.temperature = Number($(this).val());
      $("#response_refiner_temperature_value").text(
        settings.temperature.toFixed(2),
      );
      saveSettings();
    });
  $("#response_refiner_temperature_value").text(
    Number(settings.temperature).toFixed(2),
  );
  $("#response_refiner_max_tokens")
    .val(settings.maxTokens)
    .on("input", function () {
      settings.maxTokens = Number($(this).val());
      saveSettings();
    });
  $("#response_refiner_filter_enabled")
    .prop("checked", settings.filterEnabled)
    .on("change", function () {
      settings.filterEnabled = $(this).prop("checked");
      ensureBodyRule(settings, true);
      saveSettings();
      renderFormatRules();
    });
  $("#response_refiner_filter_regex")
    .val(settings.filterRegex)
    .on("input", function () {
      settings.filterRegex = String($(this).val());
      ensureBodyRule(settings, true);
      saveSettings();
      renderFormatRules();
      updatePromptPreview();
    });
  $("#response_refiner_stream_status_enabled")
    .prop("checked", settings.streamStatusEnabled)
    .on("change", function () {
      settings.streamStatusEnabled = $(this).prop("checked");
      saveSettings();
    });

  $("#response_refiner_completion_chain_regex")
    .val(settings.completionChainRegex || settings.completionOutlineRegex || "")
    .on("input", function () {
      settings.completionChainRegex = String($(this).val());
      settings.completionOutlineRegex = settings.completionChainRegex;
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_completion_prompt")
    .val(settings.completionPrompt)
    .on("input", function () {
      settings.completionPrompt = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_completion_system_template")
    .val(
      settings.completionSystemTemplate || DEFAULT_COMPLETION_SYSTEM_TEMPLATE,
    )
    .on("input", function () {
      settings.completionSystemTemplate =
        String($(this).val()) || DEFAULT_COMPLETION_SYSTEM_TEMPLATE;
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_refine_user_template")
    .val(
      getSettingValue(
        settings.refineUserTemplate,
        DEFAULT_REFINE_USER_TEMPLATE,
      ),
    )
    .on("input", function () {
      settings.refineUserTemplate = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_format_replacement_system_template")
    .val(
      settings.formatReplacementSystemTemplate ||
        DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE,
    )
    .on("input", function () {
      settings.formatReplacementSystemTemplate =
        String($(this).val()) || DEFAULT_FORMAT_REPLACEMENT_SYSTEM_TEMPLATE;
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_format_replacement_user_template")
    .val(
      getSettingValue(
        settings.formatReplacementUserTemplate,
        DEFAULT_FORMAT_REPLACEMENT_USER_TEMPLATE,
      ),
    )
    .on("input", function () {
      settings.formatReplacementUserTemplate = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_format_full_user_template")
    .val(
      getSettingValue(
        settings.formatFullUserTemplate,
        DEFAULT_FORMAT_FULL_USER_TEMPLATE,
      ),
    )
    .on("input", function () {
      settings.formatFullUserTemplate = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_completion_user_template")
    .val(
      getSettingValue(
        settings.completionUserTemplate,
        DEFAULT_COMPLETION_USER_TEMPLATE,
      ),
    )
    .on("input", function () {
      settings.completionUserTemplate = String($(this).val());
      saveSettings();
      updatePromptPreview();
    });
  $("#response_refiner_reset_connection").on("click", resetConnectionSettings);
  $("#response_refiner_reset_refine").on("click", resetRefineSettings);
  $("#response_refiner_reset_format").on("click", resetFormatSettings);
  $("#response_refiner_reset_completion").on("click", resetCompletionSettings);
  $("#response_refiner_prompt_template_reset").on(
    "click",
    resetStructureTemplateSettings,
  );
  $("#response_refiner_reset_generation").on("click", resetGenerationSettings);
  $("#response_refiner_reset_prompt_preview").on(
    "click",
    resetPromptPreviewSettings,
  );
  $("#response_refiner_reset_extract").on("click", resetExtractSettings);
  $(
    "#response_refiner_prompt_preview_type, #response_refiner_prompt_preview_part, #response_refiner_prompt_preview_source",
  ).on("input change", updatePromptPreview);
  $("#response_refiner_refresh_prompt_preview").on("click", function () {
    fillPromptPreviewSourceFromLatestAssistant(true);
    updatePromptPreview();
  });

  populateExtractPromptInputs();
  syncStructureTemplateInputs(settings);
  $(
    "#response_refiner_extract_rules_prompt, #response_refiner_extract_chain_rule_prompt",
  ).on("input", syncExtractPromptSettingsFromDom);

  $("#response_refiner_open_extract_rules").on("click", function () {
    state.extractedRules = [];
    populateExtractPromptInputs();
    $("#response_refiner_extract_output").text("暂无提取结果");
    $("#response_refiner_extract_apply").prop("disabled", true);
    $("#response_refiner_extract_modal").show();
  });
  $(
    "#response_refiner_extract_close, #response_refiner_extract_modal .response-refiner-modal-backdrop",
  ).on("click", function () {
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

  $(document).on(
    "input change",
    "#response_refiner_format_rules input, #response_refiner_format_rules textarea",
    syncFormatRulesFromDom,
  );
  $(document).on("click", ".response-refiner-rule-delete", function () {
    const index = Number(
      $(this).closest(".response-refiner-rule-card").data("index"),
    );
    syncFormatRulesFromDom();
    settings.formatRules.splice(index, 1);
    saveSettings();
    renderFormatRules();
  });
  $(document).on("click", ".response-refiner-rule-up", function () {
    const index = Number(
      $(this).closest(".response-refiner-rule-card").data("index"),
    );
    syncFormatRulesFromDom();
    if (isStartOnlyAnchoredRule(settings.formatRules[index])) return;
    if (
      index > 0 &&
      !isStartOnlyAnchoredRule(settings.formatRules[index - 1])
    ) {
      [settings.formatRules[index - 1], settings.formatRules[index]] = [
        settings.formatRules[index],
        settings.formatRules[index - 1],
      ];
      saveSettings();
      renderFormatRules();
    }
  });
  $(document).on("click", ".response-refiner-rule-down", function () {
    const index = Number(
      $(this).closest(".response-refiner-rule-card").data("index"),
    );
    syncFormatRulesFromDom();
    if (isStartOnlyAnchoredRule(settings.formatRules[index])) return;
    if (index < settings.formatRules.length - 1) {
      [settings.formatRules[index + 1], settings.formatRules[index]] = [
        settings.formatRules[index],
        settings.formatRules[index + 1],
      ];
      settings.formatRules.sort(
        (a, b) =>
          Number(isStartOnlyAnchoredRule(b)) -
          Number(isStartOnlyAnchoredRule(a)),
      );
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
  const relativeUrl = new URL("./settings.html", import.meta.url).href;
  const urls = [
    relativeUrl,
    `/scripts/extensions/third-party/${MODULE_NAME}/settings.html`,
    `/scripts/extensions/${MODULE_NAME}/settings.html`,
  ];

  for (const url of urls) {
    try {
      const response = await fetch(url);
      if (response.ok) {
        return response.text();
      }
    } catch (_error) {}
  }

  throw new Error(
    "无法加载 settings.html，请确认插件目录内存在 settings.html，且插件文件未损坏。",
  );
}

async function addUi() {
  const settingsHtml = await loadSettingsHtml();
  $("#extensions_settings2").append(settingsHtml);

  const $panel = $("<div>", { id: "response_refiner_comparison_panel" }).hide();
  const $header = $("<div>", {
    class: "response-refiner-comparison-panel-header",
  });
  $header.append($("<h3>", { text: "Response Refiner 预览" }));
  $header.append(
    $("<button>", {
      id: "response_refiner_toggle_comparison",
      class: "menu_button",
      type: "button",
    }).append($("<i>", { class: "fa-solid fa-chevron-down" })),
  );
  $panel.append(
    $header,
    $("<div>", { class: "response-refiner-comparison-content" }).append(
      $("<p>", { text: "暂无处理结果" }),
    ),
  );
  $("#chat").before($panel);
}

function onCharacterMessageRendered(messageId) {
  if (typeof messageId !== "number" || messageId < 0) return;
  const message = /** @type {RefinerChatMessage | undefined} */ (
    chat[messageId]
  );
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

    eventSource.on(
      event_types.CHARACTER_MESSAGE_RENDERED,
      onCharacterMessageRendered,
    );
    if (event_types.MESSAGE_UPDATED) {
      eventSource.on(event_types.MESSAGE_UPDATED, onCharacterMessageRendered);
    }
    eventSource.on(event_types.CHAT_CHANGED, refreshAllMessageButtons);
    refreshAllMessageButtons();
  } catch (error) {
    state.initialized = false;
    toastr?.error?.(
      error instanceof Error ? error.message : String(error),
      "Response Refiner 初始化失败",
    );
  }
});
