import { logWarn } from "./log";

let cachedDefaultLangId: number | null = null

export function getDefaultLanguageId(): number {
    if (cachedDefaultLangId !== null) {
        return cachedDefaultLangId
    }

    cachedDefaultLangId = Application.Language
    return cachedDefaultLangId
}

/**
 * 国际化钩子函数，用于根据路径获取本地化文本
 * @param data - 本地化数据对象，包含不同语言的文本数据，文本中可以包含`{param}`格式的占位符
 * @returns 返回一个函数，该函数可以根据路径、参数和语言ID获取对应的本地化字符串
 */
export const useI10n = (data: LocaleData) => {
    const replaceParams = (template: string, params?: LocaleParams): string => {
        if (!params) {
            return template;
        }
        let result = template;
        for (const [paramKey, paramValue] of Object.entries(params)) {
            result = result.replaceAll(`{${paramKey}}`, String(paramValue));
        }
        return result;
    };

    /**
     * 根据路径获取本地化文本
     * @param path - 本地化文本的路径，使用点号分隔的键名
     * @param params - 可选的参数对象，用于替换模板中的占位符
     * @param langId - 可选的语言ID，如果不提供则使用默认语言
     * @returns 返回本地化后的字符串
     */
    return (path: string, params?: LocaleParams, langId?: number): string => {
        const currentLangId = langId ?? getDefaultLanguageId();
        // 优先匹配当前语言，回退到英语
        const root = data[currentLangId] ?? data[wps.Enum.msoLanguageIDEnglishUS];
        
        if (!root) {
            logWarn("I18n", `No locale data found for language ID: ${currentLangId}.`);
            return path;
        }
        
        const keys = path.split(".");
        let current: LocaleStructure | string = root;

        for (const key of keys) {
            if (typeof current === "object" && key in current) {
                current = (current as LocaleStructure)[key];
            }
            else {
                logWarn("I18n", `Key '${key}' in path '${path}' not found in locale data.`);
                return path;
            }
        }
        
        if (typeof current === "string") {
            return replaceParams(current, params);
        }

        return path;
    };
};

