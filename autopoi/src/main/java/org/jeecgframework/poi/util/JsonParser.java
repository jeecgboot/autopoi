package org.jeecgframework.poi.util;

import java.util.*;
import java.lang.reflect.*;
import java.math.BigDecimal;
import java.math.BigInteger;

/**
 * json格式处理
 */
public class JsonParser {

    /**
     * 将JSON数组字符串解析为List<Object>
     * @param json JSON数组字符串
     * @return List集合，包含解析后的对象
     */
    public static List<Object> parseJsonArrayToList(String json) {
        if (json == null || json.trim().isEmpty()) {
            return new ArrayList<>();
        }

        String trimmed = json.trim();
        if (!trimmed.startsWith("[") || !trimmed.endsWith("]")) {
            throw new IllegalArgumentException("Input is not a valid JSON array");
        }

        // 移除外层的方括号
        String content = trimmed.substring(1, trimmed.length() - 1).trim();
        if (content.isEmpty()) {
            return new ArrayList<>();
        }

        List<Object> result = new ArrayList<>();
        List<String> elements = splitJsonElements(content);

        for (String element : elements) {
            result.add(parseJsonValue(element.trim()));
        }

        return result;
    }

    /**
     * 分割JSON数组中的元素
     */
    private static List<String> splitJsonElements(String content) {
        List<String> elements = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        int braceDepth = 0;
        int bracketDepth = 0;
        boolean inString = false;
        char stringChar = '"';
        boolean escapeNext = false;

        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);

            if (escapeNext) {
                current.append(c);
                escapeNext = false;
                continue;
            }

            if (c == '\\') {
                escapeNext = true;
                current.append(c);
                continue;
            }

            if (!inString) {
                if (c == ',') {
                    if (braceDepth == 0 && bracketDepth == 0) {
                        elements.add(current.toString());
                        current = new StringBuilder();
                        continue;
                    }
                } else if (c == '{') {
                    braceDepth++;
                } else if (c == '}') {
                    braceDepth--;
                } else if (c == '[') {
                    bracketDepth++;
                } else if (c == ']') {
                    bracketDepth--;
                } else if (c == '"' || c == '\'') {
                    inString = true;
                    stringChar = c;
                }
            } else {
                if (c == stringChar) {
                    inString = false;
                }
            }

            current.append(c);
        }

        if (current.length() > 0) {
            elements.add(current.toString());
        }

        return elements;
    }

    /**
     * 解析JSON值
     */
    private static Object parseJsonValue(String jsonValue) {
        if (jsonValue == null || jsonValue.isEmpty()) {
            return null;
        }

        String trimmed = jsonValue.trim();

        // 处理null
        if ("null".equalsIgnoreCase(trimmed)) {
            return null;
        }

        // 处理布尔值
        if ("true".equalsIgnoreCase(trimmed)) {
            return true;
        }
        if ("false".equalsIgnoreCase(trimmed)) {
            return false;
        }

        // 处理字符串
        if (isQuotedString(trimmed)) {
            return parseString(trimmed);
        }

        // 处理数字
        if (isNumber(trimmed)) {
            return parseNumber(trimmed);
        }

        // 处理对象
        if (trimmed.startsWith("{")) {
            return parseJsonObject(trimmed);
        }

        // 处理数组
        if (trimmed.startsWith("[")) {
            return parseJsonArray(trimmed);
        }

        // 如果都不是，尝试作为数字处理，否则返回字符串
        try {
            return parseNumber(trimmed);
        } catch (NumberFormatException e) {
            return trimmed;
        }
    }

    /**
     * 解析JSON对象
     */
    private static Map<String, Object> parseJsonObject(String json) {
        if (!json.startsWith("{") || !json.endsWith("}")) {
            throw new IllegalArgumentException("Invalid JSON object: " + json);
        }

        String content = json.substring(1, json.length() - 1).trim();
        if (content.isEmpty()) {
            return new LinkedHashMap<>();
        }

        Map<String, Object> result = new LinkedHashMap<>();
        List<String> pairs = splitJsonPairs(content);

        for (String pair : pairs) {
            String[] keyValue = splitKeyValue(pair.trim());
            if (keyValue.length == 2) {
                String key = parseString(keyValue[0].trim());
                Object value = parseJsonValue(keyValue[1].trim());
                result.put(key, value);
            }
        }

        return result;
    }

    /**
     * 解析JSON数组
     */
    private static List<Object> parseJsonArray(String json) {
        if (!json.startsWith("[") || !json.endsWith("]")) {
            throw new IllegalArgumentException("Invalid JSON array: " + json);
        }

        String content = json.substring(1, json.length() - 1).trim();
        if (content.isEmpty()) {
            return new ArrayList<>();
        }

        List<Object> result = new ArrayList<>();
        List<String> elements = splitJsonElements(content);

        for (String element : elements) {
            result.add(parseJsonValue(element.trim()));
        }

        return result;
    }

    /**
     * 分割JSON对象中的键值对
     */
    private static List<String> splitJsonPairs(String content) {
        List<String> pairs = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        int braceDepth = 0;
        int bracketDepth = 0;
        boolean inString = false;
        char stringChar = '"';
        boolean escapeNext = false;

        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);

            if (escapeNext) {
                current.append(c);
                escapeNext = false;
                continue;
            }

            if (c == '\\') {
                escapeNext = true;
                current.append(c);
                continue;
            }

            if (!inString) {
                if (c == ',') {
                    if (braceDepth == 0 && bracketDepth == 0) {
                        pairs.add(current.toString());
                        current = new StringBuilder();
                        continue;
                    }
                } else if (c == '{') {
                    braceDepth++;
                } else if (c == '}') {
                    braceDepth--;
                } else if (c == '[') {
                    bracketDepth++;
                } else if (c == ']') {
                    bracketDepth--;
                } else if (c == '"' || c == '\'') {
                    inString = true;
                    stringChar = c;
                }
            } else {
                if (c == stringChar) {
                    inString = false;
                }
            }

            current.append(c);
        }

        if (current.length() > 0) {
            pairs.add(current.toString());
        }

        return pairs;
    }

    /**
     * 分割键值对
     */
    private static String[] splitKeyValue(String pair) {
        int colonIndex = -1;
        boolean inString = false;
        char stringChar = '"';
        boolean escapeNext = false;

        for (int i = 0; i < pair.length(); i++) {
            char c = pair.charAt(i);

            if (escapeNext) {
                escapeNext = false;
                continue;
            }

            if (c == '\\') {
                escapeNext = true;
                continue;
            }

            if (!inString) {
                if (c == ':') {
                    colonIndex = i;
                    break;
                } else if (c == '"' || c == '\'') {
                    inString = true;
                    stringChar = c;
                }
            } else {
                if (c == stringChar) {
                    inString = false;
                }
            }
        }

        if (colonIndex == -1) {
            throw new IllegalArgumentException("Invalid key-value pair: " + pair);
        }

        return new String[] {
                pair.substring(0, colonIndex),
                pair.substring(colonIndex + 1)
        };
    }

    /**
     * 解析字符串，处理转义字符
     */
    private static String parseString(String str) {
        if (!isQuotedString(str)) {
            return str;
        }

        char quoteChar = str.charAt(0);
        String content = str.substring(1, str.length() - 1);
        StringBuilder result = new StringBuilder();

        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);

            if (c == '\\' && i + 1 < content.length()) {
                char next = content.charAt(i + 1);
                switch (next) {
                    case '"': result.append('"'); i++; break;
                    case '\'': result.append('\''); i++; break;
                    case '\\': result.append('\\'); i++; break;
                    case '/': result.append('/'); i++; break;
                    case 'b': result.append('\b'); i++; break;
                    case 'f': result.append('\f'); i++; break;
                    case 'n': result.append('\n'); i++; break;
                    case 'r': result.append('\r'); i++; break;
                    case 't': result.append('\t'); i++; break;
                    case 'u': // Unicode转义
                        if (i + 5 < content.length()) {
                            String hex = content.substring(i + 2, i + 6);
                            try {
                                int codePoint = Integer.parseInt(hex, 16);
                                result.append((char) codePoint);
                                i += 5;
                            } catch (NumberFormatException e) {
                                result.append(c);
                            }
                        } else {
                            result.append(c);
                        }
                        break;
                    default:
                        result.append(c);
                }
            } else {
                result.append(c);
            }
        }

        return result.toString();
    }

    /**
     * 解析数字
     */
    private static Object parseNumber(String numStr) {
        try {
            // 尝试解析为整数
            if (numStr.indexOf('.') == -1 && numStr.indexOf('e') == -1 && numStr.indexOf('E') == -1) {
                try {
                    long longValue = Long.parseLong(numStr);
                    if (longValue >= Integer.MIN_VALUE && longValue <= Integer.MAX_VALUE) {
                        return (int) longValue;
                    }
                    return longValue;
                } catch (NumberFormatException e) {
                    // 可能是大整数
                    return new BigInteger(numStr);
                }
            } else {
                // 尝试解析为浮点数
                double doubleValue = Double.parseDouble(numStr);
                if (Math.abs(doubleValue) <= Float.MAX_VALUE) {
                    float floatValue = (float) doubleValue;
                    // 检查是否实际上是整数
                    if (floatValue == Math.floor(floatValue) && !Double.isInfinite(floatValue)) {
                        if (floatValue >= Integer.MIN_VALUE && floatValue <= Integer.MAX_VALUE) {
                            return (int) floatValue;
                        } else {
                            return (long) floatValue;
                        }
                    }
                    return floatValue;
                }
                return doubleValue;
            }
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("Invalid number: " + numStr, e);
        }
    }

    /**
     * 判断是否为引号包裹的字符串
     */
    private static boolean isQuotedString(String str) {
        return (str.startsWith("\"") && str.endsWith("\"")) ||
                (str.startsWith("'") && str.endsWith("'"));
    }

    /**
     * 判断是否为数字
     */
    private static boolean isNumber(String str) {
        if (str == null || str.isEmpty()) {
            return false;
        }

        String trimmed = str.trim();
        if (trimmed.isEmpty()) {
            return false;
        }

        char firstChar = trimmed.charAt(0);
        if (firstChar != '-' && firstChar != '+' && !Character.isDigit(firstChar)) {
            return false;
        }

        boolean hasDecimalPoint = false;
        boolean hasExponent = false;
        boolean hasDigit = false;

        for (int i = (firstChar == '-' || firstChar == '+') ? 1 : 0; i < trimmed.length(); i++) {
            char c = trimmed.charAt(i);

            if (c >= '0' && c <= '9') {
                hasDigit = true;
            } else if (c == '.') {
                if (hasDecimalPoint || hasExponent) {
                    return false;
                }
                hasDecimalPoint = true;
            } else if (c == 'e' || c == 'E') {
                if (hasExponent || !hasDigit) {
                    return false;
                }
                hasExponent = true;
                hasDigit = false; // 指数后面必须有数字
            } else if (c == '+' || c == '-') {
                // 只允许在指数符号后出现
                char prev = (i > 0) ? trimmed.charAt(i - 1) : '\0';
                if (!(prev == 'e' || prev == 'E')) {
                    return false;
                }
            } else {
                return false;
            }
        }

        return hasDigit;
    }

    // 测试方法
    public static void main(String[] args) {
        // 测试用例1：基本数组
        String json1 = "[{\"name\":\"John\",\"age\":30}, {\"name\":\"Jane\",\"age\":25}]";
        List<Object> result1 = parseJsonArrayToList(json1);
        System.out.println("Test 1 Result: " + result1);

        // 测试用例2：嵌套结构
        String json2 = "[{\"name\":\"John\",\"scores\":[85,90,78],\"address\":{\"city\":\"NY\",\"zip\":10001}}]";
        List<Object> result2 = parseJsonArrayToList(json2);
        System.out.println("Test 2 Result: " + result2);

        // 测试用例3：各种数据类型
        String json3 = "[123, 45.67, \"hello\", true, false, null]";
        List<Object> result3 = parseJsonArrayToList(json3);
        System.out.println("Test 3 Result: " + result3);

        // 测试用例4：转义字符
        String json4 = "[\"Line1\\nLine2\", \"Quote:\\\"Hello\\\"\", \"Backslash:\\\\\"]";
        List<Object> result4 = parseJsonArrayToList(json4);
        System.out.println("Test 4 Result: " + result4);
    }
}
