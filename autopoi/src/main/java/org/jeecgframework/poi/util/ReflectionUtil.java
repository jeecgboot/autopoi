package org.jeecgframework.poi.util;

import java.lang.reflect.Field;

public class ReflectionUtil {

    /**
     * 获取指定对象的字段值
     *
     * @param obj    目标对象
     * @param fieldName 字段名
     * @return 字段值
     */
    public static Object getFieldValue(Object obj, String fieldName) {
        if (obj == null || fieldName == null || fieldName.isEmpty()) {
            return null;
        }

        try {
            // 获取对象的类
            Class<?> clazz = obj.getClass();
            // 获取字段
            Field field = clazz.getDeclaredField(fieldName);
            // 如果字段是私有的，需要设置可访问
            field.setAccessible(true);
            // 返回字段的值
            return field.get(obj);
        } catch (NoSuchFieldException e) {
            System.err.println("No such field: " + fieldName);
        } catch (IllegalAccessException e) {
            System.err.println("Field is not accessible: " + fieldName);
        }
        return null;
    }

//    public static void main(String[] args) {
//        // 示例类
//        class Example {
//            private String name = "Hutool";
//            private int age = 10;
//        }
//
//        Example example = new Example();
//        System.out.println(getFieldValue(example, "name")); // 输出: Hutool
//        System.out.println(getFieldValue(example, "age"));  // 输出: 10
//    }
}
