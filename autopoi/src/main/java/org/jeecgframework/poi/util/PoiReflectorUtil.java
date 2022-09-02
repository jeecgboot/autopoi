package org.jeecgframework.poi.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ReflectPermission;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;


/**
 * 反射工具类,缓存读取的class信息,省的一直获取
 * @Description [LOWCOD-2521]【autopoi】大数据导出方法【全局】
 * @author liusq
 * @date  2022年1月4号
 */
public final class PoiReflectorUtil {

    private static final Map<Class<?>, PoiReflectorUtil> CACHE_REFLECTOR = new ConcurrentHashMap<Class<?>, PoiReflectorUtil>();

    private Map<String, Method> getMethods = new HashMap<String, Method>();
    private Map<String, Method> setMethods = new HashMap<String, Method>();
    private Map<String, Method> enumMethods = new HashMap<String, Method>();
    private List<Field> fieldList = new ArrayList<Field>();

    private Class<?> type;

    private PoiReflectorUtil(Class<?> clazz) {
        this.type = clazz;
        addGetMethods(clazz);
        addFields(clazz);
        addSetMethods(clazz);
    }

    public static PoiReflectorUtil forClass(Class<?> clazz) {
        return new PoiReflectorUtil(clazz);
    }

    public static PoiReflectorUtil fromCache(Class<?> clazz) {
        if (!CACHE_REFLECTOR.containsKey(clazz)) {
            CACHE_REFLECTOR.put(clazz, new PoiReflectorUtil(clazz));
        }
        return CACHE_REFLECTOR.get(clazz);
    }

    private void addGetMethods(Class<?> cls) {
        Map<String, List<Method>> conflictingGetters = new HashMap<String, List<Method>>();
        Method[] methods = getClassMethods(cls);
        for (Method method : methods) {
            String name = method.getName();
            if (name.startsWith("get") && name.length() > 3) {
                if (method.getParameterTypes().length == 0) {
                    name = methodToProperty(name);
                    addMethodConflict(conflictingGetters, name, method);
                }
            } else if (name.startsWith("is") && name.length() > 2) {
                if (method.getParameterTypes().length == 0) {
                    name = methodToProperty(name);
                    addMethodConflict(conflictingGetters, name, method);
                }
            }
        }
        resolveGetterConflicts(conflictingGetters);
    }

    private void resolveGetterConflicts(Map<String, List<Method>> conflictingGetters) {
        for (String propName : conflictingGetters.keySet()) {
            List<Method> getters = conflictingGetters.get(propName);
            Iterator<Method> iterator = getters.iterator();
            Method firstMethod = iterator.next();
            if (getters.size() == 1) {
                addGetMethod(propName, firstMethod);
            } else {
                Method getter = firstMethod;
                Class<?> getterType = firstMethod.getReturnType();
                while (iterator.hasNext()) {
                    Method method = iterator.next();
                    Class<?> methodType = method.getReturnType();
                    if (methodType.equals(getterType)) {
                        throw new RuntimeException(
                                "Illegal overloaded getter method with ambiguous type for property "
                                        + propName + " in class "
                                        + firstMethod.getDeclaringClass()
                                        + ".  This breaks the JavaBeans "
                                        + "specification and can cause unpredicatble results.");
                    } else if (methodType.isAssignableFrom(getterType)) {
                        // OK getter type is descendant
                    } else if (getterType.isAssignableFrom(methodType)) {
                        getter = method;
                        getterType = methodType;
                    } else {
                        throw new RuntimeException(
                                "Illegal overloaded getter method with ambiguous type for property "
                                        + propName + " in class "
                                        + firstMethod.getDeclaringClass()
                                        + ".  This breaks the JavaBeans "
                                        + "specification and can cause unpredicatble results.");
                    }
                }
                addGetMethod(propName, getter);
            }
        }
    }

    private void addGetMethod(String name, Method method) {
        if (isValidPropertyName(name)) {
            getMethods.put(name, method);
        }
    }

    private void addSetMethods(Class<?> cls) {
        Map<String, List<Method>> conflictingSetters = new HashMap<String, List<Method>>();
        Method[] methods = getClassMethods(cls);
        for (Method method : methods) {
            String name = method.getName();
            if (name.startsWith("set") && name.length() > 3) {
                if (method.getParameterTypes().length == 1) {
                    name = methodToProperty(name);
                    addMethodConflict(conflictingSetters, name, method);
                }
            }
        }
        resolveSetterConflicts(conflictingSetters);
    }

    private static String methodToProperty(String name) {
        if (name.startsWith("is")) {
            name = name.substring(2);
        } else if (name.startsWith("get") || name.startsWith("set")) {
            name = name.substring(3);
        } else {
            throw new RuntimeException("Error parsing property name '" + name
                    + "'.  Didn't start with 'is', 'get' or 'set'.");
        }

        if (name.length() == 1 || (name.length() > 1 && !Character.isUpperCase(name.charAt(1)))) {
            name = name.substring(0, 1).toLowerCase(Locale.ENGLISH) + name.substring(1);
        }

        return name;
    }

    private void addMethodConflict(Map<String, List<Method>> conflictingMethods, String name,
                                   Method method) {
        List<Method> list = conflictingMethods.get(name);
        if (list == null) {
            list = new ArrayList<Method>();
            conflictingMethods.put(name, list);
        }
        list.add(method);
    }

    private void resolveSetterConflicts(Map<String, List<Method>> conflictingSetters) {
        for (String propName : conflictingSetters.keySet()) {
            List<Method> setters = conflictingSetters.get(propName);
            Method firstMethod = setters.get(0);
            if (setters.size() == 1) {
                addSetMethod(propName, firstMethod);
            } else {
                Iterator<Method> methods = setters.iterator();
                Method setter = null;
                while (methods.hasNext()) {
                    Method method = methods.next();
                    if (method.getParameterTypes().length == 1) {
                        setter = method;
                        break;
                    }
                }
                if (setter == null) {
                    throw new RuntimeException(
                            "Illegal overloaded setter method with ambiguous type for property "
                                    + propName + " in class "
                                    + firstMethod.getDeclaringClass()
                                    + ".  This breaks the JavaBeans "
                                    + "specification and can cause unpredicatble results.");
                }
                addSetMethod(propName, setter);
            }
        }
    }

    private void addSetMethod(String name, Method method) {
        if (isValidPropertyName(name)) {
            setMethods.put(name, method);
        }
    }

    private void addFields(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (canAccessPrivateMethods()) {
                try {
                    field.setAccessible(true);
                } catch (Exception e) {
                    // Ignored. This is only a final precaution, nothing we can do.
                }
            }
            if (field.isAccessible() && !"serialVersionUID".equalsIgnoreCase(field.getName())) {
                this.fieldList.add(field);
            }
        }
        if (clazz.getSuperclass() != null) {
            addFields(clazz.getSuperclass());
        }
    }

    private boolean isValidPropertyName(String name) {
        return !(name.startsWith("$") || "serialVersionUID".equals(name) || "class".equals(name));
    }

    /*
     * This method returns an array containing all methods
     * declared in this class and any superclass.
     * We use this method, instead of the simpler Class.getMethods(),
     * because we want to look for private methods as well.
     *
     * @param cls The class
     * @return An array containing all methods in this class
     */
    private Method[] getClassMethods(Class<?> cls) {
        HashMap<String, Method> uniqueMethods = new HashMap<String, Method>();
        Class<?> currentClass = cls;
        while (currentClass != null) {
            addUniqueMethods(uniqueMethods, currentClass.getDeclaredMethods());

            // we also need to look for interface methods - 
            // because the class may be abstract
            Class<?>[] interfaces = currentClass.getInterfaces();
            for (Class<?> anInterface : interfaces) {
                addUniqueMethods(uniqueMethods, anInterface.getMethods());
            }

            currentClass = currentClass.getSuperclass();
        }

        Collection<Method> methods = uniqueMethods.values();

        return methods.toArray(new Method[methods.size()]);
    }

    private void addUniqueMethods(HashMap<String, Method> uniqueMethods, Method[] methods) {
        for (Method currentMethod : methods) {
            if (!currentMethod.isBridge()) {
                String signature = getSignature(currentMethod);
                // check to see if the method is already known
                // if it is known, then an extended class must have
                // overridden a method
                if (!uniqueMethods.containsKey(signature)) {
                    if (canAccessPrivateMethods()) {
                        try {
                            currentMethod.setAccessible(true);
                        } catch (Exception e) {
                            // Ignored. This is only a final precaution, nothing we can do.
                        }
                    }

                    uniqueMethods.put(signature, currentMethod);
                }
            }
        }
    }

    private String getSignature(Method method) {
        StringBuilder sb = new StringBuilder();
        Class<?> returnType = method.getReturnType();
        if (returnType != null) {
            sb.append(returnType.getName()).append('#');
        }
        sb.append(method.getName());
        Class<?>[] parameters = method.getParameterTypes();
        for (int i = 0; i < parameters.length; i++) {
            if (i == 0) {
                sb.append(':');
            } else {
                sb.append(',');
            }
            sb.append(parameters[i].getName());
        }
        return sb.toString();
    }

    private boolean canAccessPrivateMethods() {
        try {
            SecurityManager securityManager = System.getSecurityManager();
            if (null != securityManager) {
                securityManager.checkPermission(new ReflectPermission("suppressAccessChecks"));
            }
        } catch (SecurityException e) {
            return false;
        }
        return true;
    }

    public Method getGetMethod(String propertyName) {
        Method method = getMethods.get(propertyName);
        if (method == null) {
            throw new RuntimeException(
                    "There is no getter for property named '" + propertyName + "' in '" + type + "'");
        }
        return method;
    }

    public Method getSetMethod(String propertyName) {
        Method method = setMethods.get(propertyName);
        if (method == null) {
            throw new RuntimeException(
                    "There is no setter for property named '" + propertyName + "' in '" + type + "'");
        }
        return method;
    }

    /**
     * 获取field 值
     *
     * @param obj
     * @param property
     * @return
     */
    public Object getValue(Object obj, String property) {
        Object value = null;
        Method m = getMethods.get(property);
        if (m != null) {
            try {
                value = m.invoke(obj, new Object[]{});
            } catch (Exception ex) {
            }
        }
        return value;
    }

    /**
     * 设置field值
     *
     * @param obj      对象
     * @param property
     * @param object   属性值
     * @return
     */
    public boolean setValue(Object obj, String property, Object object) {
        Method m = setMethods.get(property);
        if (m != null) {
            try {
                m.invoke(obj, object);
                return true;
            } catch (Exception ex) {
                return false;
            }
        }
        return false;
    }

    public Map<String, Method> getGetMethods() {
        return getMethods;
    }

    public List<Field> getFieldList() {
        return fieldList;
    }

    public Object execEnumStaticMethod(String staticMethod, Object params) {
        if (!enumMethods.containsKey(setMethods)) {
            try {
                enumMethods.put(staticMethod,type.getMethod(staticMethod,params.getClass()));
            } catch (NoSuchMethodException e) {
                throw new RuntimeException(
                        "There is no enum for property named '" + staticMethod + "' in '" + type + "'");
            }
        }
        try {
            return enumMethods.get(staticMethod).invoke(null,params);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
