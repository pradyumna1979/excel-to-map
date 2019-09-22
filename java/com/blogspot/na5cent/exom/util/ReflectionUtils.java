/*
 * code https://github.com/jittagornp/excel-object-mapping
 */
package com.blogspot.na5cent.exom.util;

import com.blogspot.na5cent.exom.annotation.Column;
import com.blogspot.na5cent.exom.converter.TypeConverter;
import com.blogspot.na5cent.exom.converter.TypeConverters;
import static com.blogspot.na5cent.exom.util.CollectionUtils.isEmpty;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author redcrow
 */
public class ReflectionUtils {
    
    private static final Logger LOG = LoggerFactory.getLogger(ReflectionUtils.class);

    private static String toUpperCaseFirstCharacter(String str) {
        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }

    public static void setValueOnField(Object instance, Field field, Object value) throws Exception {
        Class clazz = instance.getClass();
        String setMethodName = "set" + toUpperCaseFirstCharacter(field.getName());

        for (Map.Entry<Class, TypeConverter> entry : TypeConverters.getConverter().entrySet()) {
            if (field.getType().equals(entry.getKey())) {
                Method method = clazz.getDeclaredMethod(setMethodName, entry.getKey());
                Column column = field.getAnnotation(Column.class);
                        
                method.invoke(
                        instance,
                        entry.getValue().convert(
                                value,
                                column == null ? null : column.pattern()
                        )
                );
            }
        }
    }
    public static void setValueMap(Map<String,String> instance, String field, String value) throws Exception {
      instance.put(field, value);
        
    }

    public static void eachFields(List<String> list, EachFieldCallback callback) throws Throwable {
        
        if (!isEmpty(list)) {
            for (String field : list) {
                callback.each(
                   field
                );
            }
        }
    }
}
