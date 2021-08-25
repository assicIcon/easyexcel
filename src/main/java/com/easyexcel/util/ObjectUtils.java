package com.easyexcel.util;

import java.util.Collection;

/**
 * @author wrh
 * @date 2021/8/25
 */
public class ObjectUtils {

    private ObjectUtils() {}

    public static boolean isEmpty(Object obj) {
        if (obj == null) {
            return true;
        }
        if (obj instanceof Collection) {
            return ((Collection) obj).isEmpty();
        }
        if (obj instanceof String) {
            return ((String) obj).isEmpty();
        }
        return false;
    }

}
