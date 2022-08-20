package com.ontoweb.pois.utils;

public class StringUtils {
    public static Boolean isEmpty(String str) {
        return str == null || str.equals("");
    }

    public static Boolean isNotEmpty(String str) {
        return !isEmpty(str);
    }

}
