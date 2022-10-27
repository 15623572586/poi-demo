package com.ontoweb.pois.xlsx;

import com.ontoweb.pois.utils.StringUtils;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class GenEntityList {

    public static List<?> reflectionDemo() throws ClassNotFoundException, InstantiationException, IllegalAccessException {

        HashMap<String, Integer> fieldsMap = new HashMap<String, Integer>(){
            {
                put("id", 0);
                put("username", 1);
                put("birthDate", 2);
                put("age", 3);
            }
        };

        List<List<String>> data = new ArrayList<List<String>>(){
            {
                add(new ArrayList<String>(){
                    {
                        add("001");
                        add("yuyunhu");
                        add("980225");
                    }
                });
                add(new ArrayList<String>(){
                    {
                        add("002");
                        add("zhangjinlong");
                        add("000000");
                        add("22");
                    }
                });
            }
        };
        List<Object> entities = new ArrayList<>();
        for (List<String> list: data) {
            Class<?> clazz = Class.forName("com.ontoweb.pois.xlsx.Entity");
            Object entity = clazz.newInstance();
            for(String key : fieldsMap.keySet()) {
                Field field = null;  // 获取私有属性字段
                try {
                    field = clazz.getDeclaredField(key);
                    field.setAccessible(true);  // 设置私有属性修改权限
                    Integer col = fieldsMap.get(key);  // 获取列
                    if(col > list.size() - 1) continue;  //  不在列范围内
                    String value = list.get(col);  // 获取col列的值
                    if (StringUtils.isNotEmpty(value)) {  // 如果数据不为空设置字段值
                        try {
                            field.set(entity, value);
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        }
                    }
                } catch (NoSuchFieldException e) {
                    e.printStackTrace();
                }

            }
            entities.add(entity);
        }
        return entities;
    }

    public static List<String> getFieldsDemo() throws ClassNotFoundException {
        Class<?> entityClass = Class.forName("com.ontoweb.pois.xlsx.Entity");
        Field[] fields = entityClass.getDeclaredFields();  // 私有属性
        List<String> fs = new ArrayList<>();
        for(Field field: fields){
            fs.add(field.getName());
        }
        return fs;
    }
    public static void main(String[] args) throws ClassNotFoundException, InstantiationException, IllegalAccessException {
        System.out.println(reflectionDemo());
        System.out.println(getFieldsDemo());
    }

}

@Data
class Entity {
    private String id;
    private String username;
    private String birthDate;
    private String age;
}