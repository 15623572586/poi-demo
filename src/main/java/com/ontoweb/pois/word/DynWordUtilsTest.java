package com.ontoweb.pois.word;


import java.util.*;

/**
 * Create by IntelliJ Idea 2018.2
 *
 * @author: qyp
 * Date: 2019-10-26 17:34
 */
public class DynWordUtilsTest {

    /**
     * 说明 普通占位符位${field}格式
     * 表格中的占位符为${tbAddRow:tb1}  tb1为唯一标识符
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) {
//        testImage();
        testT();
    }

    public static void testT() {

        // 模板全的路径
        String templatePaht = "D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\merge-test.docx";
//        String templatePaht = "D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\审查报告模板1023体检表.docx";
        // 输出位置
        String outPath = "D:\\projects\\java-demo\\data-process\\src\\out\\武汉市居民健康素养监测报告.docx";

        Map<String, Object> paramMap = new HashMap<>(16);
        // 普通的占位符示例 参数数据结构 {str,str}
//        paramMap.put("title", "德玛西亚");
//        paramMap.put("startYear", "2010");
//        paramMap.put("endYear", "2020");
//        paramMap.put("currentYear", "2019");
//        paramMap.put("currentMonth", "10");
//        paramMap.put("currentDate", "26");
//        paramMap.put("name", "黑色玫瑰");

        // 段落中的动态段示例 [str], 支持动态行中添加图片
//        List<Object> list1 = new ArrayList<Object>(Arrays.asList("2、list1_11111", "3、list1_2222", "${image:image0}"));
//        ImageEntity imgEntity = new ImageEntity();
//        imgEntity.setHeight(200);
//        imgEntity.setWidth(300);
//        imgEntity.setUrl("D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\image1.jpg");
//        imgEntity.setTypeId(ImageUtils.ImageType.JPG);
//
//        paramMap.put("image:image0", imgEntity);
//        paramMap.put("list1", list1);
//
//        List<String> list2 = new ArrayList<>(Arrays.asList("2、list2_11111", "3、list2_2222"));
//        paramMap.put("list2", list2);

        // 表格中的参数示例 参数数据结构 [[str]]
        List<List<String>> tbRow1 = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            List<String> row = null;
            if (i<8) {
                row = new ArrayList<>(Arrays.asList("科学健康观", "母乳喂养对婴儿的好处"+i, ""+i, ""+i+i));
            }else {
                row = new ArrayList<>(Arrays.asList("传染病防治", "咳嗽、打喷嚏时，正确的处理方法"+i, ""+i, ""+i+i));
            }
            tbRow1.add(row);
        }
        paramMap.put("tbAddRow:analysis", tbRow1);
//        paramMap.put(PoiWordUtils.addRowText + "result.section1.questionnaire.recovery.list", tbRow1);

//        List<List<String>> tbRow2 = new ArrayList<>();
//        List<String> tbRow2_row1 = new ArrayList<>(Arrays.asList("指标c", "指标c的意见"));
//        List<String> tbRow2_row2 = new ArrayList<>(Arrays.asList("指标d", "指标d的意见"));
//        tbRow2.add(tbRow2_row1);
//        tbRow2.add(tbRow2_row2);
//        paramMap.put(PoiWordUtils.addRowText + "tb2", tbRow2);
//
//        List<List<String>> tbRow3 = new ArrayList<>();
//        List<String> tbRow3_row1 = new ArrayList<>(Arrays.asList("3", "耕地估值"));
//        List<String> tbRow3_row2 = new ArrayList<>(Arrays.asList("4", "耕地归属", "平方公里"));
//        tbRow3.add(tbRow3_row1);
//        tbRow3.add(tbRow3_row2);
//        paramMap.put(PoiWordUtils.addRowText + "tb3", tbRow3);
//
//        // 支持在表格中动态添加图片
//        List<List<String>> tbRow4 = new ArrayList<>();
//        List<String> tbRow4_row1 = new ArrayList<>(Arrays.asList("03", "旅游用地", "18.8m2"));
//        List<String> tbRow4_row2 = new ArrayList<>(Arrays.asList("04", "建筑用地"));
//        List<String> tbRow4_row3 = new ArrayList<>(Arrays.asList("04", "${image:image3}"));
//        tbRow4.add(tbRow4_row3);
//        tbRow4.add(tbRow4_row1);
//        tbRow4.add(tbRow4_row2);
//
//        ImageEntity imgEntity3 = new ImageEntity();
//        imgEntity3.setHeight(100);
//        imgEntity3.setWidth(100);
//        imgEntity3.setUrl("D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\image1.jpg");
//        imgEntity3.setTypeId(ImageUtils.ImageType.JPG);
//
//        paramMap.put(PoiWordUtils.addRowText + "tb4", tbRow4);
//        paramMap.put("image:image3", imgEntity3);
//
//        // 图片占位符示例 ${image:imageid} 比如 ${image:image1}, ImageEntity中的值就为image:image1
//        // 段落中的图片
//        ImageEntity imgEntity1 = new ImageEntity();
//        imgEntity1.setHeight(500);
//        imgEntity1.setWidth(400);
//        imgEntity1.setUrl("D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\image1.jpg");
//        imgEntity1.setTypeId(ImageUtils.ImageType.JPG);
//        paramMap.put("image:image1", imgEntity1);
//
//        // 表格中的图片
//        ImageEntity imgEntity2 = new ImageEntity();
//        imgEntity2.setHeight(200);
//        imgEntity2.setWidth(100);
//        imgEntity2.setUrl("D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\image1.jpg");
//        imgEntity2.setTypeId(ImageUtils.ImageType.JPG);
//        paramMap.put("image:image2", imgEntity2);

        DynWordUtils.process(paramMap, templatePaht, outPath);
    }

    public static void testImage() {

        Map<String, Object> paramMap = new HashMap<>(16);
        String templatePath = "D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\11.docx";
        String outPath = "D:\\projects\\java-demo\\data-process\\src\\out\\3.docx";
        ImageEntity imgEntity1 = new ImageEntity();
//        imgEntity1.setHeight(500);
//        imgEntity1.setWidth(400);
        imgEntity1.setUrl("D:\\projects\\java-demo\\data-process\\src\\main\\java\\com\\ontoweb\\pois\\word\\wordtemplate\\image1.jpg");
        imgEntity1.setTypeId(ImageUtils.ImageType.JPG);

        paramMap.put("image:img1", imgEntity1);
        DynWordUtils.process(paramMap, templatePath, outPath);
    }
}
