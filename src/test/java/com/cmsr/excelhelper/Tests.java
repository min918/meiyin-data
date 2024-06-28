package com.cmsr.excelhelper;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;
import com.google.common.base.Strings;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

@Slf4j
public class Tests {

    @Test
    void test() throws IOException {

        String fileName = "C:\\test\\项目册印刷版0611.xlsx";
        // 这里 也可以不指定class，返回一个list，然后读取第一个sheet 同步读取会自动finish
        List<Map<Integer, String>> listMap = EasyExcel.read(fileName).sheet().doReadSync();
        JSONArray jsonArray = new JSONArray();
        for (Map<Integer, String> data : listMap) {
            // 返回每条数据的键值对 表示所在的列 和所在列的值
            log.info("读取到数据:{}", JSON.toJSONString(data));
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("value", data.get(0));
            jsonObject.put("fun1", data.get(1));
            jsonObject.put("time1", data.get(2));
            jsonObject.put("req1", data.get(3));
            jsonObject.put("factor1", data.get(4));
            jsonObject.put("sign1", data.get(5));
            jsonArray.add(jsonObject);
        }

        Files.write(Paths.get("C:\\test\\data1.json"), jsonArray.toJSONString().getBytes());

    }

    @Test
    void test2() throws IOException {

        String fileName = "C:\\test\\项目6.24_11.xlsx";
        // 写法2：
        // 匿名内部类 不用额外写一个DemoDataListener
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, new AnalysisEventListener<Map<String, String>>() {
            /**
             * 单次缓存的数据量
             */
            public static final int BATCH_COUNT = 1000;
            /**
             *临时存储
             */
            private List<Map<String, String>> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
            private List<Map<Integer, String>> keyList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);

            int i = 0;
            JSONArray jsonArray = new JSONArray();

            JSONArray allLst = new JSONArray();

            @Override
            public void invoke(Map<String, String> data, AnalysisContext context) {
                if (null == data.get(1) && null == data.get(2)) {
                    i++;
                    JSONObject jsonObject = new JSONObject();
//                    jsonObject.put("number",context.getCurrentRowNum());
                    jsonObject.put("value1", data.get(0));
                    jsonObject.put("label1", data.get(0));
                    jsonArray.add(jsonObject);
                    allLst.add(jsonObject);
                } else {
//                    System.out.println(allLst);
                    if (allLst.size() > 0) {

                        JSONObject tmp = new JSONObject();

                        for (int j = 2; j < 6; j++) {
                            if (null == data.get(j)) {
                                System.out.printf("====> " + context.getCurrentRowNum() + data.get(0 )+ "\n");
                            }
                        }

//                        Strings.nullToEmpty()

                        tmp.put("value", data.get(0));
//                        tmp.put("fun1", data.get(1));
//                        tmp.put("time1", data.get(2));
//                        tmp.put("price1", data.get(3));
//                        tmp.put("req1", data.get(4));
//                        tmp.put("factor1", data.get(5));
//                        tmp.put("sign1", data.get(6));
//                        tmp.put("mark1", data.get(7));

                        tmp.put("fun1", data.get(1));
                        tmp.put("time1", Strings.nullToEmpty(data.get(2)));
                        tmp.put("price1", Strings.nullToEmpty(data.get(3)));
                        tmp.put("req1", Strings.nullToEmpty(data.get(4)));
                        tmp.put("factor1", Strings.nullToEmpty(data.get(5)));
                        tmp.put("sign1", Strings.nullToEmpty(data.get(6)));
                        tmp.put("mark1", Strings.nullToEmpty(data.get(7)));


                        JSONObject jsonObject = JSONObject.from(allLst.get(allLst.size() - 1));
                        if (null == jsonObject.get("data")) {
                            JSONArray jsonArrayData = new JSONArray();

                            jsonArrayData.add(tmp);
                            jsonObject.put("data", jsonArrayData);
                        } else {
                            jsonObject.getJSONArray("data").add(tmp);
                        }
                        cachedDataList.add(data);
                    }

                }
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                try {
                    saveData();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }

            /**
             * 加上存储数据库
             */
            private void saveData() throws IOException {
//                log.info("{}条数据，开始存储数据库！", cachedDataList.size());
//                log.info("{}条数据，开始存储数据库！", keyList.size());
                log.info("{}条数据，开始存储数据库！", allLst.size());
                log.info("======>\n {}", allLst);

                Files.write(Paths.get("C:\\test\\data3.json"), JSON.toJSONBytes(allLst));
                log.info("======>\n");
            }
        }).sheet().doRead();

    }

}
