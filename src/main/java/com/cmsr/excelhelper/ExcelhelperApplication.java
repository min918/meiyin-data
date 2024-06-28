package com.cmsr.excelhelper;

import com.alibaba.excel.EasyExcel;
import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

@SpringBootApplication
@Slf4j
public class ExcelhelperApplication implements CommandLineRunner {

	public static void main(String[] args) {
		SpringApplication.run(ExcelhelperApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {
		String fileName = "C:\\test\\项目6.20.xlsx";
		// 这里 也可以不指定class，返回一个list，然后读取第一个sheet 同步读取会自动finish
		List<Map<Integer, String>> listMap = EasyExcel.read(fileName).sheet().doReadSync();
		JSONArray jsonArray = new JSONArray();
		for (Map<Integer, String> data : listMap) {
			// 返回每条数据的键值对 表示所在的列 和所在列的值
			log.info("读取到数据:{}", JSON.toJSONString(data));
			JSONObject jsonObject = new JSONObject();
			jsonObject.put("value",data.get(0));
			jsonArray.add(jsonObject);
		}

		Files.write(Paths.get("C:\\test\\project.json"), jsonArray.toJSONString().getBytes());






	}

}
