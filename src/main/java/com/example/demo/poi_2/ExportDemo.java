package com.example.demo.poi_2;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.hibernate.transform.Transformers;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.persistence.EntityManager;
import javax.persistence.Query;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**从数据库查询大量数据导入excel解决方案：分页查询+SXSSFWorkbook内存刷新
 * @author created by shaos on 2019/11/26
 */
@RestController
@RequestMapping("/api/poi/")
public class ExportDemo {

    private Logger logger = LoggerFactory.getLogger(this.getClass());
    private static Map<String, String> headMap = new LinkedHashMap<>();

    static {
        headMap.put("id", "订单ID");
        headMap.put("name", "订单名称");
        headMap.put("userName", "客户名称");
    }

    private EntityManager entityManager;

    public ExportDemo(EntityManager entityManager) {
        this.entityManager = entityManager;
    }


    /** 分页导出excel案例
     * @author shaos
     * @date 2019/11/26 18:24
     */
    @GetMapping("/export")
    public void export(HttpServletResponse response) throws IOException {

        long startTime = System.currentTimeMillis();
        Query query = entityManager.createNativeQuery("select count(1) from t_order");
        Object singleResult = query.getSingleResult();
        Integer totalCount = Integer.valueOf(String.valueOf(singleResult));
        int PAGESIZE = 10000;
        int totalPage = totalCount%PAGESIZE == 0 ? totalCount/PAGESIZE : totalCount/PAGESIZE + 1;
        int pageStart = 0;
        logger.info(String.format("共计[%s]条数据,每页[%s]条,共计[%s]页", totalCount, PAGESIZE, totalPage));


        String[] titles = new String[1];
        Map[] maps = new Map[1];
        titles[0] = "订单信息";
        maps[0] = headMap;
        JSONArray[] jsonArrays = new JSONArray[1];
        jsonArrays[0] = new JSONArray(new ArrayList<>());


        List<ResultMapDto> pageList = new ArrayList<>();
        SXSSFWorkbook workbook = new SXSSFWorkbook(1000);


        // 循环导出，每页10000条数据
        for (int i = 0; i < totalPage; i++) {
            logger.info(String.format("导出第[%s]页，from[%s]to[%s]", i + 1, pageStart, pageStart + PAGESIZE -1));
            Query nativeQuery = entityManager.createNativeQuery("select * from t_order limit :pageStart,:pageSize")
                    .setParameter("pageStart", pageStart)
                    .setParameter("pageSize", PAGESIZE);
            List list = nativeQuery.unwrap(org.hibernate.query.Query.class).setResultTransformer(Transformers.aliasToBean(JSONObject.class)).list();
            for (Object o : list) {
                String jsonString = JSONObject.toJSONString(o);
                ResultMapDto record = JSONObject.parseObject(jsonString, ResultMapDto.class);
                pageList.add(record);
            }

            jsonArrays[0] = new JSONArray(new ArrayList<>(pageList));
            ExcelUtils.exportExcelByMultiSheet(titles, maps, jsonArrays, null, 0, workbook, i);
            pageStart = (i + 1) * PAGESIZE;
            pageList.clear();

        }

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename="
                + URLEncoder.encode(titles[0],"UTF-8")+".xlsx");
        OutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        // 刷新缓冲区
        outputStream.flush();
        // 不要对HttpServletResponse对象的ServletOutputStream做流关闭处理,否则会报UT010029: Stream is closed异常
        // outputStream.close();
        // 释放内存空间
        workbook.dispose();

        long stopTime = System.currentTimeMillis();
        logger.info("导出到excel时间共计: " + (stopTime - startTime)/1000 + "s");

    }



}
