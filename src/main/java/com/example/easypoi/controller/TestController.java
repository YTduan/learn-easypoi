package com.example.easypoi.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import com.example.easypoi.excel.ExcelUtils;
import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.util.*;

/**
 * @author duanyuantong
 * @version Id: TestController, v 0.1 2019-11-25 16:34 duanyuantong Exp $
 */
@Controller
public class TestController {

    /**
     * 根据Map创建对应的Excel
     * @param entity
     *            表格标题属性
     * @param entityList
     *            Map对象列表
     * @param dataSet
     *            Excel对象数据List
     */
    public static Workbook exportExcel(ExportParams entity, List<ExcelExportEntity> entityList, Collection<?> dataSet) {
        //        Workbook workbook = getWorkbook(entity.getType(),dataSet.size());
        Workbook workbook = ExcelExportUtil.exportExcel(entity, entityList, dataSet);
        new ExcelExportService().createSheetForMap(workbook, entity, entityList, dataSet);
        return workbook;
    }

    @RequestMapping("/test")
    public String export(HttpServletResponse response) {

        //表头设置
        List<ExcelExportEntity> colList = new ArrayList<ExcelExportEntity>();

        ExcelExportEntity colEntity = new ExcelExportEntity("生产日期", "time");
        colEntity.setNeedMerge(true);
        colList.add(colEntity);

        colEntity = new ExcelExportEntity("产品名称", "name");
        colEntity.setNeedMerge(true);
        colList.add(colEntity);

//        colEntity = new ExcelExportEntity("产品UV", "uv");
//        colEntity.setNeedMerge(true);
//        colList.add(colEntity);

        ExcelExportEntity group_1 = new ExcelExportEntity("业务数据", "data");
        List<ExcelExportEntity> exportEntities = new ArrayList<>();
            exportEntities.add(new ExcelExportEntity("长度(cm)", "length" ));
            exportEntities.add(new ExcelExportEntity("宽度(cm)", "width"));
            exportEntities.add(new ExcelExportEntity("高度(cm)", "height"));
            exportEntities.add(new ExcelExportEntity("弧度(cm)", "radian"));
        group_1.setList(exportEntities);
        colList.add(group_1);

//        for (int i = 0; i < 1; i++) {
//            Map<String, Object> valMap = new HashMap<String, Object>();
//            valMap.put("dt", "日期" + i);
//            valMap.put("pv", "pv" + i);
//            valMap.put("uv", "uv" + i);
//            {
//                List<Map<String, Object>> list_1 = new ArrayList<Map<String, Object>>();
//                Map<String, Object> valMap_1 = new HashMap<String, Object>();
//                for (int j = 1; j < 5; j++) {
//                    valMap_1.put("data" + j, "数据" + j);
//                }
//                list_1.add(valMap_1);
//
//                valMap.put("businessData", list_1);
//            }
//
//        }
        //文件数据
        List<Map<String, Object>> list = sourceData();

        //导出
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("产品表", "数据"), colList, list);
        setResponseHeader(response, "产品表");
        write(workbook, response);
        return "/update";
    }

    /**
     * 设置导出文件头
     *
     * @param response 相应servlet
     * @param fileName 文件名称--不能是中文
     */
    private static void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            LocalDateTime localTime=LocalDateTime.now();
            fileName = fileName+localTime.getYear()+localTime.getMonthValue()+localTime.getDayOfMonth()+localTime.getHour()+localTime.getMinute()+localTime.getSecond()+".xls";
            response.setCharacterEncoding("UTF-8");
//            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setContentType("text/html;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        } catch (Exception ex) {

        }
    }

    /**
     * 将book 导出excel
     *
     * @param book
     * @param response
     */
    private static void write(Workbook book, HttpServletResponse response) {
        try {
            OutputStream os = response.getOutputStream();
            book.write(os);
            os.flush();
            os.close();
        } catch (IOException e) {

        }
    }

    private List<Map<String, Object>> sourceData() {
        List<Map<String, Object>> newData = Lists.newArrayList();//返回的数据
        Map<String, Object> map = new LinkedHashMap<>();
        map.put("time", "2018-07-23 11:20:30");
        map.put("name", "水杯");
        List<Map<String,Object>> list = Lists.newArrayList();

        Map<String,Object> data = new LinkedHashMap<>();
        data.put("length",10);
        data.put("width",10);
        data.put("height",30);
        data.put("radian",5);
        list.add(data);
        map.put("data",list);
        newData.add(map);


        return newData;

    }

    private List<ExcelExportEntity> excelHeader(){
        //表头设置
        List<ExcelExportEntity> colList = new ArrayList<ExcelExportEntity>();

        ExcelExportEntity colEntity = new ExcelExportEntity("生产日期", "time");
        colEntity.setNeedMerge(true);
        colList.add(colEntity);

        colEntity = new ExcelExportEntity("产品名称", "name");
        colEntity.setNeedMerge(true);
        colList.add(colEntity);

//        colEntity = new ExcelExportEntity("产品UV", "uv");
//        colEntity.setNeedMerge(true);
//        colList.add(colEntity);

        ExcelExportEntity group_1 = new ExcelExportEntity("业务数据", "data");
        List<ExcelExportEntity> exportEntities = new ArrayList<>();
        exportEntities.add(new ExcelExportEntity("长度(cm)", "length" ));
        exportEntities.add(new ExcelExportEntity("宽度(cm)", "width"));
        exportEntities.add(new ExcelExportEntity("高度(cm)", "height"));
        exportEntities.add(new ExcelExportEntity("弧度(cm)", "radian"));
        group_1.setList(exportEntities);
        colList.add(group_1);
        return colList;
    }

    @RequestMapping(value = "/ceshi")
    public void ceshi(HttpServletResponse response){
        ExcelUtils.exportExcel("产品表","产品表","产品",excelHeader(),sourceData(),response);
    }

}
