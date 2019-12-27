package com.example.easypoi.controller;

import cn.afterturn.easypoi.entity.vo.MapExcelConstants;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import com.example.easypoi.entity.People;
import com.example.easypoi.excel.ExcelUtils;
import com.example.easypoi.file.FileUtils;
import com.google.common.collect.Lists;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author duanyuantong
 * @version Id: IndexController, v 0.1 2019-11-20 16:59 duanyuantong Exp $
 */

@Controller
public class IndexController {

    //    @RequestMapping(value = "/index")
    //    public String index(HttpServletResponse response) throws IOException {
    //        return "index";
    //    }

    @RequestMapping(value = "/index")
    public String load(HttpServletResponse response) throws IOException {
        List<People> list = Lists.newArrayList();
        People people = new People();
        people.setName("李四");
        File file = new File(" user/yinzhe/dream-project/easypoi/src/main/resources/static/img/image.jpg");
        people.setPic(FileUtils.file2byte(file));
        ExcelUtils.exportExcel(list, "员工表", "员工表", People.class, "工作表.xlsx", response);

        return "index";
    }

    @RequestMapping(value = "/table")
    public String table(ModelMap mmp) throws Exception {
        ImportParams params = new ImportParams();
        params.setTitleRows(0);
        params.setHeadRows(1);
        List<People> list = ExcelImportUtil.importExcel(new File("ceshi.xlsx"), People.class, params);
        List<People> peopleList = ExcelUtils.importExcel("ceshi.xlsx", 0, 1, People.class);
        mmp.put("lists", peopleList);
        return "table";
    }

    @RequestMapping(value = "/update")
    public String update() {
        return "update";
    }

    @RequestMapping(value = "/save", method = RequestMethod.POST)
    public String file(@RequestParam("filePath") MultipartFile filePath) {
        List<People> peopleList = ExcelUtils.importExcel(filePath, 0, 1, People.class);
        return "update";
    }

    @RequestMapping(value = "exportXls")
    public String exportXls(ModelMap modelMap, HttpServletRequest request) {

        modelMap.put(MapExcelConstants.ENTITY_LIST, newAddExcel());
        modelMap.put(MapExcelConstants.MAP_LIST, sourceData());
        modelMap.put(MapExcelConstants.FILE_NAME, "新增数据统计");
        modelMap.put(MapExcelConstants.PARAMS, new ExportParams("新增数据统计", "导出人:铁头娃" , "导出信息"));

        Workbook workbook = ExcelExportUtil.exportExcel(sourceData(), ExcelType.HSSF);
        FileOutputStream fos= null;
        try {
            fos = new FileOutputStream("数据统计.xlsx");
            workbook.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return "update";
    }

    private List<ExcelExportEntity> newAddExcel() {
        List<ExcelExportEntity> entityList = Lists.newArrayList();
        ExcelExportEntity xh = new ExcelExportEntity("序号", "nu", 30);
        xh.setFormat("isAddIndex");
        entityList.add(xh);
        entityList.add(new ExcelExportEntity("新增时间", "clrsj", 30));
        entityList.add(new ExcelExportEntity("案件名称", "cajmc", 50));
        entityList.add(new ExcelExportEntity("案件类型", "cajlx", 30));
        entityList.add(new ExcelExportEntity("案件编号", "cajbh", 40));
        entityList.add(new ExcelExportEntity("当事人姓名", "cdsrxm", 25));
        entityList.add(new ExcelExportEntity("性别", "sex", 15));
        entityList.add(new ExcelExportEntity("当事人身份证号", "cdsrsfz", 40));
        entityList.add(new ExcelExportEntity("代保管金额（元）", "cdbgje", 50));
        entityList.add(new ExcelExportEntity("代保管物品（件）", "cdbgwp", 30));
        return entityList;
    }

    private List<Map<String, Object>> sourceData() {

        Map<String, Object> map = new LinkedHashMap<>();
        map.put("clrsj", "2018-07-23 11:20:30");
        map.put("cajmc", "李刚儿子撞人案");
        Map<String, Object> map2 = new LinkedHashMap<>();
        map2.put("clrsj", "2018-07-23 11:24:30");
        map2.put("cajmc", "黄志慧打人案");
        List<Map<String, Object>> newData = Lists.newArrayList();
        newData.add(map);
        newData.add(map2);

        return newData;

    }

}
