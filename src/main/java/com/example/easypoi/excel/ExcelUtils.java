package com.example.easypoi.excel;


import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;

/**
 * @author duanyuantong
 * @version Id: ExcelUtils, v 0.1 2019-11-20 20:49 duanyuantong Exp $
 */
public class ExcelUtils {

    /**
     * 导出excel
     * @param list 对象集合
     * @param title 表头
     * @param sheetName sheet名称
     * @param pojoClass 实体类
     * @param fileName 导出的文件名
     * @param isCreateHeader
     * @param response
     */
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, boolean isCreateHeader, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setCreateHeadRows(isCreateHeader);
        defaultExport(list, pojoClass, fileName, response, exportParams);
    }

    /**
     *
     * @param list
     * @param title
     * @param sheetName
     * @param pojoClass
     * @param fileName
     * @param response
     */
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, HttpServletResponse response) {
        defaultExport(list, pojoClass, fileName, response, new ExportParams(title, sheetName));
    }

    /**
     * 导出Excel 数据量小于6万条
     * @param list
     * @param title
     * @param sheetName
     * @param pojoClass
     * @param fileName
     */
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName) {
        defaultExport(list, pojoClass, fileName, new ExportParams(title, sheetName));
    }

    /**
     * 导出Excel 数据大于6万条
     * @param list
     * @param title
     * @param sheetName
     * @param pojoClass
     * @param fileName
     */
    public static void exportBigExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName) {
        defaultBigExport(list, pojoClass, fileName, new ExportParams(title, sheetName));
    }

    /**
     * 数据小于6万条
     * @param list
     * @param fileName
     * @param response
     */
    public static void exportExcel(List<Map<String, Object>> list, String fileName, HttpServletResponse response) {
        defaultExport(list, fileName, response);
    }


    /**
     *  读取模版，按照模版下载
     * @param param 模版
     * @param map 数据
     * @param fileName 文件名
     * @param response
     */
    public static void exportExcel(TemplateExportParams param,Map<String,Object> map,String fileName,HttpServletResponse response){
        setResponseHeader(response,fileName);
        Workbook workbook = ExcelExportUtil.exportExcel(param, map);
        defaultExport(workbook,response);
    }

    /**
     * 数据小于6万条
     * @param list
     * @param pojoClass
     * @param fileName
     * @param response
     * @param exportParams
     */
    private static void defaultExport(List<?> list, Class<?> pojoClass, String fileName, HttpServletResponse response, ExportParams exportParams) {
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, pojoClass, list);
        if (workbook != null);downLoadExcel(fileName, response, workbook);
    }

    /**
     * 默认的导出Excel，数据量小于6万条
     * @param list
     * @param pojoClass
     * @param fileName
     * @param exportParams
     */
    private static void defaultExport(List<?> list, Class<?> pojoClass, String fileName, ExportParams exportParams) {
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, pojoClass, list);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(fileName);
            workbook.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    /**
     * 默认的导出Excel，数据量大于6万条
     * @param list
     * @param pojoClass
     * @param fileName
     * @param exportParams
     */
    private static void defaultBigExport(List<?> list, Class<?> pojoClass, String fileName, ExportParams exportParams) {
        Workbook workbook = ExcelExportUtil.exportBigExcel(exportParams, pojoClass, list);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(fileName);
            workbook.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    /**
     * 导出的数据大于6万条
     * @param list
     * @param pojoClass
     * @param fileName
     * @param response
     * @param exportParams
     */
    private static void exportBigExcel(List<?> list, Class<?> pojoClass, String fileName, HttpServletResponse response, ExportParams exportParams) {
        Workbook workbook = ExcelExportUtil.exportBigExcel(exportParams, pojoClass, list);
        if (workbook != null);downLoadExcel(fileName, response, workbook);
    }

    /**
     *
     * @param fileName
     * @param response
     * @param workbook
     */
    private static void downLoadExcel(String fileName, HttpServletResponse response, Workbook workbook) {
        try {
            response.setCharacterEncoding("UTF-8");
//            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setContentType("text/html;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            workbook.write(generateResponseExcel(fileName,response));
        } catch (IOException e) {
            e.printStackTrace();
            //throw new NormalException(e.getMessage());
        }
    }

    /**
     *
     * @param list
     * @param fileName
     * @param response
     */
    private static void defaultExport(List<Map<String, Object>> list, String fileName, HttpServletResponse response) {
        Workbook workbook = ExcelExportUtil.exportExcel(list, ExcelType.HSSF);
        if (workbook != null);downLoadExcel(fileName, response, workbook);
    }

    /**
     *
     * @param filePath 文件路径
     * @param titleRows 表格标题行数
     * @param headerRows 表头行数
     * @param pojoClass 实体类
     * @param <T>
     * @return
     */
    public static <T> List<T> importExcel(String filePath, Integer titleRows, Integer headerRows, Class<T> pojoClass) {
        if (StringUtils.isBlank(filePath)) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setTitleRows(titleRows);
        params.setHeadRows(headerRows);
        List<T> list = null;
        try {
            list = ExcelImportUtil.importExcel(new File(filePath), pojoClass, params);
        } catch (NoSuchElementException e) {
            //throw new NormalException("模板不能为空");
        } catch (Exception e) {
            e.printStackTrace();
            //throw new NormalException(e.getMessage());
        }
        return list;
    }

    /**
     * 表格标题行数,默认0,表头行数,默认1
     * @param filePath 文件路径
     * @param pojoClass 实体类
     * @param <T>
     * @return
     */
    public static <T> List<T> importExcel(String filePath, Class<T> pojoClass) {
        if (StringUtils.isBlank(filePath)) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setTitleRows(0);
        params.setHeadRows(1);
        List<T> list = null;
        try {
            list = ExcelImportUtil.importExcel(new File(filePath), pojoClass, params);
        } catch (NoSuchElementException e) {
            //throw new NormalException("模板不能为空");
        } catch (Exception e) {
            e.printStackTrace();
            //throw new NormalException(e.getMessage());
        }
        return list;
    }

    /**
     *
     * @param file 文件
     * @param titleRows 表格标题行数
     * @param headerRows 表头行数
     * @param pojoClass 实体类
     * @param <T>
     * @return
     */
    public static <T> List<T> importExcel(MultipartFile file, Integer titleRows, Integer headerRows, Class<T> pojoClass) {
        if (file == null) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setTitleRows(titleRows);
        params.setHeadRows(headerRows);
        List<T> list = null;
        try {
            list = ExcelImportUtil.importExcel(file.getInputStream(), pojoClass, params);
        } catch (NoSuchElementException e) {
            // throw new NormalException("excel文件不能为空");
        } catch (Exception e) {
            //throw new NormalException(e.getMessage());
            System.out.println(e.getMessage());
        }
        return list;

    }

    /**
     * 表格标题行数,默认0,表头行数,默认1
     * @param file 文件
     * @param pojoClass 实体类
     * @param <T>
     * @return
     */
    public static <T> List<T> importExcel(MultipartFile file, Class<T> pojoClass) {
        if (file == null) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setKeyIndex(0);
        params.setTitleRows(0);
        params.setHeadRows(1);
        List<T> list = null;
        try {
            list = ExcelImportUtil.importExcel(file.getInputStream(), pojoClass, params);
        } catch (NoSuchElementException e) {
            // throw new NormalException("excel文件不能为空");
        } catch (Exception e) {
            //throw new NormalException(e.getMessage());
            System.out.println(e.getMessage());
        }
        return list;

    }
    public static void exportBigExcel(String fileName, HttpServletResponse response, Workbook workbook) {
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setCharacterEncoding("UTF-8");
        try {
            response.setHeader("Content-Disposition", "attachment;filename="+
                    java.net.URLEncoder.encode(fileName,"UTF-8"));
        } catch (UnsupportedEncodingException e) {
            //e.printStackTrace();
//            return AitgResponse.err(AitgResponse.CODE_ERROR, e.getMessage());

        }
        try {
            OutputStream os = response.getOutputStream();
            workbook.write(os);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 复杂表头数据导出
     * @param fileName
     * @param title
     * @param sheetName
     * @param entityList
     * @param data
     * @param response
     */
    public static void exportExcel(String fileName, String title, String sheetName, List<ExcelExportEntity> entityList, Collection<?> data, HttpServletResponse response){
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams(title,sheetName), entityList, data);
        setResponseHeader(response, fileName);
        defaultExport(workbook,response);
    }

    private static void defaultExport(Workbook workbook,HttpServletResponse response){
        try {
            OutputStream os = response.getOutputStream();
            workbook.write(os);
            os.flush();
            os.close();
        } catch (IOException e) {

        }
    }
    private static ServletOutputStream generateResponseExcel(String fileName, HttpServletResponse response) throws IOException {
        LocalDateTime localTime=LocalDateTime.now();
        fileName = fileName+localTime.getYear()+localTime.getMonthValue()+localTime.getDayOfMonth()+localTime.getHour()+localTime.getMinute()+localTime.getSecond()+".xls";
        fileName = fileName == null || "".equals(fileName) ? "excel" : URLEncoder.encode(fileName, "UTF-8");
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName);

        return response.getOutputStream();
    }

    /**
     * 设置导出文件头
     *
     * @param response 相应servlet
     * @param fileName 文件名称带时间信息
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


}
