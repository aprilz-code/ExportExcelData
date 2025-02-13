package com.wm.file.util;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import com.wm.file.util.easypoi.ExcelStyleUtil;
import com.wm.file.util.easypoi.MyExcelExportService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: ExcelExportUtil
 * @Description: Excel导出工具类
 * @Author: WM
 * @Date: 2021-07-24 18:47
 **/
@Slf4j
@Component
public class MyExcelExportUtil {
    @Resource
    private IExcelExportServer excelExportServer;
    /**
     * 小量数据允许导出的最大条数
     */
    private static final Integer EXPORT_EXCEL_BASE_MAX_NUM = 100000;
    public static int USE_SXSSF_LIMIT = 10000;//万级以上，HSSFWorkbook 和XSSFWorkbook 容易报内存不足
    //HSSFWorkbook 用于Excel2003版及更早版本(扩展名为.xls)的导出。
    //
    //XSSFWorkbook 用于Excel2007版(扩展名为.xlsx)的导出。

    /**
     * 获取导出的 Workbook对象
     * 普通导出
     *
     * @param title     大标题
     * @param sheetName 页签名
     * @param object    导出实体
     * @param list      普通导出传入的数据集合
     * @param list      数据集合
     * @return Workbook
     */
    public static Workbook getWorkbook(String title, String sheetName, Class<?> object, List<?> list, ExcelType excelType) {
        // 判断导出数据是否为空
        if (list == null) {
            list = new ArrayList<>();
        }
        // 判断导出数据数量是否超过限定值
//        if (list.size() > EXPORT_EXCEL_BASE_MAX_NUM) {
//            title = "导出数据行数超过:" + EXPORT_EXCEL_BASE_MAX_NUM + "条，无法导出！";
//            list = new ArrayList<>();
//        }
        // 获取导出参数
        ExportParams exportParams = new ExportParams(title, sheetName, excelType);
        // 设置导出样式
        exportParams.setStyle(ExcelStyleUtil.class);
        // 设置行高
        exportParams.setHeight((short) 8);
        // 普通导出，输出Workbook流
//        return ExcelExportUtil.exportExcel(exportParams, object, list);
        return createExcel(exportParams, object, list);
    }

    /**
     * 导出大数据量
     *
     * @param title
     * @param sheetName
     * @param object
     * @param queryParams 大数据量导出，查询数据的参数条件
     * @param excelType
     * @return
     */
    public Workbook getWorkbook(String title, String sheetName, Class<?> object, Object queryParams, ExcelType excelType) {
        // 获取导出参数
        ExportParams exportParams = new ExportParams(title, sheetName, excelType);
        // 设置导出样式
        exportParams.setStyle(ExcelStyleUtil.class);
        // 设置行高
        exportParams.setHeight((short) 6);
        // 导出大数据量，输出Workbook流
        return ExcelExportUtil.exportBigExcel(exportParams, object, excelExportServer, queryParams);
    }

    /**
     * 获取导出的 Workbook对象
     *
     * @param path 模板路径
     * @param map  导出内容map
     * @return Workbook
     */
    public static Workbook getWorkbook(String path, Map<String, Object> map) {
        // 获取导出模板
        TemplateExportParams params = new TemplateExportParams(path);
        // 设置导出样式
        params.setStyle(ExcelStyleUtil.class);
        // 输出Workbook流
        return ExcelExportUtil.exportExcel(params, map);
    }

    /**
     * 导出Excel
     *
     * @param workbook workbook流
     * @param fileName 文件名
     * @param response 响应
     */
    public static void exportExcel(Workbook workbook, String fileName, HttpServletResponse response) {
        // 输出文件
        try (OutputStream out = response.getOutputStream()) {
            // 获取文件名并转码
            String name = URLEncoder.encode(fileName, "UTF-8");
            // 编码
            response.setCharacterEncoding("UTF-8");
            // 设置强制下载不打开
            response.setContentType("application/force-download");
            // 下载文件的默认名称
            response.setHeader("Content-Disposition", "attachment;filename=" + name + ".xlsx");
            // 输出表格
            workbook.write(out);
        } catch (IOException e) {
            log.error("文件导出异常，详情如下:", e);
            throw new RuntimeException("文件导出异常");
        } finally {
            try {
                if (workbook != null) {
                    // 关闭输出流
                    workbook.close();
                }
            } catch (IOException e) {
                log.error("文件导出异常,详情如下:", e);
            }
        }
    }

    public static void exportExcel2(Workbook workbook, File file) {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(file);
            // 输出表格
            workbook.write(out);
        } catch (IOException e) {
            log.error("文件导出异常，详情如下:", e);
            throw new RuntimeException("文件导出异常");
        } finally {
            try {
                if (workbook != null) {
                    // 关闭输出流
                    workbook.close();
                }
                if (out != null) {
                    // 关闭输出流
                    out.close();
                }
            } catch (IOException e) {
                log.error("文件导出异常,详情如下:", e);
            }
        }
    }


    /**
     * 生成Workbook
     *
     * @param entity
     * @param pojoClass
     * @param dataSet
     * @return
     */
    public static Workbook createExcel(ExportParams entity, Class<?> pojoClass, Collection<?> dataSet) {
        Workbook workbook = getWorkbook(entity.getType(), dataSet.size());
        (new MyExcelExportService()).createSheet(workbook, entity, pojoClass, dataSet);
        return workbook;
    }

    private static Workbook getWorkbook(ExcelType type, int size) {
        if (ExcelType.HSSF.equals(type)) {
            return new HSSFWorkbook();
        } else {
            return size < USE_SXSSF_LIMIT ? new XSSFWorkbook() : new SXSSFWorkbook();
        }
    }
}
