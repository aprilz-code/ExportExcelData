package com.wm.file.service.impl;

import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.wm.file.entity.MsgClient;
import com.wm.file.util.MyExcelExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CountDownLatch;

/**
 * @ClassName: IAsynExportExcelServiceImpl
 * @Description:
 * @Author: WM
 * @Date: 2021-08-06 20:06
 **/
public class MyThreadA implements Runnable {

    private Map<String, Object> map;
    private CountDownLatch cdl;
    private List<Object> list = new ArrayList<>();

    private MyExcelExportUtil myExcelExportUtil;

    public MyThreadA(Map<String, Object> map, CountDownLatch cdl,MyExcelExportUtil myExcelExportUtil) {
        this.map = map;
        this.cdl = cdl;
        this.myExcelExportUtil =myExcelExportUtil;
        for (int i = 0; i < 1000000; i++) {  //模拟库中100万数据量
            MsgClient client = new MsgClient();
            client.setBirthday(new Date());
            client.setClientName("小明xxxsxsxsxsxsxsxsxsxsx" + i);
            client.setClientPhone("18797" + i);
            client.setCreateBy("JueYue");
            client.setId("1" + i);
            client.setRemark("测试" + i);
            list.add(client);
        }
    }

    public void run() {
        long start = System.currentTimeMillis();
        int currentPage = (int) map.get("page");
        int pageSize = (int) map.get("limit");
        List subList = new ArrayList(page(list, pageSize, currentPage));
        int count = subList.size();
        System.out.println("线程：" + Thread.currentThread().getName() + " , 读取数据，耗时 ：" + (System.currentTimeMillis() - start));
        String filePath = map.get("path").toString() + map.get("page") + ".xlsx";
        // 调用导出的文件方法
        Workbook workbook = myExcelExportUtil.getWorkbook("计算机一班学生", "学生", MsgClient.class, subList, ExcelType.XSSF);
        File file = new File(filePath);
        MyExcelExportUtil.exportExcel2(workbook, file);
        long end = System.currentTimeMillis();
        System.out.println("线程：" + Thread.currentThread().getName() + " , 导出excel" + map.get("page") + ".xlsx成功 , 导出数据：" + count + " ,耗时 ：" + (end - start) + "ms");
        // 执行完线程数减1
        cdl.countDown();
        System.out.println("剩余任务数  ===========================> " + cdl.getCount());
    }

    // 手动分页方法
    public List page(List list, int pageSize, int page) {
        int totalcount = list.size();
        int pagecount = 0;
        int m = totalcount % pageSize;
        if (m > 0) {
            pagecount = totalcount / pageSize + 1;
        } else {
            pagecount = totalcount / pageSize;
        }
        List<Integer> subList = new ArrayList<>();
        if (pagecount < page) {
            return subList;
        }

        if (m == 0) {
            subList = list.subList((page - 1) * pageSize, pageSize * (page));
        } else {
            if (page == pagecount) {
                subList = list.subList((page - 1) * pageSize, totalcount);
            } else {
                subList = list.subList((page - 1) * pageSize, pageSize * (page));
            }
        }
        return subList;
    }


}
