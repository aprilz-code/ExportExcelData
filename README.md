# 多线程实现百万级数据导出到excel

## **考虑前提：**

大数据量导出到文件，首先需要考虑的是内存溢出的场景：数据库读取数据到内存中、
将数据写入到excel进行大量的IO操作。
然后考虑到一个文件数据过大，用户打开慢，体验不好。
针对这些问题的考虑，采用多线程的方式多个线程同时处理查询数据，
一个线程生成一个excel，最后在合并数据返回，以达到提高效率的目的。

## **实现思路：**

1. 计算导出数据的总条数：dataTotalCount
2. 合理设置每个excel的数据的数量（避免打开一个excel时间过长）：LIMIT 		
3. 计算出需要导出的excel个数（线程个数）：count=dataTotalCount/ LIMIT + (dataTotalCount% LIMIT > 0 ? 1 : 0)
4. 将分页、生成文件路径信息，初始化到一个队列里面，队列的大小是线程的数量，对每个文件开启一个线程，异步执行导出，文件全部导出结束，此时异步转成同步，将最终生成的excel文件生成zip压缩包。



万级以上，HSSFWorkbook 和XSSFWorkbook 容易报内存不足
HSSFWorkbook 用于Excel2003版及更早版本(扩展名为.xls)的导出。
XSSFWorkbook 用于Excel2007版(扩展名为.xlsx)的导出。

故
```java
public class AsynExcelExportUtil {
    
    // .... somethings
    
    private static Workbook getWorkbook(ExcelType type, int size) {
        if (ExcelType.HSSF.equals(type)) {
            return new HSSFWorkbook();
        } else {
            return size < USE_SXSSF_LIMIT ? new XSSFWorkbook() : new SXSSFWorkbook();
        }
    }
}
```

每页显示 40000条
400000   -27750ms

1000000  8个线程 16G内存    74546ms~112475ms
1000000  16个线程 16G内存  41770ms~67382ms
1000000  17个线程 16G内存  38590ms~ 78863ms
