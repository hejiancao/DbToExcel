# DbToExcel
从数据库读取大量数据导入到excle,避免内存溢出解决方案

## 问题
当从数据库读取几十万数据然后导入到excel时，往往会发生堆溢出问题，一方面是因为从数据库读取数据放到List中，数据太多，
另一方面是因为传统的HSSFWorkbook对象，把数据都放在内存中然后一次性写入磁盘，如果数据量太大的话就会内存溢出。
知道原因了那么我们可以从以下的方面来解决问题

## 解决方案
1. 从数据库分页读取
2. 使用XSSFWorkbook对象,设置内存中保留的数据量，超过的写入磁盘


## 项目结构
1. BigDBExportToExcel类：一个简单的demo，实现分页查询和分sheet导入
2. ExportDemo类： 封装了一个ExcelUtils工具类，在实际开发中，可以按照这种模板去导出数据

