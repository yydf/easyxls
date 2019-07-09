# easyxls
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/cn.4coder/easyxls/badge.svg)](https://maven-badges.herokuapp.com/maven-central/cn.4coder/easyxls/)
[![GitHub release](https://img.shields.io/github/release/yydf/easyxls.svg)](https://github.com/yydf/easyxls/releases)
[![license](https://img.shields.io/github/license/mashape/apistatus.svg)](https://raw.githubusercontent.com/yydf/easyxls/master/LICENSE)
![Jar Size](https://img.shields.io/badge/jar--size-18.3k-blue.svg)

环境
-------------
- JDK 7

如何使用
-----------------------

* 添加dependency到POM文件::

```
<dependency>
    <groupId>cn.4coder</groupId>
    <artifactId>easyxls</artifactId>
    <version>0.0.2</version>
</dependency>
```

* 编码:

```
Workbook w = new Workbook("test");
w.addTitle("删掉");
w.addTitle("订单");
w.addData("sdf", 1214);
w.addData("sdf", 679);
Sheet sheet = w.addSheet("test1");
sheet.addTitle("大小");
sheet.addTitle("多少");
sheet.addData("sdfsdg", 234, "abc");
sheet.addData("sdfsdg2", 2234, "abcd");
w.write(new FileOutputStream(new File("d://sdff.xlsx")));
w.close();
```
