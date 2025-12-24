[中文](./README.zh-CN.md) | [English](./README.md)


AutoPOI (Excel and Word Easy Utility)
===========================
AutoPOI, as its name suggests "auto", pursues automation. It enables anyone without POI experience to quickly implement Excel import/export and Word template export in a foolproof manner. You can complete Excel import/export with just 5 lines of code.

Current Version: 2.0.4 (Released: 2025-12-21)

---------------------------
Key Features of AutoPOI
--------------------------
1. Elegant design, easy to use
2. Rich interfaces, easy to extend
3. Many default values, write less do more
4. AbstractView support, web export made simple

---------------------------
Main Utility Classes
---------------------------

1. ExcelExportUtil - Excel export (normal export, template export)
2. ExcelImportUtil - Excel import
3. WordExportUtil - Word export (only supports docx, doc version has image bugs in POI, not supported yet)

---------------------------
Difference Between XLS and XLSX Export
---------------------------

1. Export time: XLS is 2-3x faster than XLSX
2. Export size: XLS is 2-3x larger than XLSX or more
3. Need to consider both network speed and local processing speed

---------------------------
Project Modules
---------------------------
1. autopoi-parent - Parent POM
2. autopoi - Core utility package for Excel export/import and Word export
3. autopoi-spring-boot-2-starter - Spring Boot 2.x support (compatible with javax.servlet)
4. autopoi-spring-boot-3-starter - Spring Boot 3.x support (compatible with jakarta.servlet)
5. SAX import uses xercesImpl package (may cause unexpected issues), Word export uses poi-scratchpad, all as optional dependencies

--------------------------
Maven Dependencies
--------------------------

**For Spring Boot 2.x projects:**
```xml
<dependency>
 <groupId>org.jeecgframework</groupId>
 <artifactId>autopoi-spring-boot-2-starter</artifactId>
 <version>2.0.4</version>
</dependency>
```

**For Spring Boot 3.x projects:**
```xml
<dependency>
 <groupId>org.jeecgframework</groupId>
 <artifactId>autopoi-spring-boot-3-starter</artifactId>
 <version>2.0.4</version>
</dependency>
```

**For pure Java projects (without Spring):**
```xml
<dependency>
 <groupId>org.jeecgframework</groupId>
 <artifactId>autopoi</artifactId>
 <version>2.0.4</version>
</dependency>
```

--------------------------
Template Expression Support
--------------------------
- Space separation
- Ternary operator: {{test ? obj:obj2}}
- n: Indicates numeric cell type {{n:}}
- le: Represents length {{le:()}} used in if/else {{le:() > 8 ? obj1 : obj2}}
- fd: Format date {{fd:(obj;yyyy-MM-dd)}}
- fn: Format number {{fn:(obj;###.00)}}
- fe: Iterate data, create row
- !fe: Iterate data without creating row
- $fe: Insert by moving down, move current and below rows down by .size() rows, then insert
- !if: Delete current column {{!if:(test)}}
- Single quotes for constant values '', e.g., '1' outputs 1


---------------------------
Export Examples
---------------------------
1. Annotations - Both import and export are annotation-based. Add annotations to entities to mark export objects and perform operations.

```Java
@ExcelTarget("courseEntity")
public class CourseEntity implements java.io.Serializable {
    /** Primary Key */
    private String id;
    /** Course Name */
    @Excel(name = "Course Name", orderNum = "1", needMerge = true)
    private String name;
    /** Teacher */
    @ExcelEntity(id = "yuwen")
    @ExcelVerify()
    private TeacherEntity teacher;
    /** Math Teacher */
    @ExcelEntity(id = "shuxue")
    private TeacherEntity shuxueteacher;

    @ExcelCollection(name = "Students", orderNum = "4")
    private List<StudentEntity> students;
}
```

2. Basic Export
Pass export parameters, export object, and object list to complete the export.

```Java
HSSFWorkbook workbook = ExcelExportUtil.exportExcel(new ExportParams(
    "2412312", "Test", "Test"), CourseEntity.class, list);
```

3. Export with Index
Set a parameter value to add an index column in the export.

```Java
ExportParams params = new ExportParams("2412312", "Test", "Test");
params.setAddIndex(true);
HSSFWorkbook workbook = ExcelExportUtil.exportExcel(params,
    TeacherEntity.class, telist);
```

4. Export Map
Create annotation-like collections to complete Map export.

```Java
List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
entity.add(new ExcelExportEntity("Name", "name"));
entity.add(new ExcelExportEntity("Gender", "sex"));

List<Map<String, String>> list = new ArrayList<Map<String, String>>();
Map<String, String> map;
for (int i = 0; i < 10; i++) {
    map = new HashMap<String, String>();
    map.put("name", "1" + i);
    map.put("sex", "2" + i);
    list.add(map);
}

HSSFWorkbook workbook = ExcelExportUtil.exportExcel(new ExportParams(
    "Test", "Test"), entity, list);
```

5. Template Export
Complete export based on template configuration.

```Java
TemplateExportParams params = new TemplateExportParams();
params.setHeadingRows(2);
params.setHeadingStartRow(2);
Map<String,Object> map = new HashMap<String, Object>();
map.put("year", "2013");
map.put("sunCourses", list.size());
Map<String,Object> obj = new HashMap<String, Object>();
map.put("obj", obj);
obj.put("name", list.size());
params.setTemplateUrl("org/jeecgframework/poi/excel/doc/exportTemp.xls");
Workbook book = ExcelExportUtil.exportExcel(params, CourseEntity.class, list, map);
```

6. Import
Set import parameters, pass file or stream to get the corresponding list.

```Java
ImportParams params = new ImportParams();
params.setTitleRows(2);
params.setHeadRows(2);
//params.setSheetNum(9);
params.setNeedSave(true);
long start = new Date().getTime();
List<CourseEntity> list = ExcelImportUtil.importExcel(new File(
    "d:/tt.xls"), CourseEntity.class, params);
```

7. Seamless Spring MVC Integration
Excel export done with just a few lines.

```Java
@RequestMapping(value = "/exportXls")
public ModelAndView exportXls(HttpServletRequest request, HttpServletResponse response) {
    ModelAndView mv = new ModelAndView(new JeecgEntityExcelView());
    List<JeecgDemo> pageList = jeecgDemoService.list();
    // Export file name
    mv.addObject(NormalExcelConstants.FILE_NAME, "Export Excel File Name");
    // Annotated object Class
    mv.addObject(NormalExcelConstants.CLASS, JeecgDemo.class);
    // Custom table parameters
    mv.addObject(NormalExcelConstants.PARAMS, new ExportParams("Custom Export Excel Title", "Custom Sheet Name"));
    // Export data list
    mv.addObject(NormalExcelConstants.DATA_LIST, pageList);
    return mv;
}
```

| Custom View | Purpose | Description |
| ------ | ------ | ------ |
| JeecgEntityExcelView | Entity object export view | e.g., List&lt;JeecgDemo&gt; |
| JeecgMapExcelView | Map object export view | List&lt;Map&lt;String, String&gt;&gt; list |
| JeecgTemplateExcelView | Excel template export view | - |
| JeecgTemplateWordView | Word template export view | - |

8. Excel Import Validation
Filter data that doesn't meet rules, append error messages to Excel. Provides common validation rules and generic validation interface.

```Java
/**
 * Email validation
 */
@Excel(name = "Email", width = 25)
@ExcelVerify(isEmail = true, notNull = true)
private String email;

/**
 * Mobile phone validation
 */
@Excel(name = "Mobile", width = 20)
@ExcelVerify(isMobile = true, notNull = true)
private String mobile;

ExcelImportResult<ExcelVerifyEntity> result = ExcelImportUtil.importExcelVerify(
    new File("d:/tt.xls"), ExcelVerifyEntity.class, params);
for (int i = 0; i < result.getList().size(); i++) {
    System.out.println(ReflectionToStringBuilder.toString(result.getList().get(i)));
}
```

9. Import Map
Set import parameters, pass file or stream to get the corresponding list. Custom Key requires implementing IExcelDataHandler interface.

```Java
ImportParams params = new ImportParams();
List<Map<String,Object>> list = ExcelImportUtil.importExcel(new File(
    "d:/tt.xls"), Map.class, params);
```

10. Dictionary Usage
Add dicCode="" in the entity property Excel annotation, where dicCode is the Code of the data dictionary in the jeecg system.

```Java
@Excel(name="Gender", width=15, dicCode="sex")
private java.lang.String sex;
```

11. Dictionary Table Usage
dictTable is the database table name, dicCode is the associated field name, dicText is the field corresponding to the content displayed in Excel.

```Java
@Excel(name="Department", dictTable="t_s_depart", dicCode="id", dicText="departname")
private java.lang.String depart;
```

12. Replace Usage
If database stores 0/1, Excel cells display Female/Male.

```Java
@Excel(name="Test Replace", width=15, replace={"Male_1","Female_0"})
private java.lang.String fdReplace;
```

13. Advanced Field Conversion
- exportConvert: Set to true to replace values during export, add a method with "convert" prefix before the original get method name.
- importConvert: Set to true to replace values during import, add a method with "convert" prefix before the original set method name.

```Java
@Excel(name="Test Convert", width=15, exportConvert=true, importConvert=true)
private java.lang.String fdConvert;

/**
 * Conversion example: Add suffix to the field value
 * @return
 */
public String convertgetFdConvert(){
    return this.fdConvert + " Yuan";
}

/**
 * Conversion example: Replace "Yuan" in Excel cell
 * @return
 */
public void convertsetFdConvert(String fdConvert){
    this.fdConvert = fdConvert.replace(" Yuan", "");
}
```

---------------------------
Excel Annotation Reference
---------------------------

**@Excel**

| Property | Type | Default | Description |
|----------------|----------|------------------|------------------------------------------------------------------------|
| name | String | null | Column name, supports name_id |
| needMerge | boolean | false | Whether to merge cells vertically (for single cells in a list, merge multiple rows created by list) |
| orderNum | String | "0" | Column order, supports name_id |
| replace | String[] | {} | Value replacement, export {a_id,b_id}, import reversed |
| savePath | String | "upload" | Import file save path, can be filled for images, default is upload/className/ |
| type | int | 1 | Export type: 1=text, 2=image, 3=function, 10=number, default is text |
| width | double | 10 | Column width |
| height | double | 10 | Column height (will be deprecated, use @ExcelTarget height instead) |
| isStatistics | boolean | false | Auto statistics, append a statistics row with all data summed |
| isHyperlink | boolean | FALSE | Hyperlink, need to implement interface to return object |
| isImportField | boolean | TRUE | Validate field exists in imported Excel, supports name_id |
| exportFormat | String | "" | Export date format |
| importFormat | String | "" | Import date format |
| format | String | "" | Date format, equivalent to setting both exportFormat and importFormat |
| databaseFormat | String | "yyyyMMddHHmmss" | Database format for string type date fields |
| numFormat | String | "" | Number format, Pattern parameter, uses DecimalFormat |
| imageType | int | 1 | Image type: 1=from file, 2=from database, default is file |
| suffix | String | "" | Text suffix, e.g., % makes 90 become 90% |
| isWrap | boolean | TRUE | Whether to wrap, supports \n |
| mergeRely | int[] | {} | Merge cell dependencies, e.g., {0} for second column based on first |
| mergeVertical | boolean | false | Vertically merge cells with same content |
| fixedIndex | int | -1 | Corresponds to Excel column, ignore name |
| isColumnHidden | boolean | FALSE | Export hidden column |

**@ExcelCollection**

| Property | Type | Default | Description |
|----------|----------|-----------------|------------------|
| id | String | null | Define ID |
| name | String | null | Define collection column name, supports name_id |
| orderNum | int | 0 | Order, supports name_id |
| type | Class<?> | ArrayList.class | Used to create objects during import |

**Single Table Export Entity Example**

```Java
public class SysUser implements Serializable {

    /** id */
    private String id;

    /** Login Account */
    @Excel(name = "Login Account", width = 15)
    private String username;

    /** Real Name */
    @Excel(name = "Real Name", width = 15)
    private String realname;

    /** Avatar */
    @Excel(name = "Avatar", width = 15)
    private String avatar;

    /** Birthday */
    @Excel(name = "Birthday", width = 15, format = "yyyy-MM-dd")
    private Date birthday;

    /** Gender (1: Male 2: Female) */
    @Excel(name = "Gender", width = 15, dicCode="sex")
    private Integer sex;

    /** Email */
    @Excel(name = "Email", width = 15)
    private String email;

    /** Phone */
    @Excel(name = "Phone", width = 15)
    private String phone;

    /** Status (1: Normal 2: Frozen) */
    @Excel(name = "Status", width = 15, replace={"Normal_1","Frozen_0"})
    private Integer status;
}
```

**One-to-Many Export Entity Example**

```Java
@Data
public class JeecgOrderMainPage {
    
    /** Primary Key */
    private java.lang.String id;
    
    /** Order Number */
    @Excel(name="Order Number", width=15)
    private java.lang.String orderCode;
    
    /** Order Type */
    private java.lang.String ctype;
    
    /** Order Date */
    @Excel(name="Order Date", width=15, format = "yyyy-MM-dd")
    private java.util.Date orderDate;
    
    /** Order Amount */
    @Excel(name="Order Amount", width=15)
    private java.lang.Double orderMoney;
    
    /** Order Note */
    private java.lang.String content;
    
    /** Created By */
    private java.lang.String createBy;
    
    /** Create Time */
    private java.util.Date createTime;
    
    /** Updated By */
    private java.lang.String updateBy;
    
    /** Update Time */
    private java.util.Date updateTime;
    
    @ExcelCollection(name="Customer")
    private List<JeecgOrderCustomer> jeecgOrderCustomerList;
    
    @ExcelCollection(name="Ticket")
    private List<JeecgOrderTicket> jeecgOrderTicketList;
}
```

---------------------------
Example Code
---------------------------
- [Example Code](./autopoi-spring-boot-2-starter/src/test/java/) - Unit test code location

---------------------------
License
---------------------------
Apache License 2.0

