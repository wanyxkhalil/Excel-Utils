# Excel 工具类

## 使用

- 使用JavaBean

```groovy
//def list = new ArrayList<{JavaBean with filed annotation}>()
// ...... 添加数据

def path = ExcelWriter.builder()
                .fileName()
                .sheet(list)
                .build()
```

- 使用动态数据，主要用于一维表难以表述的情况

```groovy
def list = new ArrayList<DynamicData>()
// ...... 添加数据

def path = ExcelDynamicWriter.builder()
                .fileName()
                .sheet(list)
                .build()
```

## 主要方法列表

1. fileName
   1. `fileName()` 不添加自定义文件前缀，使用自动生成的文件名。
   2. `fileName(String fileName)` 添加文件前缀。需为英文。
2. sheet
   1. `sheet(List data)` 只添加数据
   2. `sheet(String sheetName, List data)` 添加sheet名称、数据
   3. `sheet(String sheetName, Class clz, List data)` 添加sheet名称、数据的类信息、数据。将使用clz作为数据的Bean来解析