## Excelize简介
Excelize 是 Go 语言编写的一个用来操作 Office Excel 文档类库，基于 ECMA-376 Office OpenXML 标准。可以使用它来读取、写入 XLSX 文件。

### 引入excelize库
``` golang
import (
	"github.com/360EntSecGroup-Skylar/excelize"
)
```

### 创建excel
使用 NewFile 新建 Excel 工作薄，新创建的工作簿中会默认包含一个名为 Sheet1 的工作表。
``` golang
func NewFile() *File
```

### 打开excel
``` golang
func OpenFile(filename string) (*File, error)
```

### 保存 / 另存为
``` golang
func (f *File) Save() error
func (f *File) SaveAs(name string) error
```

## 工作表(sheet)的操作
- 新建工作表
``` golang
func (f *File) NewSheet(name string) int
```
根据给定的工作表名称添加新的工作表，并返回工作表索引。如果添加的工作表名已经存在, 则直接返回工作表的索引

- 修改工作表名称
``` golang
func (f *File) SetSheetName(oldName, newName string)
```

- 列出所有的工作表的名称及索引
``` golang
func (f *File) GetSheetMap() map[int]string
```

## 行 / 列的操作
- 插入 / 删除行
``` golang
func (f *File) InsertRow(sheet string, row int) error
func (f *File) RemoveRow(sheet string, row int) error
```
根据给定的工作表名称（大小写敏感）和行号，在指定行之后插入空白行。

- 插入 / 删除列
``` golang
func (f *File) InsertCol(sheet, column string) error
func (f *File) RemoveCol(sheet, column string) error
```

- 按行赋值
``` golang
func (f *File) SetSheetRow(sheet, axis string, slice interface{}) error
```

## 单元格的操作
- 设置单元格的值
``` golang
func (f *File) SetCellValue(sheet, axis string, value interface{}) error
```

- 获取指定单元格的值
``` golang
func (f *File) GetCellValue(sheet, axis string) (string, error)
```

- 获取全部单元格的值
``` golang
func (f *File) GetRows(sheet string) ([][]string, error)
```

## 样式
``` golang
// 创建样式
func (f *File) NewStyle(style interface{}) (int, error)
// 把样式应用到指定的单元格上
func (f *File) SetCellStyle(sheet, hcell, vcell string, styleID int) 
// 把样式应用到列上
func (f *File) SetColStyle(sheet, columns string, styleID int) error
```

- 填充背景色
``` golang
style, err := xlsx.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
if err != nil {
    fmt.Println(err)
}
xlsx.SetCellStyle("Sheet1", "H9", "H9", style)
```

- 对齐方式
``` golang
style, err := f.NewStyle(`{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"","wrap_text":true}}`)
if err != nil {
	fmt.Println(err)
}
f.SetCellStyle("Sheet1", "H9", "H9", style)
```