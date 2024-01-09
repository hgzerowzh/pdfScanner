# pdfScanner
## 功能选项
- 扫描pdf文件中的单词高亮, 按照不同的高亮颜色进行分类, 支持文件夹批量扫描
- 翻译单词: 本地词典+联网翻译单词
- 添加单词本: 自动添加单词到网易有道词典单词本

## 生成文件
- 单词Excel: 将扫描的结果归类整理到Excel, 并按照pdf文件名保存为一个个Excel worksheet, 并进行排序(output.xlsx)
- 单词本xml: 生成网易有道词典单词本xml文件(youdao_wordbook.xml),（功能迭代的产物, 其实已经可以自动添加到单词本了）

## 说明
- 功能选项和生成文件的选项之间存在依赖关系
- 生成的产物都保存在pdf同级目录下
- 本地翻译需要加载词典, 可以使用同级目录下的两个英汉大词典, 注意修改代码中的路径
- 需要事先获取有道词典的cookie
- 需要事先获取百度翻译api的key和secret

## 效果图
<img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/54fef90b-a4d7-4061-b44a-d8e1d99f6150">
   <img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/1d68322d-7b2d-4d19-ad69-d41494c3c08a">  <img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/9b4d442f-4159-49ca-b071-c9d4c67a1634">  <img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/c0d8ff81-0624-4315-b298-5458fd447a71">




<img width="600" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/59ea605c-55d1-4293-aa89-9dbc73b843c8">

## 打赏
<img width="196" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/76a3ab13-23f7-45f3-b978-ec4ac580e140">
