# pdfScanner
## 功能
- 扫描pdf文件中的单词高亮, 按照不同的高亮颜色进行分类, 支持文件夹批量扫描
- 本地词典+联网翻译单词
- 将结果按照pdf文件名保存为一个个Excel worksheet, 并进行排序
- 生成网易有道词典单词本xml
- 勾选“单词本”选项即可自动添加单词到网易有道词典单词本

## 生成产物
- output.xlsx 保存了所有扫描的单词和其对应翻译，并进行了归类
- youdao_wordbook.xml 生成的有道词典单词本的xml, 可以直接导入到有道词典单词本（功能迭代的产物, 其实以及可以自动添加到单词本了）

## 说明
- 生成的产物都保存在pdf同级目录下
- 本地翻译需要加载词典, 可以使用同级目录下的两个英汉大词典, 注意修改代码中的路径
- 需要事先获取有道词典的cookie
- 需要事先获取百度翻译api的key和secret

## 效果图
<img width="280" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/b7afac88-8e17-4bb7-888a-818b922a86d8">               <img width="280" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/b828f7a4-9a00-4b2a-baa3-a4e49cb1e3b8">
<img width="600" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/59ea605c-55d1-4293-aa89-9dbc73b843c8">

## 打赏
<img width="196" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/76a3ab13-23f7-45f3-b978-ec4ac580e140">
