# MDword
## 项目介绍
当前版本：Alpha  
主要用途：动态生成word  
优势：生成word只需关注动态数据及逻辑，无需关注式样的调整（式样可以借助office word调整母版即可）

## 使用方法
```
//引入自动加载类
require_once('Autoloader.php');
//新建类 加载 母版
$TemplateProcessor = new WordProcessor();
$template = 'temple.docx';
$TemplateProcessor->load($template);

//赋值

//图片替换

//克隆多列

//克隆并复制


```
## 项目进展
- [x] 目录随内容更改而更新 （2020/04/09 完成）
- [ ] readme 完善
- [ ] tests框架完善
- [x] 赋值支持克隆属性（字体属性） （2020/04/16 完成）
- [ ] 支持XML赋值
- [x] 删除区块  （2020/04/16 完成）
- [x] 支持换行符&分页符
- [x] 克隆支持段落（cloneP） 
- [x] 速度优化 (第一轮)
- [x] 添加bind数据类（简化克隆和赋值操作）
- [ ] 赋值支持图片
- [ ] API
- [x] 目录更新优化 （2020/07/20 完成）
- [x] 优化速度 （2020/07/20 完成）
- [ ] 与echarts对接，支持动态绘图(取代word中自带的绘图)

## 名称介绍
母版：在某个word基础上修改的,这个word命名为“母版”
