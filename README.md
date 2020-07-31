# MDword
## 项目介绍
当前版本：Alpha  
主要用途：动态生成word  
优势：生成word只需关注动态数据及逻辑，无需关注式样的调整（式样可以借助office word调整母版即可）

## 教程
+ ### 安装
```
//方法一
composer require mkdreams/mdword
//方法二，手动引入自动加载类
require_once('Autoloader.php');
```
+ ### 使用方法（可参考此实例：tests\samples\simple for readme）
```
//新建类 加载 母版
$TemplateProcessor = new WordProcessor();
$template = 'temple.docx';
$TemplateProcessor->load($template);

//赋值
$TemplateProcessor->setValue('value', 'r-value');

//克隆并复制
$TemplateProcessor->clone('people', 3);

$TemplateProcessor->setValue('name#0', 'colin0');
$TemplateProcessor->setValue('name#1', [['text'=>'colin1','style'=>'style','type'=>MDWORD_TEXT]]);
$TemplateProcessor->setValue('name#2', 'colin2');

$TemplateProcessor->setValue('sex#1', 'woman');

$TemplateProcessor->setValue('age#0', '280');
$TemplateProcessor->setValue('age#1', '281');
$TemplateProcessor->setValue('age#2', '282');

//图片复制
$TemplateProcessor->setImageValue('image', dirname(__FILE__).'/logo.jpg');

//删除某行
$TemplateProcessor->deleteP('style');

//保存
$rtemplate = __DIR__.'/r-temple.docx';
$TemplateProcessor->saveAs($rtemplate);
```
## 项目进展
- [x] 目录随内容更改而更新 （2020/04/09 完成）
- [ ] readme 完善
- [ ] tests框架完善
- [x] 赋值支持克隆属性（字体属性） （2020/04/16 完成）
- [x] 支持XML赋值 （2020/07/29 完成）
- [x] 删除区块  （2020/04/16 完成）
- [x] 支持换行符&分页符
- [x] 克隆支持段落（cloneP） 
- [x] 速度优化 (第一轮)
- [x] 添加bind数据类（简化克隆和赋值操作）
- [x] 赋值支持图片 （2020/07/29 完成）
- [ ] API
- [x] 目录更新优化 （2020/07/20 完成）
- [x] 优化速度  (第二轮)（2020/07/20 完成）
- [ ] 与echarts对接，支持动态绘图(取代word中自带的绘图)
- [x] 删除分页前换行
- [ ] 图片支持式样控制，长&宽

## 名称介绍
母版：在某个word基础上修改的,这个word命名为“母版”
