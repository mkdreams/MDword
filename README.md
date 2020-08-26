# MDword
中文 | English

## 项目通用名称
母版：在某个word基础上修改的,这个word命名为“母版”   
block：区块，也就是要被替换或者克隆的部分呢


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
+ ### 给母版“temple.docx”添加批注
![image](https://github.com/mkdreams/MDword/blob/master/tests/samples/simple%20for%20readme/word%20config.png) 
+ ### 调用方法（可参考此实例：[tests\samples\simple for readme](https://github.com/mkdreams/MDword/blob/master/tests/samples/simple%20for%20readme/index.php)）
```
//新建类 加载 母版
$TemplateProcessor = new WordProcessor();
$template = 'temple.docx';
$TemplateProcessor->load($template);

//赋值
$TemplateProcessor->setValue('value', 'r-value');

//克隆并复制
$TemplateProcessor->clones('people', 3);

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

+ ### 结果
![image](https://github.com/mkdreams/MDword/blob/master/tests/samples/simple%20for%20readme/word%20result.png) 

## 更多案例
- [简单的综合案例](https://github.com/mkdreams/MDword/tree/master/tests/samples/simple%20for%20readme)
- [带式样的文字](https://github.com/mkdreams/MDword/tree/master/tests/samples/text)
- [添加图片](https://github.com/mkdreams/MDword/tree/master/tests/samples/image)
- [克隆](https://github.com/mkdreams/MDword/tree/master/tests/samples/clone)
- [多种方式设置block，解决无法添加批注问题](https://github.com/mkdreams/MDword/tree/master/tests/samples/block)

## [项目进展](https://github.com/mkdreams/MDword/projects/1#column-10318470)

