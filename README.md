# MDword

## 项目通用名称
母版：在某个word基础上修改的,这个word命名为“母版”   
区块：也就是要被替换或者克隆的部分呢


## 项目介绍
主要用途：动态生成word  
优势：生成word只需关注动态数据及逻辑，无需关注式样的调整（式样可以借助office word调整母版即可）

## 与PHPWord的爱恨情仇
+ ### 共同点
1. PHP编写的库（资源包）
2. 用于生成office word

+ ### 不同点
1. PHPWord 需要一个元素一个元素的写入，而MDword则是在母版的基础上修改，编码效率更高
2. 修改文字式样，增加封面，修改页眉页脚MDword只需用word编辑软件调整母版，而PHPWord需要繁琐的去调整每个元素

## 教程
+ ### 安装
```
//方法一
composer require mkdreams/mdword
//方法二，手动引入自动加载类
require_once('Autoloader.php');
```
+ ### 给母版“temple.docx”添加批注(或采用“${value/}”方式，注意结尾有个“/”)
![image](https://user-images.githubusercontent.com/12422458/111026036-1c647700-8423-11eb-9df2-e9a2e5530007.png) 
+ ### 调用方法（更多更丰富的调用方式，参考案例：[tests\samples\simple for readme](https://github.com/mkdreams/MDword/blob/master/tests/samples/simple%20for%20readme/index.php)，例如：目录、序号等）
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
$TemplateProcessor->setValue('name#1', [
    ['text'=>'colin1','style'=>'style','type'=>MDWORD_TEXT],
    ['text'=>1,'type'=>MDWORD_BREAK],
    ['text'=>'86','style'=>'style','type'=>MDWORD_TEXT]
]);
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
![image](https://user-images.githubusercontent.com/12422458/111026037-1d95a400-8423-11eb-81e2-941f6b854e34.png) 

+ ### 动图
![image](https://user-images.githubusercontent.com/12422458/111026041-1ec6d100-8423-11eb-8e14-d8daf99a9704.gif) 


## 更多案例
- [简单的综合案例](https://github.com/mkdreams/MDword/tree/master/tests/samples/simple%20for%20readme)
- [带式样的文字](https://github.com/mkdreams/MDword/tree/master/tests/samples/text)
- [添加图片](https://github.com/mkdreams/MDword/tree/master/tests/samples/image)
- [克隆](https://github.com/mkdreams/MDword/tree/master/tests/samples/clone)
- [多种方式设置区块，解决无法添加批注问题](https://github.com/mkdreams/MDword/tree/master/tests/samples/block)
- [PHPWORD写入到区块](https://github.com/mkdreams/MDword/tree/master/tests/samples/phpword)

## 交流
###  请备注：MDword交流
![image](https://user-images.githubusercontent.com/12422458/111025926-5a14d000-8422-11eb-86a3-db8a0ad712f0.png) 


## [项目计划](https://github.com/mkdreams/MDword/projects/1#column-10318470)

