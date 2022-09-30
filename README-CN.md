# MDword
## [English document](https://github.com/mkdreams/MDword/blob/main/README.md)

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
1. PHPWord 专注于一个元素一个元素的写入，而MDword则是专注于在母版的基础上修改，功能更强大，编码效率更高
2. 修改文字式样，增加封面，修改页眉页脚MDword只需用word编辑软件调整母版，而PHPWord需要繁琐的去调整每个元素
3. 可以自动生成目录

## 教程
+ ### 安装
```
//方法一
composer require mkdreams/mdword
//方法二，手动引入自动加载类
require_once('Autoloader.php');
```
+ ### 给母版“temple.docx”添加批注，或者添加“${value/}”这类特殊文字，注意结尾有个“/”。
![image](https://user-images.githubusercontent.com/12422458/111026036-1c647700-8423-11eb-9df2-e9a2e5530007.png) 
+ ### 调用方法（更多更丰富的调用方式，参考案例：[tests\samples\simple for readme](https://github.com/mkdreams/MDword/blob/main/tests/samples/simple%20for%20readme/index.php)，例如：目录、序号等）
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

## 性能情况（[统计脚本](https://github.com/mkdreams/MDword/blob/main/tests/samples/performance/index.php)）
|  测试项   |  用时(S)   |
|  ----   |  ----   |
|  1页母版赋值100次   |  0.04   |
|  1页母版赋值500次   |  0.16   |
|  1页母版赋值1000次   |  0.33   |
|  1页母版赋值10000次   |  7.80   |
|  1750页母版赋值100次   |  4.61   |
|  1750页母版赋值500次   |  4.94   |
|  1750页母版赋值1000次   |  5.43   |
|  1750页母版赋值10000次   |  17.39   |

## 内存使用情况（[统计脚本](https://github.com/mkdreams/MDword/blob/main/tests/samples/memory%20use/index.php)）
|  连续运行第几次   | 累积内存使用情况 |  备注 |
|  ----  | ----  | ----  |
| 1  | 0.050590515136719 M | 首次需要加载PHP类 |
| 2  | 0.050949096679688 M |  |
| 3  | 0.050949096679688 M |  |
| 4  | 0.050949096679688 M |  |
| 5  | 0.050949096679688 M |  |
| 6  | 0.050949096679688 M |  |
| 7  | 0.050949096679688 M |  |
| 8  | 0.050949096679688 M |  |


## 更多案例
- [简单的综合案例](https://github.com/mkdreams/MDword/tree/main/tests/samples/simple%20for%20readme)
- [带式样的文字](https://github.com/mkdreams/MDword/tree/main/tests/samples/text)
- [添加图片](https://github.com/mkdreams/MDword/tree/main/tests/samples/image)
- [克隆](https://github.com/mkdreams/MDword/tree/main/tests/samples/clone)
- [多种方式设置区块，解决无法添加批注问题](https://github.com/mkdreams/MDword/tree/main/tests/samples/block)
- [PHPWORD写入到区块](https://github.com/mkdreams/MDword/tree/main/tests/samples/phpword)
- [目录嵌入到表格](https://github.com/mkdreams/MDword/tree/main/tests/samples/toc)
- [合并表格单元格](https://github.com/mkdreams/MDword/tree/main/tests/samples/merge%20table%20cells)



## 交流
###  请备注：MDword交流
![image](https://user-images.githubusercontent.com/12422458/111025926-5a14d000-8422-11eb-86a3-db8a0ad712f0.png) 


## [项目计划](https://github.com/mkdreams/MDword/projects/1#column-10318470)

