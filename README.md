# MDword 
## [中文文档](https://github.com/mkdreams/MDword/tree/master/README-CN.md)

##  Project General Name
Template: a word which will be revised.  
Block: the part which will be replaced or cloned.

## Project Introduction
Main Use: to generate word dynamically.  
Advantage: focus only on dynamic data and logic, rather than adjusting the style, which can modulate the template with the help of office word.

## Comparison between MDword & PHPword
+ ### Similarities
1. PHP Package.
2. Both can be used to generate office word.

+ ### Differences
1. PHPword concentrates on writing elements one by one. However,
it is more powerful and efficient for Mdword just to revise  them on the template.
2. For updating text styles, adding covers, headers and footers, MDword just modifies the template by office word, while PHPword complicate the task-adjusting every element.
3. Directories(Table of content) can be automatically generated.

## Tutotials
+ ### Installation
```
// Method 1
composer require mkdreams/mdword
// Method 2, Autoloading Template
require_once('Autoloader.php');
```

+ ### Add annotations or use “${value/}” to the template. Please note that there is a “/” at the end.
![image](https://user-images.githubusercontent.com/12422458/111026036-1c647700-8423-11eb-9df2-e9a2e5530007.png) 
+ ### Invocation Methods (more and richer approaches, for example: [tests\samples\simple for readme](https://github.com/mkdreams/MDword/blob/master/tests/samples/simple%20for%20readme/index.php), such as catalog, sequence number, etc.)
```
// New class,load template
$TemplateProcessor = new WordProcessor();
$template = 'temple.docx';
$TemplateProcessor->load($template);

// Set Value
$TemplateProcessor->setValue('value', 'r-value');

// Clone 
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

// set value for image
$TemplateProcessor->setImageValue('image', dirname(__FILE__).'/logo.jpg');

// Delete a paragraph
$TemplateProcessor->deleteP('style');

// Save
$rtemplate = __DIR__.'/r-temple.docx';
$TemplateProcessor->saveAs($rtemplate);
```

+ ### Result
![image](https://user-images.githubusercontent.com/12422458/111026037-1d95a400-8423-11eb-81e2-941f6b854e34.png) 

+ ### GIFs
![image](https://user-images.githubusercontent.com/12422458/111026041-1ec6d100-8423-11eb-8e14-d8daf99a9704.gif) 


## More samples
- [Simple Comprehhensive cases](https://github.com/mkdreams/MDword/tree/master/tests/samples/simple%20for%20readme)   

- [Formatted texts](https://github.com/mkdreams/MDword/tree/master/tests/samples/text)

- [Add images](https://github.com/mkdreams/MDword/tree/master/tests/samples/image)

- [Clone](https://github.com/mkdreams/MDword/tree/master/tests/samples/clone)

- [Many ways to set blocks, solving the lack of notes](https://github.com/mkdreams/MDword/tree/master/tests/samples/block)

- [Write element that is written by PHPword to the block](https://github.com/mkdreams/MDword/tree/master/tests/samples/phpword)

- [Put TOC in a table](https://github.com/mkdreams/MDword/blob/main/tests/samples/toc)

## Communication
### Note: Exchange idea on MDword.
![image](https://user-images.githubusercontent.com/12422458/111025926-5a14d000-8422-11eb-86a3-db8a0ad712f0.png) 

## [Project Plans](https://github.com/mkdreams/MDword/projects/1#column-10318470)


