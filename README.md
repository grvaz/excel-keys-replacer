Use Excel files as templates. This package will replace special labels in Excel files (both xls and xlsx) matching array keys.  
Required: [phpoffice/phpexcel](https://github.com/PHPOffice/PHPExcel).  
Install via composer:  
composer require it-poet/excel-keys-replacer

Example:  
Insert anywhere in xls_keys.xls in any cells ${key1}$ and ${key2}$.  
```php
$excel = new \ItPoet\ExcelKeysReplacer\Replacer;  
$excel->file('path/to/file/xls_keys.xls');            
$excel->data([  
    'key1' => 'some text',  
    'key2' => '12345',  
]);  
$xls = $excel->replace();  
$writer = new \PHPExcel_Writer_Excel5($xls);  
$writer->save('path/to/file/xls_replaced.xls'); 
 ```        
Optionally (with default values):  
```php
$excel->keyWrapper('${', '}$');  
$excel->sheet(0);
```
TODO:
- multi-row insert to build tables (if key is array).
