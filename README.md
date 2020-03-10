# DTY PHP EXECL 增强扩展 

### 使用方法

`composer require tdy/execl`

`use Tdy\Execl\Import;`

`$filePath="execl表格.xlsx";`

     `$obj= new Import($filePath);  //实例对象
      $data=$obj->setField("a,b,c,d,e,f,g,h,i,j,k") //获取指定单元格
      ->setImagePath(true)  //设置获取图片资源路径
      ->setThead_row(1)     //指定表头行数
      ->setMax_row(2)       //指定读取数据行数
      ->setFieldAlias(['a'=>'name',"b"=>"description",'c'=>"meta_title","d"=>'price']) //字段别名
      ->getData();`  //获取的数据


## SET方法


### 设置读取最大列记录 eg：a到z
`$obj->setMax_column("z");`

### 设置读取的最大行记录  eg:10  只读取10 行记录
`$obj->setMax_row(10);`

### 设置读取指定的列字段
`$obj->setField("a,b,c,d,e,f,g,z"); `

### 设置列字段别名
` $obj->setFieldAlias(['a'=>'field1',"b"=>"field2",'e'=>"field3"]);`

### 设置表头所占行数
`$obj->setThead_row(2)`

### 设置execl图片路径 参数为true：获取资源路径  参数为具体路径时：返回本地保存后的路径 
`$obj->setImagePath(true)`

## GET 方法

### 获取隐藏的列字段
`$obj->hide_column;  `

### 获取表头数据
`$obj->$thead_data;`

### 获取图片资源
`$obj->imageData`

### 获取数据  可选参数 $title,$therd,$hide_column 
` $obj->getData();`


```` Array
     (
         [0] => Array
             (
                 [0] => Array
                     (
                         [name] => 产品名称
                         [description] => 同轮播图
                         [meta_title] => 同标题
                         [price] => 5.5
                         [category] => BEACH CANVAS BAG
                         [type] => 82% polyester 18% spandex
                         [pic] => Array
                             (
                                 [0] => Array
                                     (
                                         [file] => zip://E:\wamp\www\CMS\opencart\image\catalog\execl\Book1.xlsx#xl/media/image1.jpeg
                                         [title] => 主图
                                     )
     
                                 [1] => Array
                                     (
                                         [file] => zip://E:\wamp\www\CMS\opencart\image\catalog\execl\Book1.xlsx#xl/media/image2.jpeg
                                         [title] => 
                                     )
     
                             )
     
                         [banner] => Array
                             (
                                 [0] => Array
                                     (
                                         [file] => zip://E:\wamp\www\CMS\opencart\image\catalog\execl\Book1.xlsx#xl/media/image2.jpeg
                                         [title] => 图片xxxxxxx
                                     )
     
                                 [1] => Array
                                     (
                                         [file] => zip://E:\wamp\www\CMS\opencart\image\catalog\execl\Book1.xlsx#xl/media/image1.jpeg
                                         [title] => 图片二
                                     )
     
                             )
     
                         [seo_title] => 同标题
                         [color] => Blue,green,white,yellow
                         [size] => S,M,L
                     )
     
             )
     
     )




