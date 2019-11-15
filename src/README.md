# DTY PHP EXECL 增强扩展 

###使用方法

`composer require tdy/execl`

`use Tdy\Execl\Import;`

`$filePath="execl表格.xlsx";`

`$obj=  new Import($filePath);`

##SET方法

###设置列字段转小写
` $obj->column2lower=true; `

###设置读取最大列记录 eg：a到z
`$obj->setMax_column("z");`

###设置读取的最大行记录  eg:10  只读取10 行记录
`$obj->setMax_row(10);`

###设置读取指定的列字段
`$obj->setField("a,b,c,d,e,f,g,z"); `

###设置列字段别名
` $obj->setFieldAlias(['a'=>'field1',"b"=>"field2",'e'=>"field3"]);`

###设置表头所占行数
`$obj->setThead_row(2)`

##GET 方法

###获取隐藏的列字段
`$visible= $obj->getVisible();  `

###获取表头数据
`$thead= $obj->getThead();`

###获取数据  可选参数 $title,$hide_column,$therd
` $obj->getData();`








