<?php
/**
 * Import class
 * Created by PhpStorm.
 * Author: DoubleY
 * Date: 2019/11/12
 * Time: 10:22
 * Email: 731633799@qq.com
 * PHPExecl 导入自定义处理
 */
namespace Tdy\Execl;

class Import
{

    private $PHPExcel = null;
    private $sheet_count = 0;
    private $column2lower = true;
    private $thead_row = 0;
    private $max_column = null;
    private $max_row = null;
    private $hide_column = [];
    private $thead_data = [];
    private $field = null;
    private $fieldAlias = null;
    private $imageData = [];
    private $imagePath = false;

    /**
     * PhpExecl 导入处理
     * import constructor.
     * @param $file
     * @throws \Exception
     */
    public function __construct($file)
    {
        if (!file_exists($file)) {
            throw new \Exception("文件不存在");
        }
        $PHPReader = new \PHPExcel_Reader_Excel2007();
        if (!$PHPReader->canRead($file)) {
            $PHPReader = new \PHPExcel_Reader_Excel5();
            if (!$PHPReader->canRead($file)) {
                $PHPReader = new \PHPExcel_Reader_CSV();
                if (!$PHPReader->canRead($file)) {
                    throw new \Exception(__('未知的数据格式'));
                }
            }
        }
        $this->PHPExcel = $PHPReader->load($file);
        $this->sheet_count = $this->PHPExcel->getSheetCount();
    }


    /**
     * 设置读取的最大单元格列
     * @param string $max_column a-z......
     * @return $this
     */
    public function setMax_column($max_column)
    {
        $max_column && $this->max_column = $max_column;
        return $this;
    }

    /**
     * 设置读取最大的行数   默认自动获取
     * @param int $max_row 1-1000......
     * @return $this
     */
    public function setMax_row($max_row)
    {
        $max_row && $this->max_row = $max_row;
        return $this;
    }


    /**
     * 设置表格头部行数
     * @param int $thead_row
     * @return $this
     */
    public function setThead_row($thead_row)
    {
        $thead_row && $this->thead_row = $thead_row;
        return $this;
    }

    /**设置读取指定单元格列字段
     * @param null|string|array $field eg:'a,c,d'
     * @return $this
     */
    public function setField($field)
    {
        $field && $this->field = is_array($field) ? $field : explode(",", $field);
        return $this;
    }

    /**
     * 设置单元格别名 eg:['a'=>'field1','b'=>'field2']
     * @param array $arr
     * @return $this
     */
    public function setFieldAlias(array $arr)
    {
        $arr && $this->fieldAlias = $arr;
        return $this;
    }

    /**
     * 设置图片本地保存路径
     * @param String $path
     * @return $this
     */
    public function setImagePath($path)
    {
        if (is_bool($path) || is_string($path)) {
            $path && $this->imagePath = $path;
        }
        return $this;
    }

    /**
     * @param $column
     * @return int
     */
    protected function column2number($column)
    {
        return \PHPExcel_Cell::columnIndexFromString($column);
    }

    /**
     * @param $number
     * @return string
     */
    protected function number2column($number)
    {
        return \PHPExcel_Cell::stringFromColumnIndex($number);
    }


    /**
     * 设置键名
     * @param $colName
     * @return string
     */
    private function aliasField($colName)
    {
        if ($this->fieldAlias) {
            $temp_arr = $this->column2lower ? array_change_key_case($this->fieldAlias, CASE_LOWER) : $this->fieldAlias;
            $this->column2lower && $colName = strtolower($colName);
            if (array_key_exists($colName, $temp_arr) && isset($temp_arr[$colName])) {
                $colName = $temp_arr[$colName];
            }
        }
        return $this->column2lower ? strtolower($colName) : $colName;
    }

    /*处理图片
    *@param String $savePath  保存路径  为空时返回资源路径
    *@return array
    */
    private function handleImage($sheet_index = 0, $savePath = null)
    {
        $imageData = [];
        $currentSheet = $this->PHPExcel->getSheet($sheet_index);
        foreach ($currentSheet->getDrawingCollection() as $k => $image) {
            $codata = $image->getCoordinates();
            $temp_file = $image->getPath();
            if ($savePath) {
                @mkdir($savePath, 0777, true);
                $imgFile = $savePath . md5_file($temp_file) . '.' . $image->getExtension();
                if (!file_exists($imgFile)) {
                    copy($temp_file, $imgFile);
                }
            } else {
                $imgFile = $temp_file;
            }
            $imageData[$codata][] = $imgFile;
        }
        $this->imageData[$sheet_index] = $imageData;
        return $imageData;
    }

    /** 获取execl隐藏的具体字段
     * @param int $sheet_index
     * @return array
     */
    private function handleVisible($sheet_index = 0)
    {
        $currentSheet = $this->PHPExcel->getSheet($sheet_index);
        $arr = [];
        if ($this->field) {      //只读取每行指定字段
            for ($currentColumn = 0; $currentColumn < count($this->field); $currentColumn++) {
                $colName = strtoupper($this->field[$currentColumn]);
                !$currentSheet->getColumnDimension($colName)->getVisible() && $arr[] = $this->aliasField($this->field[$currentColumn]);
            }
        } else {
            for ($currentColumn = 0; $currentColumn < 1; $currentColumn++) {
                $colName = $this->number2column($currentColumn);
                !$currentSheet->getColumnDimension($colName)->getVisible() && $arr[] = $this->aliasField($this->field[$currentColumn]);
            }
        }
        $this->hide_column[$sheet_index] = $arr;
        return $arr;
    }

    /**
     * 处理当前表格数据
     * @param null $currentSheet index or 当前对象
     * @param string $callback 回调函数
     * @return array
     */
    private function handelRowData($sheet_index = 0)
    {
        $currentSheet = $this->PHPExcel->getSheet($sheet_index);
        $maxRow = $this->max_row ? $this->max_row + $this->thead_row : $currentSheet->getHighestRow();  //取得一共有多少行数据
        if (is_bool($this->imagePath)) {                                      //图片处理
            $this->imagePath == true && $this->handleImage($sheet_index);
        } else {
            $this->imagePath && $this->handleImage($sheet_index, $this->imagePath);
        }
        $all_data = [];
        for ($currentRow = 1; $currentRow <= $maxRow; $currentRow++) {         //每行
            $arr = $this->handelColData($sheet_index, $currentRow);            //获取每一列具体内容
            if ($this->thead_row >= 1 && $currentRow <= $this->thead_row) {    //设置表头数据
                $temp = &$this->thead_data;
                $temp[$sheet_index][] = $arr;
                continue;
            }
            $all_data[] = $arr;
        }
        return $all_data;
    }

    /*处理每列具体字段的内容
    *@param int $sheet_index   表格索引
    *@param int $currentRow    指定的行号
    *@param function $callback 回调函数
    */
    private function handelColData($sheet_index, $currentRow, $callback = 'default_value')
    {
        $currentSheet = $this->PHPExcel->getSheet($sheet_index);
        $arr = [];
        if ($this->field) {//只读取每行指定字段
            for ($currentColumn = 0; $currentColumn < count($this->field); $currentColumn++) {
                $cell = $currentSheet->getCell(strtoupper($this->field[$currentColumn]) . $currentRow);
                $arr[$this->aliasField($this->field[$currentColumn])] =  $this->$callback($cell, $sheet_index);
            }
        } else {
            $maxColumn = !$this->max_column ? $currentSheet->getHighestDataColumn() : $this->max_column; //获取A-X最大的列号
            for ($currentColumn = 0; $currentColumn < $this->column2number($maxColumn); $currentColumn++) {
                $colName = $this->number2column($currentColumn);
                $cell = $currentSheet->getCell($colName . $currentRow);
                $arr[$this->aliasField($colName)] = $this->$callback($cell, $sheet_index);
            }
        }
        return $arr;
    }

    /**
     * 处理获取的Execl 单元格内容
     * @param $cell
     * @return
     */
    private function default_value($cell, $sheet_index = 0)
    {
        $row_content = $cell->getFormattedValue();
        if ($row_content instanceof \PHPExcel_RichText) {
            $row_content = (string)$row_content->getPlainText();
        }
        if (isset($this->imageData[$sheet_index][$cell->getCoordinate()])) {
            $arr_img = $this->imageData[$sheet_index][$cell->getCoordinate()];
            $content_array = array_pad(stristr($row_content, ",") ? explode(",", $row_content) : explode("，", $row_content), count($arr_img), null);
            $row_content = array_map(function ($file, $title) {
                return compact('file', 'title');
            }, $arr_img, $content_array);
        }
        return $row_content;
    }

    /**
     * 获取全部execl数据   配合 setThead_row方法,返回值会有所不同
     * @param bool $title
     * @param bool $therd
     * @param bool $hide_column
     * @return array
     */
    public function getData($title = false, $therd = false, $hide_column = false)
    {
        $sheet = [];
        for ($sheet_count = 0; $sheet_count < $this->sheet_count; $sheet_count++) {
            $currentSheet = $this->PHPExcel->getSheet($sheet_count);
            $all_data = $this->handelRowData($sheet_count);
            $merge = [];
            $title && $merge['title'] = $currentSheet->getTitle();
            $hide_column && $merge['hide_column'] = $this->handleVisible($sheet_count);
            $therd && $merge['therd'] = $this->thead_data[$sheet_count];
            $data = (!$title && !$therd && !$hide_column) ? $all_data : array_merge($merge, ['data' => $all_data]);
            $sheet[] = $data;
        }
        return $sheet;
    }

    public function __get($name)
    {
        if(isset($this->$name)){
            return $this->$name;
        }else{
            return null;
        }
    }


}
