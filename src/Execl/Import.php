<?php
/**
 * Created by PhpStorm.
 * Author: DoubleY
 * Date: 2019/11/12
 * Time: 10:22
 * Email: 731633799@qq.com
 */

namespace Tdy\Execl;

class Import
{

    public $PHPExcel = null;
    public $sheet_count = 0;
    private $thead_row = 1;
    private $max_column = null;
    private $max_row = null;
    private $hide_column = [];
    public $column2lower = true;
    private $thead_data = [];
    private $field = null;
    private $fieldAlias = null;

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


    /**获取execl隐藏的具体字段
     * @param null $currentSheet
     * @return array
     */
    public function getVisible($currentSheet = null)
    {
        $currentSheet = $currentSheet ? $currentSheet : $this->PHPExcel->getSheet(0);
        $maxColumn = !$this->max_column ? $currentSheet->getHighestDataColumn() : $this->max_column;
        for ($currentRow = 1; $currentRow < 2; $currentRow++) {
            if ($this->field) {
                for ($currentColumn = 0; $currentColumn < count($this->field); $currentColumn++) {
                    $colName = $this->number2column($currentColumn);
                    !$currentSheet->getColumnDimension($colName)->getVisible() && $this->hide_column[] = $this->aliasField($colName);
                }
            } else {
                for ($currentColumn = 0; $currentColumn < $this->column2number($maxColumn); $currentColumn++) {
                    $colName = $this->number2column($currentColumn);
                    !$currentSheet->getColumnDimension($colName)->getVisible() && $this->hide_column[] = $this->aliasField($colName);
                }
            }
        }
        return $this->hide_column;
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
     * @param null $currentSheet
     * @param string $callback 回调函数
     * @return array
     */
    private function getRowData($currentSheet = null, $callback = 'default_value')
    {
        $currentSheet = $currentSheet ? $currentSheet : $this->PHPExcel->getSheet(0);
        $maxColumn = !$this->max_column ? $currentSheet->getHighestDataColumn() : $this->max_column;      //获取A-X最大的列号
        $maxRow = !$this->max_row ? $currentSheet->getHighestRow() : $this->max_row;                      //取得一共有多少行数据
        $currentRow_int = 1;
        $this->thead_data && $currentRow_int = $this->thead_row + 1;
        $this->thead_data && $this->max_row && $maxRow = $maxRow + $this->thead_row;
        $this->thead_data && !$this->max_row && $maxRow = $maxRow - $this->thead_row;
        $all_data = [];
        for ($currentRow = $currentRow_int; $currentRow <= $maxRow; $currentRow++) {
            $arr = [];
            if ($this->field) {
                for ($currentColumn = 0; $currentColumn < count($this->field); $currentColumn++) {
                    $cell = $currentSheet->getCell(strtoupper($this->field[$currentColumn]) . $currentRow);
                    $arr[$this->aliasField($this->field[$currentColumn])] = (string)$this->$callback($cell);
                }
            } else {
                for ($currentColumn = 0; $currentColumn < $this->column2number($maxColumn); $currentColumn++) {
                    $colName = $this->number2column($currentColumn);
                    $cell = $currentSheet->getCell($colName . $currentRow);
                    $arr[$this->aliasField($colName)] = (string)$this->$callback($cell);;
                }
            }
            $all_data[] = $arr;
        }
        return $all_data;
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


    /**
     * 处理获取的execl 单元格内容
     * @param $cell
     * @return string
     */
    private function default_value($cell)
    {
        $val = $cell->getFormattedValue();
        if ($val instanceof \PHPExcel_RichText) {
            $val = $val->getPlainText();
        }
        return $val;
    }


    /**
     * 获取exexel 表格头部
     * @param null $currentSheet
     * @param string $callback
     * @return array
     */
    public function getThead($currentSheet = null, $callback = "default_value")
    {
        $currentSheet = $currentSheet ? $currentSheet : $this->PHPExcel->getSheet(0);
        $all_data = [];
        for ($currentRow = 1; $currentRow <= $this->thead_row; $currentRow++) {
            $arr = [];
            if ($this->field) {
                for ($currentColumn = 0; $currentColumn < count($this->field); $currentColumn++) {
                    $cell = $currentSheet->getCell(strtoupper($this->field[$currentColumn]) . $currentRow);
                    $arr[$this->aliasField($this->field[$currentColumn])] = $this->$callback($cell);
                }
            } else {
                $maxColumn = !$this->max_column ? $currentSheet->getHighestDataColumn() : $this->max_column;
                !is_numeric($maxColumn) && $maxColumn = $this->column2number($maxColumn);
                for ($currentColumn = 0; $currentColumn < $maxColumn; $currentColumn++) {
                    $colName = $this->number2column($currentColumn);
                    $cell = $currentSheet->getCell($colName . $currentRow);
                    $arr[$this->aliasField($colName)] = $this->$callback($cell);;
                }
            }
            $all_data[] = $arr;
        }
        $this->thead_data = $all_data;
        return $all_data;
    }

    /**
     * 获取全部execl   配合 getThead 方法返回值会有所不同
     * @param bool $title
     * @param bool $hide_column
     * @param bool $therd
     * @return array
     */
    public function getData($title = false, $hide_column = false, $therd = false)
    {
        $sheet = [];
        for ($sheet_count = 0; $sheet_count < $this->sheet_count; $sheet_count++) {
            $currentSheet = $this->PHPExcel->getSheet($sheet_count);
            $all_data = $this->getRowData($currentSheet);
            $title && $data['title'] = $currentSheet->getTitle();
            $hide_column && $data['hide_column'] = $this->getVisible($currentSheet);
            $therd && $data['therd'] = $this->getThead($currentSheet);
            $data = ['data' => $all_data];
            $sheet[] = $data;
        }
        return $sheet;
    }
}
