<?php

namespace common\components\excel;

use yii\base\Component;
use yii\base\InvalidConfigException;
use yii\helpers\ArrayHelper;

class ExcelImport extends Component
{
    /**
     * 文件地址
     *
     * @var string
     */
    public $file_name;

    /**
     * 文件类型
     *
     * @var string
     */
    public $format;

    /**
     * 选取列
     *
     * @var array
     */
    public $columns = [];

    /**
     * 是否设置首行为列名
     *
     * @var bool
     */
    public $first_record_as_key = true;

    /**
     * @var \PHPExcel
     */
    protected $_php_excel;

    /**
     * @inheritdoc
     */
    public function init()
    {
        if (!file_exists($this->file_name)) {
            throw new InvalidConfigException('input file not exist');
        }
        if (!$this->format) {
            $this->format = \PHPExcel_IOFactory::identify($this->file_name);
        }
        $this->_php_excel = \PHPExcel_IOFactory::createReader($this->format)->load($this->file_name);
    }

    /**
     * 输出数组
     *
     * @return array
     */
    public function toArray()
    {
        $sheet_data = $this->_php_excel->getActiveSheet()->toArray(null, true, true, true);
        if ($this->first_record_as_key) {
            $sheet_data = $this->setFirstRecordAsKey($sheet_data);
        }
        if (!empty($this->columns)) {
            $sheet_data = $this->setColumnsLabel($sheet_data);
        }
        return $sheet_data;
    }

    /**
     * 设置首行为列名
     *
     * @param array $sheet_data
     * @return array
     */
    protected function setFirstRecordAsKey($sheet_data)
    {
        $keys = array_shift($sheet_data);
        $new_data = [];
        foreach ($sheet_data as $values)
        {
            if (empty(array_filter($values))) {
                continue;
            }
            $new_data[] = array_combine($keys, $values);
        }
        return $new_data;
    }

    /**
     * 选取指定列并更改列列名
     *
     * @param array $sheet_data
     * @return array
     */
    protected function setColumnsLabel($sheet_data)
    {
        $new_data = [];
        if (ArrayHelper::isIndexed($this->columns)) {
            $columns = array_combine($this->columns, $this->columns);
        } else {
            $columns = $this->columns;
        }
        foreach ($sheet_data as $values) {
            $column = [];
            foreach ($columns as $key => $label) {
                if (isset($values[$key])) {
                    $column[$label] = $values[$key];
                }
            }
            $new_data[] = $column;
        }
        return $new_data;
    }
}