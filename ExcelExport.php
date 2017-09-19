<?php

namespace common\components\excel;

use yii\base\Component;
use yii\base\InvalidConfigException;

class ExcelExport extends Component
{
    /**
     * 数据模型
     *
     * @var \yii\base\Model
     */
    public $models;

    /**
     * 数据行
     *
     * @var array
     */
    public $rows;

    /**
     * 输出格式
     *
     * @var string
     */
    public $format = 'Excel2007';

    /**
     * 下载文件名
     *
     * @var string
     */
    public $download_name = 'export';

    /**
     * 输出路径
     *
     * @var string
     */
    public $path = 'php://output';

    /**
     * 是否采用模板
     *
     * @var bool
     */
    public $template = false;

    /**
     * 是否直接在模板后添加数据
     *
     * @var bool
     */
    public $template_append = false;

    /**
     * @inheritdoc
     */
    public function init()
    {
        if (!isset($this->models) && !isset($this->rows)) {
            throw new InvalidConfigException('models or rows must be set');
        }
        if ($this->template && !file_exists($this->template)) {
            throw new InvalidConfigException('template file not exist');
        }
    }

    /**
     * 设置下载头
     */
    public function setDownloadHeaders()
    {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $this->download_name .'"');
        header('Cache-Control: max-age=0');
    }

    /**
     * 输出excel
     */
    public function toStream()
    {
        if ($this->template) {
            $this->format = \PHPExcel_IOFactory::identify($this->template);
            $sheet = \PHPExcel_IOFactory::createReader($this->format)->load($this->template);
            $worksheet = $sheet->getActiveSheet();
            if (!$this->template_append) {
                $worksheet->removeRow(2, $worksheet->getHighestRow() - 1);
            }
        } else {
            $sheet = new \PHPExcel();
            $worksheet = $sheet->getActiveSheet();
        }

        if (isset($this->models)) {
            $this->exportFromModel($worksheet);
        } elseif (isset($this->rows)) {
            $this->exportFromRows($worksheet);
        }
//        var_dump($worksheet->toArray(null, true, true, true));die;

        $this->setDownloadHeaders();
        $this->writeFile($sheet);
    }

    /**
     * @param \PHPExcel $sheet
     */
    public function writeFile(\PHPExcel $sheet)
    {
        $writer = \PHPExcel_IOFactory::createWriter($sheet, $this->format);
        $writer->save($this->path);
        exit;
    }

    /**
     * @param \PHPExcel_Worksheet $worksheet
     */
    protected function exportFromModel(\PHPExcel_Worksheet &$worksheet)
    {
        // todo
    }

    /**
     * @param \PHPExcel_Worksheet $worksheet
     */
    protected function exportFromRows(\PHPExcel_Worksheet &$worksheet)
    {
        $row_num = 2;
        foreach ($this->rows as $row) {
            $col_num = 1;
            foreach ($row as $column) {
                $worksheet->setCellValue(chr(64 + $col_num) . $row_num, $column);
                $col_num ++;
            }
            $row_num ++;
        }
    }
}