<?php

namespace common\components\excel;

class Excel
{
    /**
     * 导入excel
     *
     * @param string $file_name
     * @param array $options
     * @return ExcelImport
     */
    public static function import($file_name, $options = [])
    {
        return new ExcelImport(array_merge(['file_name' => $file_name], $options));
    }

    /**
     * 导出excel
     *
     * @param array $options
     * @return ExcelExport
     */
    public static function export($options = [])
    {
        return new ExcelExport($options);
    }
}