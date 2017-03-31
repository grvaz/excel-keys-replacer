<?php

namespace ItPoet\ExcelKeysReplacer;

class Replacer
{
    protected $file;
    protected $left = '${';
    protected $right = '}$';
    protected $data = [];
    protected $sheet = 0;

    public function replace()
    {
        $values = $this->data;
        //$keys = array_keys($values);
        $xls = \PHPExcel_IOFactory::load($this->file);
        $xls->setActiveSheetIndex($this->sheet);
        $sheet = $xls->getActiveSheet();

        $rowIterator = $sheet->getRowIterator();
        foreach ($rowIterator as $row) {
            $cellIterator = $row->getCellIterator();
            foreach ($cellIterator as $cell) {
                preg_match_all("/". preg_quote($this->left, "/") . "(.+?)" . preg_quote($this->right, "/") . "/ism", $cell->getValue(), $toReplace);
                foreach ($toReplace[1] as $replace) {
                    if (isset($values[$replace])){
                        $sheet->setCellValue(
                            $cell->getCoordinate(),
                            str_replace($this->left . $replace . $this->right, $values[$replace], $cell->getValue())
                        );
                    }
                }
            }
        }
        return $xls;
    }

    public function file($path)
    {
        $this->file = $path;
    }

    public function keyWrapper($left, $right)
    {
        $this->left = $left;
        $this->right = $right;
    }

    public function data($array)
    {
        $this->data = $array;
    }

    public function sheet($sheet)
    {
        $this->sheet = abs((int)$sheet);
    }
}