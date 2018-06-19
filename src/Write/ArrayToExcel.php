<?php

namespace EasyExcel\Write;

class ArrayToExcel
{
    private $excelObj;
    private $colAttr = array(
        'A' => array(//列的属性设置
            'colName' => '',//第一行的列名
            'keyName' => '',//每一列对应的赋值数组的key值
            'width' => ''   //A列的宽度
        )
    );      //列属性
    private $rowAttr = array(
        'firstRowHeight' => '', //第一行的列名的高度
        'height' => ''         //2-OO无从行的高度
    );      //行属性
    private $options = array(
        'fileName' => '',           //导出的excel的文件的名称
        'sheetTitle' => '',         //每个工作薄的标题
        'creator' => '',            //创建者
        'lastModified' => '',       //最近修改时间
        'title' => '',              //当前活动的主题
        'subject' => '',
        'description' => '',
        'keywords' => '',
        'category' => '',
        'writeType' => ''           //输出类型 Excel5 Excel7 CSV
    );      //文件选项
    private $validData = array();       //数据有效性
    private $colDefaultDefine = array(
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
        'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    );  //默认列定义
    private $writeType = array(
        'excel5' => 'Excel5',
        'excel2007' => 'Excel2007',
        'csv' => 'CSV',
    );     //写入文件类型

    public function __construct(array $options)
    {
        if (!isset($options['fileName'])) {
            throw new \Exception('fileName is require');
        }
        if (!isset($options['path'])) {
            throw new \Exception('path is require');
        }
        if (!isset($options['writeType'])) {
            throw new \Exception('writeType is require');
        }

        $this->options['writeType'] = $this->writeType[strtolower($options['writeType'])];
        $this->options = $options;
        $this->excelObj = new \PHPExcel();
    }

    /**
     * 设置Excel第一行表头  一维数组/二维数组
     * @param array $col
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function setFirstRow(array $col)
    {
        if (empty($col)) {
            throw new \Exception('Col is not be Null');
        }

        $this->colAttr = $col;
        $obj = $this->excelObj->getActiveSheet();

        foreach ($col as $key => $val) {
            //设置 Vertical 和 Horizontal
            $this->excelObj->getActiveSheet()
                ->getStyle($key)
                ->getAlignment()
                ->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

            //设置第一行字段名
            $colIndex = !is_array($val) ? $this->colDefaultDefine[$key] : $key;
            $colValue = !is_array($val) ? $val : $val['colName'];
            $this->excelObj->getActiveSheet()->
            setCellValue($colIndex . '1', $colValue);

            //设置列的宽度 or 跟随字体长度宽度
            if (isset($val['width']) && !empty($val['width'])) {
                $this->excelObj->getActiveSheet()->
                getColumnDimension($colIndex)->setWidth($val['width']);
            } else {
                $this->excelObj->getActiveSheet()->
                getColumnDimension($colIndex)->setAutoSize(true);
            }
        }

        return $this;
    }

    public function fillData(array $data, array $valiData = array())
    {
        if (isset($this->options['lastModified'])) {
            $this->excelObj->getProperties()->setLastModifiedBy($this->options['lastModified']);
        }
        if (isset($this->options['title'])) {
            $this->excelObj->getProperties()->setTitle($this->options['title']);
        }
        if (isset($this->options['subject'])) {
            $this->excelObj->getProperties()->setSubject($this->options['subject']);
        }
        if (isset($this->options['description'])) {
            $this->excelObj->getProperties()->setDescription($this->options['description']);
        }
        if (isset($this->options['keywords'])) {
            $this->excelObj->getProperties()->setKeywords($this->options['keywords']);
        }
        if (isset($this->options['category'])) {
            $this->excelObj->getProperties()->setCategory($this->options['category']);
        }
        if (isset($this->options['category'])) {
            $this->excelObj->getProperties()->setTitle($this->options['category']);
        }

        //填充
        for ($p = 0; $p < count($data); $p++) {
            //行数num,第二行开始
            $row = $p + 2;
            // 设置数据的有效性
            if (isset($valiData) && !empty($valiData)) {
                //总分数据有效性下拉菜单
                $objValidNum = $this->excelObj->getActiveSheet()->getCell($valiData['list1'][0] . $row)->getDataValidation();
                $objValidNum->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST);
                $objValidNum->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION);
                $objValidNum->setAllowBlank(false);
                $objValidNum->setShowInputMessage(true);
                $objValidNum->setShowErrorMessage(true);
                $objValidNum->setShowDropDown(true);
                $objValidNum->setFormula1('"' . $valiData['list1'][1] . '"');
                $objValidNum->getActiveSheet()->getCell('F' . $row)->setDataValidation($objValidation1);

                //学期数据有效性下拉菜单
                $objValid = $this->excelObj->getActiveSheet()->getCell($valiData['list2'][0] . $row)->getDataValidation();
                $objValid->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST);
                $objValid->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION);
                $objValid->setAllowBlank(false);
                $objValid->setShowInputMessage(true);
                $objValid->setShowErrorMessage(true);
                $objValid->setShowDropDown(true);
                $objValid->setFormula1('"' . $valiData['list2'][1] . '"');
                $this->excelObj->getActiveSheet()->getCell('G' . $row)->setDataValidation($objValidation2);
            }

            //填充数据
            foreach ($this->colAttr as $colKey => $colVal) {
                $colIndex = !is_array($colVal) ? $this->colDefaultDefine[$colKey] : $colKey;
                $colVal = !is_array($colVal) ? $colVal : $colVal['colName'];
                $this->excelObj->getActiveSheet()->setCellValue($colIndex . $row, $data[$p][$colKey]);
            }

            //设置行高
            if (isset($this->rowAttr['height']) && !empty($this->rowAttr['height'])) {
                $this->excelObj->getActiveSheet()->getRowDimension($row)->setRowHeight($this->rowAttr['height']);
            }
        }
        ob_end_clean();
        ob_start();
        $objWriter = \PHPExcel_IOFactory::createWriter($this->excelObj, $this->options['writeType']);
        $objWriter->save($this->options['path'] . $this->options['fileName']);//php://output
    }
}

?>