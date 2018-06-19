<?php


namespace EasyExcel\Read;


class ExcelToArray
{
    private $path;          //完整路径
    private $fileName;
    private $ext;
    private $data = array();
    private $readObj;
    private $loadObj;
    private $fileType = array(
        'xls' => 'Excel5',
        'xlsx' => 'Excel2007',
        'csv' => 'CSV',
        'xsl' => 'SYLK',
    );

    /**
     * ExcelToArray constructor.
     * @param $path
     * @throws \Exception
     */
    public function __construct($path)
    {
        $this->path = $path;
        $this->fileName = $this->getFileName();
        $this->ext = $this->getExt();
        if (isset($this->fileType[strtolower($this->ext)])) {
            $this->readObj = \PHPExcel_IOFactory::createReader(strtolower($this->ext));
            $this->readObj->setReadDataOnly(true); //只读取数据
        } else {
            throw new \Exception('File Extension Is Not Illegal');
        }
        return $this;
    }

    /**
     * 获取路径中的文件名
     * @return string
     */
    private function getFileName(){
        return basename($this->path);
    }

    /**
     * 获取文件后缀
     *
     * @return mixed
     */
    public function getExt()
    {
        $extName = explode('.', $this->fileName);
        return end($extName);
    }

    /**
     * 获取行数
     * @return mixed
     */
    public function getRowNumber()
    {
        if (!$this->loadObj) {
            $this->load();
        }
        return $this->loadObj->getActiveSheet()->getHighestRow();
    }

    /**
     * 加载文件到内存
     * @return $this
     */
    public function load(){
        $this->loadObj = $this->readObj->load($this->path);
        return $this;
    }

    /**
     * 分批加载文件到内存
     * @param ChunkReadFilter $chunkReadFilter
     * @return $this
     */
    public function loadByChunk(ChunkReadFilter $chunkReadFilter)
    {
        $this->readObj->setReadFilter($chunkReadFilter);
        $this->loadObj = $this->readObj->load($this->path);
        return $this;
    }

    /**
     * 切换工作薄
     * @param $index
     * @return $this
     */
    public function setSheet($index)
    {
        $this->loadObj->setActiveSheetIndex($index);
        return $this;
    }

    /**
     * 获取所有Sheets
     * @return mixed
     */
    public function getAllSheet(){
        return $this->loadObj->getAllSheets();
    }

    /**
     * 获取内容
     * @return array
     */
    public function getData()
    {
        $loadedWorkSheet = $this->loadObj->getActiveSheet(); //获取当前激活Sheet
        $maxRow = $loadedWorkSheet->getHighestRow(); //获取最大行 int
        $maxColumn = $loadedWorkSheet->getHighestColumn(); //获取最大列 (A-Z)
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($maxColumn); //根据列名获取index

        //从第二行获取行数据 （第一行是字段）
        for ($row=2; $row <= $maxRow; $row++) {
            //从第一列获取列的数据
            for ($col = 0; $col < $highestColumnIndex; $col++) {
                $this->data[$row - 2][] = (string)$loadedWorkSheet->getCellByColumnAndRow($col, $row)->getValue();
            }
        }
        $this->loadObj->disconnectWorksheets();
        unset($objReader, $objPHPExcel, $objWorkSheet, $highestColumnIndex);
        return $this->data;
    }
}