<?php
/**
 * Created by Andrey Stepanenko.
 * User: webnitros
 * Date: 11.08.2020
 * Time: 11:45
 */
namespace Excel\Xlsx;

use Excel\Xlsx\Reader;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Settings;
use PhpOffice\PhpSpreadsheet\Collection\CellsFactory;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;
use Exception;

class ExcelReader extends Reader
{
    /** @var Spreadsheet $reader */
    public $reader = null;

    /**
     * @param int $seek
     */
    public function setSeek($seek)
    {
        $this->seek = $seek;
    }
    public function read(array $provider, $callback = null)
    {
        if (!file_exists($provider['file'])) {
            #$this->modx->log(modX::LOG_LEVEL_ERROR, sprintf($this->modx->lexicon('msimportexport.err_nf_file'), $provider['file']));
            return false;
        }

        $this->provider = $provider;
        if (!isset($this->provider['seek'])) {
            $this->provider['seek'] = 0;
        }

        try {
            if ($this->initReader()) {
                $objSheet = $this->reader->getActiveSheet();
                $rowIterator = $objSheet->getRowIterator();
                $emptyValue = 0;
                $index = 1;

                if (isset($this->provider['seek']) && !empty($this->provider['seek'])) {
                    $index = $this->provider['seek'];
                    $rowIterator->resetStart($index);
                }

                while ($rowIterator->valid()) {
                    $cellIterator = $rowIterator->current()->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);
                    $data = array();
                    while ($cellIterator->valid()) {
                        $val = $cellIterator->current()->getValue();
                        if (!empty($val)) $emptyValue++;
                        $data[] = $val;
                        $cellIterator->next();
                    }

                    $rowIterator->next();
                    $index++;

                    if (empty($emptyValue)) {
                        $this->setSeek(-1);
                    } else {
                        $this->setSeek($index);
                    }

                    if (is_callable($callback)) {
                        if (empty($emptyValue) || $callback($this, $data) !== true) {
                            unset($data);
                            unset($objSheet);
                            $this->disconnect();
                            return true;
                        }
                    }
                    $emptyValue = 0;
                }
                unset($data);
                unset($objSheet);
                $this->disconnect();
            }
        } catch (Exception $e) {
            if ($e->getLine() == 125 || $e->getLine() == 72 ) {
                $this->setSeek(-1);
                return true;
            }
            return 'Exception ' . $e->getMessage() . '. Info:';
        }

        return true;
    }

    /**
     * @return int|null
     */
    public function getTotalRows()
    {
        return -1;
    }


    private function disconnect()
    {
        $this->reader->disconnectWorksheets();
        unset($this->reader);
    }

    /**
     * @return bool
     */
    private function initReader()
    {
        if (!$this->reader) {
            // TODO тут шаги передавались
            $chunkSize = $this->limitStep;
            try {

                /*if ($cache = $this->getCacheAdapter()) {
                    Settings::setCache($cache);
                }*/

                $this->inputFileType = IOFactory::identify($this->provider['file']);
                $objReader = IOFactory::createReader($this->inputFileType);
                $objReader->setReadDataOnly(true);

                $readFilter = new ReadFilter();
                $readFilter->setRows((int)$this->provider['seek'], $chunkSize + 1);
                $objReader->setReadFilter($readFilter);
                $this->reader = $objReader->load($this->provider['file']);
                $this->reader->setActiveSheetIndex(0);

                unset($objReader);

            } catch (Exception $e) {
                return  'Exception ' . $e->getMessage();
            }
            return true;
        }
        return false;
    }


    private $limitStep = 500;

    private function setLimitStep($limit)
    {
        $this->limitStep = $limit;
    }

}