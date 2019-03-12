<?php
/**
 * Created by PhpStorm.
 * User: Mohamadreza
 * Date: 3/11/2019
 * Time: 12:19 PM
 */

namespace Waxwink\Report;


use Illuminate\Support\Collection;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excel
{
    /**
     * @var Collection
     */
    protected $collection;

    /**
     * @var array
     */
    protected $keys = [];

    /**
     * @var Spreadsheet
     */
    protected $spreadsheet;

    /**
     * @var Xlsx
     */
    protected $writer;

    protected $titles = [];

    /**
     * Excel constructor.
     * @param Collection $collection
     * @param array $titles
     */
    public function __construct(Collection $collection, array $titles =[])
    {
        $this->setCollection($collection);
        $this->setTitles($titles);
        $this->setSpreadsheet();
    }

    /**
     * @return array
     */
    public function getTitles(): array
    {
        return $this->titles;
    }

    /**
     * @param array $titles
     */
    public function setTitles(array $titles): void
    {
        $this->titles = $titles;
        $this->setKeys();
    }

    /**
     * @return Spreadsheet
     */
    public function getSpreadsheet(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    /**
     * @param Spreadsheet $spreadsheet
     */
    protected function setSpreadsheet(Spreadsheet $spreadsheet = null): void
    {
        if ($spreadsheet == null){
            $spreadsheet = new Spreadsheet();
        }
        $this->spreadsheet = $spreadsheet;
    }

    /**
     * @return array
     */
    public function getKeys(): array
    {
        return $this->keys;
    }

    /**
     */
    public function setKeys(): void
    {
        foreach ($this->getTitles() as $key => $title) {
            $this->keys[] = $key;
        }
    }

    /**
     * @return Collection
     */
    public function getCollection(): Collection
    {
        return $this->collection;
    }

    /**
     * @param Collection $collection
     */
    public function setCollection(Collection $collection): void
    {
        $this->collection = $collection;
    }

    public function putTitles()
    {
        $titles = \array_values($this->getTitles());

        $this->spreadsheet
            ->getActiveSheet()
            ->fromArray($titles, null, "A1");

    }

    public function putData()
    {
        $sheet = $this->spreadsheet->getActiveSheet();
        $data = [];
        $keys = $this->getKeys();
        $this->getCollection()->each(function ($item, $i) use(&$data, $keys) {
            foreach ($keys as $key) {
                $data[$i][$key] = (is_array($item))? $item[$key]:$item->$key;
            }
        });

        $sheet->fromArray($data, null,'A2');

    }

    public function getLastColumn()
    {
        return $this->spreadsheet->getActiveSheet()->getHighestColumn();
    }

    public function getLastRow()
    {
        return $this->spreadsheet->getActiveSheet()->getHighestRow();
    }

    public function makeTitlesBold()
    {
        $from = "A1";
        $to = $this->getLastColumn() . "1";
        $this->getSpreadsheet()->getActiveSheet()
            ->getStyle("$from:$to")
            ->getFont()->setBold(true);

        return $this;
    }

    public function makeTitlesColorful()
    {
        $from = "A1";
        $to = $this->getLastColumn() . "1";
        $this->getSpreadsheet()->getActiveSheet()
            ->getStyle("$from:$to")
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_DARKGREEN);


        $this->getSpreadsheet()->getActiveSheet()
            ->getStyle("$from:$to")
            ->getFont()->setColor(new Color(Color::COLOR_WHITE));

        return $this;
    }


    public function enableAutosize()
    {
        $from = "A";
        $to = $this->getLastColumn();
        $to++;

        for ($column = $from; $column != $to; $column++) {
            $this->getSpreadsheet()->getActiveSheet()->getColumnDimension($column)->setAutoSize(true);
        }

        return $this;
    }

    public function alignCenter()
    {
        $from = "A1";
        $to = $this->getLastColumn() . "100";
        $this->getSpreadsheet()->getActiveSheet()->getStyle("$from:$to")->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $this->getSpreadsheet()->getActiveSheet()->getStyle("$from:$to")->getAlignment()
            ->setVertical(Alignment::VERTICAL_CENTER);

        return $this;
    }

    public function setColumnWidth($column, $value)
    {
        $this->getSpreadsheet()->getActiveSheet()->getColumnDimension($column)->setWidth($value)->setAutoSize(false);
        return $this;
    }

    public function wrapTextInColumn($column)
    {
        $from = $column . '1';
        $to = $column. $this->getLastRow();
        $this->getSpreadsheet()->getActiveSheet()->getStyle("$from:$to")->getAlignment()->setWrapText(true);
        return $this;

    }

    public function setFont($font)
    {
        $from = 'A1';
        $to = $this->getLastColumn(). $this->getLastRow();
        $this->getSpreadsheet()->getActiveSheet()->getStyle("$from:$to")->getFont()->setName($font);
        return $this;
    }

    /**
     * @param string $filename
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function export($filename = 'table')
    {
        $filename .= '.xlsx';

        $this->putTitles();
        $this->putData();

        $this->writer = new Xlsx($this->spreadsheet);
        $this->writer->save($filename);

        $this->setStyle();
        $this->update();

        return $filename;

    }

    public function setStyle(){
        $this->makeTitlesBold();
        $this->enableAutosize();
        $this->alignCenter();

    }

    public function update($filename = 'table')
    {
        $filename .= '.xlsx';
        $this->writer->save($filename);
        return $filename;
    }


}
