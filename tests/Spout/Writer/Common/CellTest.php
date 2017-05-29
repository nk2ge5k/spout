<?php

namespace Box\Spout\Writer\Common;

use Box\Spout\Writer\Style\StyleBuilder;
use PHPUnit\Framework\TestCase;

class CellTest extends TestCase
{
    protected function styleMock()
    {
        $styleMock = $this
            ->getMockBuilder('Box\Spout\Writer\Style\Style');
        return $styleMock;
    }

    public function testValidInstance()
    {
        $this->assertInstanceOf('Box\Spout\Writer\Common\Cell', new Cell('cell'));
        $this->assertInstanceOf(
            'Box\Spout\Writer\Common\Cell',
            new Cell('cell-with-style', $this->styleMock()->getMock())
        );
    }

    public function testCellTypeNumeric()
    {
        $this->assertTrue((new Cell(0))->isNumeric());
        $this->assertTrue((new Cell(1))->isNumeric());
    }

    public function testCellTypeString()
    {
        $this->assertTrue((new Cell('String!'))->isString());
    }

    public function testCellTypeEmptyString()
    {
        $this->assertTrue((new Cell(''))->isEmpty());
    }

    public function testCellTypeEmptyNull()
    {
        $this->assertTrue((new Cell(null))->isEmpty());
    }

    public function testCellTypeBool()
    {
        $this->assertTrue((new Cell(true))->isBoolean());
        $this->assertTrue((new Cell(false))->isBoolean());
    }

    public function testCellTypeError()
    {
        $this->assertTrue((new Cell([]))->isError());
    }

    public function testApplyStyle()
    {
        $o =  new Cell('style');
        $baseStyle = (new StyleBuilder())->setFontBold()->build();
        $o->applyStyle($baseStyle);
        $this->assertTrue($o->getStyle()->isFontBold());
    }
}
