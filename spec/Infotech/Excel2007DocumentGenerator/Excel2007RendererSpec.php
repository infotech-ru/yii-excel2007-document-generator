<?php

namespace spec\Infotech\Excel2007DocumentGenerator;

use PHPExcel_IOFactory;
use PhpSpec\ObjectBehavior;
use Prophecy\Argument;

class Excel2007RendererSpec extends ObjectBehavior
{
    public function getMatchers()
    {
        $createTmpFile = function ($contents) {
            $tmpfile = tempnam(sys_get_temp_dir(), 'spec_fixture_');
            file_put_contents($tmpfile, $contents);

            return $tmpfile;
        };
        $extractDocument = function ($contents) use ($createTmpFile) {

            $tmpfile = $createTmpFile($contents);

            $docContents = @file_get_contents('zip://' . $tmpfile . '#xl/workbook.xml');
            unlink($tmpfile);

            return $docContents;
        };
        return [
            'beXlsXDocument' => function ($subject) use ($extractDocument) {
                return false !== $extractDocument($subject);
            },
            'contains' => function($subject, $strings) use ($createTmpFile) {
                $tmpfile = $createTmpFile($subject);

                $book = PHPExcel_IOFactory::load($tmpfile);
                unlink($tmpfile);

                $cellValues = array();

                foreach ($book->getAllSheets() as $sheet) {
                    foreach ($sheet->getRowIterator() as $row) {
                        foreach ($row->getCellIterator() as $cell) {
                            $cellValues[] = $cell->getValue();
                        }
                    }
                }

                preg_match_all(
                    '/' . implode('|', array_map('preg_quote', (array)$strings)) . '/',
                    implode("\n", $cellValues),
                    $matches
                );

                return !array_diff((array)$strings, $matches[0]);
            }
        ];
    }

    function it_is_initializable()
    {
        $this->shouldHaveType('Infotech\Excel2007DocumentGenerator\Excel2007Renderer');
    }

    function it_should_fillup_template_with_values()
    {
        $fixtureTemplate = __DIR__ . '/../../fixtures/simple_template.xlsx';
        $data = [
            'PLACEHOLDER_1' => 'Replaced placeholder 1',
            'PLACEHOLDER_2' => 'Replaced placeholder 2',
            'PLACEHOLDER_3' => '3',
        ];

        $result = $this->render($fixtureTemplate, $data);
        $result->shouldBeXlsXDocument();
        $result->shouldContains($data);
    }

    function it_should_fillup_template_with_values_containing_xml_symbols()
    {
        $fixtureTemplate = __DIR__ . '/../../fixtures/simple_template.xlsx';
        $data = [
            'PLACEHOLDER_1' => '20 < 5',
            'PLACEHOLDER_2' => '&mdash;',
            'PLACEHOLDER_3' => '<!-- dsfsdgsdh',
        ];

        $result = $this->render($fixtureTemplate, $data);
        $result->shouldBeXlsXDocument();
        $result->shouldContains($data);
    }

    function it_should_throw_an_exception_when_template_not_found()
    {
        $fixtureTemplate = __DIR__ . '/../../fixtures/unexistent_template.xlsx';
        $data = [
            'PLACEHOLDER_1' => 'Replaced placeholder 1',
        ];

        $this->shouldThrow('CException')->during('render', array($fixtureTemplate, $data));
    }
}
