<?php

namespace spec\Infotech\Excel2007DocumentGenerator;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpSpec\ObjectBehavior;
use ZipArchive;

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

                $book = IOFactory::load($tmpfile);
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
            },
            'notContains' => function($subject, $strings) use ($createTmpFile) {
                $tmpfile = $createTmpFile($subject);

                $book = IOFactory::load($tmpfile);
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

                return !array_diff((array)$strings, array_intersect((array)$strings, $matches[0]));
            },
            'zippedFilesExists' => function($subject, $filePaths) use ($createTmpFile) {
                $zip = new ZipArchive();
                $zip->open($createTmpFile($subject));


                $filePatterns = array_map(function ($pattern) {
                    return '/^' . str_replace(array('\*', '\?'), array('.*', '.'), preg_quote($pattern, '/')) . '$/';
                }, (array)$filePaths);
                for ($i = 0; $file = $zip->statIndex($i); $i++) {
                    foreach ($filePatterns as $filePatternIdx => $filePattern) {
                        if (preg_match($filePattern, $file['name'])) {
                            unset($filePatterns[$filePatternIdx]);
                        }
                    }
                }
                unlink($zip->filename);

                return count($filePatterns) === 0;
            },
            'zippedFilesCountEquals' => function($subject, $filePaths, $count) use ($createTmpFile) {
                $zip = new ZipArchive();
                $zip->open($createTmpFile($subject));

                $actuallyCount = 0;
                $filePatterns = array_map(function ($pattern) {
                    return '/^' . str_replace(array('\*', '\?'), array('.*', '.'), preg_quote($pattern, '/')) . '$/';
                }, (array)$filePaths);
                for ($i = 0; $file = $zip->statIndex($i); $i++) {
                    foreach ($filePatterns as $filePatternIdx => $filePattern) {
                        if (preg_match($filePattern, $file['name'])) {
                            $actuallyCount++;
                            break;
                        }
                    }
                }
                unlink($zip->filename);

                return $actuallyCount == $count;
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

    function it_should_insert_image_file()
    {
        $fixtureTemplate = __DIR__ . '/../../fixtures/simple_template.xlsx';
        $data = [
            'IMAGE' => file_get_contents(__DIR__ . '/../../fixtures/picture1.jpg'),
            'IMAGE2' => file_get_contents(__DIR__ . '/../../fixtures/picture2.jpeg'),
        ];

        $result = $this->render($fixtureTemplate, $data);
        $result->shouldBeXlsXDocument();
        $result->shouldNotContains(['${IMAGE:200x300}']);
        $result->shouldNotContains(['${IMAGE2:200x300}']);
        $result->shouldNotContains(['${IMAGE:100x200}']);
        $result->shouldZippedFilesExists(['xl/media/*']);
        $result->shouldZippedFilesCountEquals(['xl/media/*'], 2);
    }
}
