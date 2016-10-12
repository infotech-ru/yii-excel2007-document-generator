<?php
/*
 * This file is part of the infotech/yii-excel2007-document-generator package.
 *
 * (c) Infotech, Ltd
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Infotech\Excel2007DocumentGenerator;

use DOMDocument;
use DOMElement;
use DOMNode;
use Infotech\DocumentGenerator\Renderer\RendererInterface;
use CException;

class Excel2007Renderer implements RendererInterface
{
    const RELTYPE_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
    const RELTYPE_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
    const RELTYPE_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

    /**
     * Render template with data.
     *
     * @param string $templatePath
     * @param array $data
     * @throws \CException While template read error or temporary document write error
     * @return string Rendered document as binary string
     */
    public function render($templatePath, array $data)
    {
        $tmpPath = sys_get_temp_dir() . '/' . uniqid('excel_render_');

        if (!@copy($templatePath, $tmpPath)) {
            throw new CException('Error while reading a template file');
        }

        $xlsx = new Excel2007File($tmpPath);

        $imageSubstitutions = array();
        $doc = $xlsx->fetchXml('xl/sharedStrings.xml');
        $stringNodes = self::xpath($doc, array('ws' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'))
            ->query('/ws:sst/ws:si');
        foreach ($stringNodes as $stringIdx => $stringNode) {
            /** @var DOMNode $stringNode */
            if ($placeholderParams = self::fetchImagePlaceholderParams($stringNode->textContent)) {
                $dataIdx = array_shift($placeholderParams);
                if (isset($data[$dataIdx])) {
                    $imageFile = 'xl/media/' . md5($dataIdx) . '.jpeg';
                    $imageSubstitutions[$stringIdx] = array(
                        'dataIdx' => $dataIdx,
                        'params' => $placeholderParams,
                        'filename' => $imageFile,
                    );
                    $xlsx->putEntry($imageFile, $data[$dataIdx], 'image/jpeg');
                    unset($data[$dataIdx]);
                    $stringNode->parentNode->replaceChild(self::createDomFragment($doc, '<si><t> </t></si>'), $stringNode);
                }
            } else {
                $modifiedXml = preg_replace_callback('/\$\{([^}]+)\}/', function ($matches) use ($data) {
                    return !isset($data[$dataIdx = strip_tags($matches[1])])
                        ? $matches[0]
                        : htmlspecialchars($data[$dataIdx]);
                }, $doc->saveXML($stringNode));
                $stringNode->parentNode->replaceChild(self::createDomFragment($doc, $modifiedXml), $stringNode);
            }
        }

        if ($imageSubstitutions) {
            $imageNodesXPath = '/ws:worksheet/ws:sheetData/ws:row/ws:c[@t="s" and (ws:v="'
                . implode('" or ws:v="', array_keys($imageSubstitutions))
                . '")]';
            $sheetFiles = $this->fetchRelatedFiles(
                $xlsx->fetchXml($this->getRelsFileName('xl/workbook.xml')),
                self::RELTYPE_WORKSHEET
            );
            foreach ($sheetFiles as $sheetFile) {
                $sheetFile = 'xl/' . $sheetFile;
                $drawingFile = $this->getDrawingFileNameForWorkSheet($xlsx, $sheetFile);
                $drawingDoc = $xlsx->fetchXml($drawingFile);
                $drawingRelsDoc = $this->fetchRelsXmlFor($xlsx, $drawingFile);

                $doc = $xlsx->fetchXml($sheetFile);
                $xpath = self::xpath($doc, array('ws' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'));
                $imageNodes = $xpath->query($imageNodesXPath);
                foreach ($imageNodes as $imageNode) {
                    /** @var \DOMElement $imageNode */

                    $imageSubstitution = $imageSubstitutions[(int)$xpath->evaluate('string(ws:v)', $imageNode)];
                    $cellAddress = preg_split('/(?<=[A-Z])(?=[0-9])/', $imageNode->getAttribute('r'));
                    $rowFrom = -1 + (int)$cellAddress[1];
                    $colFrom = -1 + array_reduce(
                        preg_split('//', $cellAddress[0], -1, PREG_SPLIT_NO_EMPTY),
                        function ($acc, $digit) { return $acc * 26 + ord($digit) - ord('A') + 1; },
                        0
                    );
                    list($width, $height) = explode('x', $imageSubstitution['params'][0]);
                    $this->addRelationToRelsXml(
                        $drawingRelsDoc,
                        self::RELTYPE_IMAGE,
                        $imageSubstitution['filename'],
                        $imageRelId
                    );
                    $this->addImageToDrawingXml($drawingDoc, array(
                        'column' => $colFrom,
                        'row' => $rowFrom,
                        'width' => $width,
                        'height' => $height,
                        'title' => basename($imageSubstitution['filename']),
                        'relId' => $imageRelId,
                    ));
                }
            }
        }

        if (!$xlsx->close() ) {
            throw new CException('Error while writing temporary rendered file');
        }

        return $this->getTemporaryFileContents($tmpPath);
    }

    /**
     * @param string $filePath
     * @return string
     */
    private function getTemporaryFileContents($filePath)
    {
        $contents = file_get_contents($filePath);
        unlink($filePath);

        return $contents;
    }

    /**
     * @param DOMDocument $relsXml
     *
     * @return array
     */
    private function fetchRelatedFiles($relsXml, $worksheetRelType)
    {
        $relatedFiles = array();
        $rels = self::xpath($relsXml, array('r' => 'http://schemas.openxmlformats.org/package/2006/relationships'))
            ->query("/r:Relationships/r:Relationship[@Type='{$worksheetRelType}']");
        foreach ($rels as $rel) {
            /** @var DomElement $rel */
            $relatedFiles[] = $rel->getAttribute('Target');
        }

        return $relatedFiles;
    }

    /**
     * @param $file
     *
     * @return string
     */
    private function getRelsFileName($file)
    {
        return dirname($file) . '/_rels/' . basename($file) . '.rels';
    }

    /**
     * @param DOMDocument $doc
     * @param array  $image
     *
     * @return DOMDocument
     */
    private function addImageToDrawingXml(DOMDocument $doc, $image)
    {
        $drawingsNode = $doc->firstChild;
        $drawingsNode->appendChild($drawingNode = $doc->createElement('xdr:oneCellAnchor'));
        $drawingNode->appendChild($drawingFromNode = $doc->createElement('xdr:from'));
        $drawingFromNode->appendChild($doc->createElement('xdr:col', $image['column']));
        $drawingFromNode->appendChild($doc->createElement('xdr:colOff', self::pixelsToEMUs(0)));
        $drawingFromNode->appendChild($doc->createElement('xdr:row', $image['row']));
        $drawingFromNode->appendChild($doc->createElement('xdr:rowOff', self::pixelsToEMUs(0)));
        $drawingNode->appendChild($pictureExtNode = $doc->createElement('xdr:ext'));
        $pictureExtNode->setAttribute('cx', self::pixelsToEMUs($image['width']));
        $pictureExtNode->setAttribute('cy', self::pixelsToEMUs($image['height']));
        $drawingNode->appendChild($pictureNode = $doc->createElement('xdr:pic'));
        $pictureNode->appendChild($pictureNvPropsNode = $doc->createElement('xdr:nvPicPr'));
        $pictureNvPropsNode->appendChild($pictureNvPropNode = $doc->createElement('xdr:cNvPr'));
        $pictureNvPropNode->setAttribute('id', 0);
        $pictureNvPropNode->setAttribute('name', $image['title']);
        $pictureNvPropsNode->appendChild($pictureCNvPropNode = $doc->createElement('xdr:cNvPicPr'));
        $pictureNode->appendChild($imageContainerNode = $doc->createElement('xdr:blipFill'));
        $imageContainerNode->appendChild($imageNode = $doc->createElement('a:blip'));
        $imageNode->setAttribute('r:embed', $image['relId']);
        $imageNode->setAttribute('cstate', 'print');
        $imageContainerNode->appendChild($imageStretchNode = $doc->createElement('a:stretch'));
        $pictureNode->appendChild($pictureSpPropsNode = $doc->createElement('xdr:spPr'));
        $pictureSpPropsNode->setAttribute('bwMode', 'auto');
        $pictureSpPropsNode->appendChild($pictureXfrmNode = $doc->createElement('a:xfrm'));
        $pictureXfrmNode->appendChild($pictureXfrmOffNode = $doc->createElement('a:off'));
        $pictureXfrmOffNode->setAttribute('x', self::pixelsToEMUs(0));
        $pictureXfrmOffNode->setAttribute('y', self::pixelsToEMUs(0));
        $pictureXfrmNode->appendChild($pictureXfrmExtNode = $doc->createElement('a:ext'));
        $pictureXfrmExtNode->setAttribute('cx', self::pixelsToEMUs($image['width']));
        $pictureXfrmExtNode->setAttribute('cy', self::pixelsToEMUs($image['height']));
        $pictureSpPropsNode->appendChild($picturePresetGeomNode = $doc->createElement('a:prstGeom'));
        $picturePresetGeomNode->setAttribute('prst', 'rect');
        $picturePresetGeomNode->appendChild($doc->createElement('a:avLst'));
        $pictureSpPropsNode->appendChild($lnNode = $doc->createElement('a:ln'));
        $lnNode->appendChild($doc->createElement('a:noFill'));
        $drawingNode->appendChild($doc->createElement('xdr:clientData'));

        return $doc;
    }

    /**
     * @param DOMDocument $doc
     * @param $relType
     * @param $fileName
     * @param $relId
     *
     * @return DOMDocument
     */
    private function addRelationToRelsXml(DOMDocument $doc, $relType, $fileName, &$relId = null)
    {
        $relationsNode = $doc->firstChild;
        $relId = 'rId' . ($relationsNode->childNodes->length + 1);
        $relationshipNode = $doc->createElement('Relationship');
        $relationshipNode->setAttribute('Id', $relId);
        $relationshipNode->setAttribute('Type', $relType);
        $relationshipNode->setAttribute('Target', '../' . substr($fileName, 3));
        $relationsNode->appendChild($relationshipNode);

        return $doc;
    }

    /**
     * @param Excel2007File $xlsx
     * @param string        $fileName
     *
     * @return DOMDocument
     */
    private function fetchRelsXmlFor(Excel2007File $xlsx, $fileName)
    {
        $relsFileName = $this->getRelsFileName($fileName);
        if (!$xlsx->hasEntry($relsFileName)) {
            $xlsx->putEntry($relsFileName, <<<'XML'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
XML
            , 'application/vnd.openxmlformats-package.relationships+xml');
        }

        return $xlsx->fetchXml($relsFileName);
    }

    /**
     * @param DOMDocument $doc
     *
     * @param array       $namespaces
     *
     * @return \DOMXPath
     */
    private static function xpath(DOMDocument $doc, $namespaces = [])
    {
        $xpath = new \DOMXPath($doc);

        foreach ($namespaces as $prefix => $url) {
            $xpath->registerNamespace($prefix, $url);
        }

        return $xpath;
    }

    /**
     * @param Excel2007File $xlsx
     * @param string $sheetFile
     *
     * @return string
     */
    private function getDrawingFileNameForWorkSheet(Excel2007File $xlsx, $sheetFile)
    {
        $sheetRelsDoc = $this->fetchRelsXmlFor($xlsx, $sheetFile);
        $drawingFiles = $this->fetchRelatedFiles($sheetRelsDoc, self::RELTYPE_DRAWING);

        if (!$drawingFiles) {
            $drawingFile = 'xl/drawings/' . str_replace('sheet', 'drawing', basename($sheetFile));
            $xlsx->putEntry($drawingFile, <<<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
XML
            , 'application/vnd.openxmlformats-officedocument.drawing+xml');
            $this->addRelationToRelsXml($sheetRelsDoc, self::RELTYPE_DRAWING, $drawingFile, $drawingRelationId);

            $sheetDoc = $xlsx->fetchXml($sheetFile);
            $sheetDoc->firstChild->appendChild($drawingNode = $sheetDoc->createElement('drawing'));
            $drawingNode->setAttribute('r:id', $drawingRelationId);
        } else {
            $drawingFile = dirname($sheetFile) . '/' . $drawingFiles[0];
        }

        return $drawingFile;
    }

    /**
     * @param string $cellText
     *
     * @return array|false
     */
    private static function fetchImagePlaceholderParams($cellText)
    {
        $cellText = strip_tags($cellText);

        if (substr($cellText, 0, 2) == '${' && substr($cellText, -1) == '}') {
            $params = explode(':', strip_tags(substr($cellText, 2, -1)));
            if (count($params) > 1) {
                return $params;
            }
        }

        return false;
    }

    /**
     * @param DOMDocument $doc
     * @param string $xml
     *
     * @return \DOMDocumentFragment
     */
    private static function createDomFragment(DOMDocument $doc, $xml)
    {
        $fragment = $doc->createDocumentFragment();
        $fragment->appendXML($xml);

        return $fragment;
    }

    private static function pixelsToEMUs($pixels)
    {
        return $pixels * 914400 / 72;
    }
}
