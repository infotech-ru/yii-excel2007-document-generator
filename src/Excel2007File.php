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

use CException;
use DOMDocument;
use DOMXPath;
use ZipArchive;

class Excel2007File
{
    /** @var ZipArchive */
    private $zip;
    private $changedFiles = array();
    private $types = array();
    /** @var DOMDocument */
    private $typesDoc;
    /** @var DOMXPath  */
    private $typesXpath;


    public function __construct($filename)
    {
        $this->zip = new ZipArchive();
        if (!file_exists($filename)) {
            throw new CException('File "' . $filename . '" is not found');
        }
        if (!is_readable($filename)) {
            throw new CException('Unable to read file "' . $filename . '"');
        }
        if (!$this->zip->open($filename)
                || !$this->typesDoc = self::createDomDocument($this->zip->getFromName('[Content_Types].xml'))
        ) {
            throw new CException('Unsupported file format. Required xlsx (Excel2007) file');
        }

        $this->typesXpath = new DOMXPath($this->typesDoc);
        $this->typesXpath->registerNamespace('t', 'http://schemas.openxmlformats.org/package/2006/content-types');
    }

    public function close()
    {
        foreach (array_keys($this->changedFiles) as $changedFileName) {
            if (!$contents = $this->fetchString($changedFileName)) {
                $this->zip->deleteName($changedFileName);
                if ($node = $this->typesXpath->query(self::getTypeOverrideSelector($changedFileName))->item(0)) {
                    $node->parentNode->removeChild($node);
                }
            } else {
                $this->zip->addFromString($changedFileName, $contents);
                if (!$this->getTypeFor($changedFileName)) {
                    $this->typesDoc->firstChild->appendChild($typeOverride = $this->typesDoc->createElement('Override'));
                    $typeOverride->setAttribute('PartName', '/' . $changedFileName);
                    $typeOverride->setAttribute('ContentType', $this->types[$changedFileName]);
                }
            }
        }

        $this->zip->addFromString('[Content_Types].xml', $this->typesDoc->saveXML());

        return $this->zip->close();
    }

    public function hasEntry($path)
    {
        return false !== $this->fetchString($path);
    }

    /**
     * @param string $path
     *
     * @return DOMDocument|false DOMDocument object or FALSE if content is not exists or is not valid XML
     */
    public function fetchXml($path)
    {
        if (!isset($this->changedFiles[$path])) {
            $raw = $this->zip->getFromName($path);
            $this->changedFiles[$path] = self::createDomDocument($raw) ?: $raw;
            $this->types[$path] = $this->getTypeFor($path);
        } elseif (!$this->changedFiles[$path] instanceof DOMDocument) {
            if ($doc = self::createDomDocument((string)$this->changedFiles[$path])) {
                $this->changedFiles[$path] = $doc;
            }
        }

        return $this->changedFiles[$path] instanceof DOMDocument ? $this->changedFiles[$path] : new DOMDocument();
    }

    /**
     * @param string $path
     *
     * @return string|false binary string or FALSE if content is not exists
     */
    public function fetchString($path)
    {
        if (!isset($this->changedFiles[$path])) {
            $string = $this->changedFiles[$path] = $this->zip->getFromName($path);
            $this->types[$path] = $this->getTypeFor($path);
        } elseif ($this->changedFiles[$path] instanceof DOMDocument) {
            $string = $this->changedFiles[$path]->saveXML();
        } elseif (false === $this->changedFiles[$path]) {
            $string = false;
        } else {
            $string = (string)$this->changedFiles[$path];
        }

        return $string;
    }

    public function putEntry($path, $data, $type)
    {
        $this->changedFiles[$path] = $data;
        $this->types[$path] = $type;
    }

    /**
     * @param string $data
     *
     * @return DOMDocument|false
     */
    private static function createDomDocument($data)
    {
        if (!$data) {
            return false;
        }

        $doc = new DOMDocument();
        $doc->loadXML($data);

        return $doc;
    }

    private function getTypeFor($path)
    {
        return $this->typesXpath->evaluate('string(' . self::getTypeOverrideSelector($path) . '/@ContentType)');
    }

    /**
     * @param string $path
     *
     * @return string
     */
    private static function getTypeOverrideSelector($path)
    {
        return '/t:Types/t:Override[@PartName="/' . $path . '"]';
    }

}
