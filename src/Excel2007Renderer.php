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

use Infotech\DocumentGenerator\Renderer\RendererInterface;
use ZipArchive;
use CException;

class Excel2007Renderer implements RendererInterface
{
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

        $zip = new ZipArchive();
        if (!$zip->open($tmpPath) || !$sharedStrings = $zip->getFromName('xl/sharedStrings.xml')) {
            throw new CException('Unsupported template file format. Required xlsx (Excel2007) file');
        }

        $replaceFn = function ($matches) use ($data) {
            $placeholderName = strip_tags($matches[1]);

            return isset($data[$placeholderName]) ? htmlspecialchars($data[$placeholderName]) : $matches[0];
        };

        $zip->addFromString('xl/sharedStrings.xml', preg_replace_callback('/\$\{([^}]+)\}/', $replaceFn, $sharedStrings));

        if (!$zip->close() ) {
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
}
