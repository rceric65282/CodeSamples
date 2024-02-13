<?php

require_once("vendor/autoload.php");
require_once("config.php");
require_once("print-handler.php");
require_once("PDFMerger.php");
require_once("twigFunctions.php");

use Ramsey\Uuid\Uuid;
use Ramsey\Uuid\Exception\UnsatisfiedDependencyException;
use Phpdocx\Create\CreateDocx;
use Phpdocx\Elements\WordFragment;
use GuzzleHttp\Psr7;
use GuzzleHttp\Exception\ClientException;
use GuzzleHttp\Exception\RequestException;
use GuzzleHttp\Exception\ServerException;
use GuzzleHttp\Exception\TransferException;
use KubAT\PhpSimple\HtmlDomParser;

/**
 * Class responsible for retrieving data from PeopleSoft and creating PDF
 * content.
 */
class Poke {
    // Transaction ID for the creation of a PDF file.
    private $transactionId;

    private $marginSettings = array(
        'marginLeft' => 1440,
        'marginRight' => 1440,
        'marginTop' => 1440,
        'marginBottom' => 1440);

    function __construct() {
        openlog("POKE", LOG_INFO, LOG_USER);
        $this->transactionId = self::createUniqueId();
    }

    /**
     * Converts HTML to DOCX.
     * @param $htmlContent string containing the HTML to convert.
     * @param $xmlElement XML document containing any info you may need.
     * @return string with the docx filename or null if something went wrong.
     */
    private function convertHtmlToDocx($htmlContent, $xmlElement, $settings) {
        $uniq = self::createUniqueId();
        $docxFilename = TMP_DIR . "/" . $uniq . ".docx";
        $format = $settings['ecvFormat'];
        try {
            if ($format == 'far') {
                $docxFilename = $this->farPrint($htmlContent, $xmlElement, $settings);
            } else {
                $docx = new Phpdocx\Create\CreateDocx();
                $docx->modifyPageLayout('letter', $this->marginSettings);
                $docx->embedHTML($htmlContent, array("removeLineBreaks" => True));
                $footer = $this->createFooterDocx($docx, $xmlElement, $settings);
                $docx->addFooter(array('default' => $footer));
                $docx->createDocx($docxFilename);
            }
        } catch (\Exception $e) {
            $this->recordLog("DOCX creation failed: " . $e->getMessage());
            return null;
        }
        if (file_exists($docxFilename)) {
            $this->recordLog("DOCX created in " . $docxFilename);
            return $docxFilename;
        } else {
            $this->recordLog("DOCX creation failed.");
            return null;
        }
    }

    /**
     * Converts HTML to PDF.
     * @param $htmlContent string HTML content.
     * @return string filename of the PDF file or null if conversion failed.
     */
    private function convertHtmlToPdf($htmlContent, $xmlElement, $settings) {
        $ecvFormat = $settings['ecvFormat'];
        if ($ecvFormat == 'far') {
            return $this->farPrint($htmlContent, $xmlElement, $settings);
        } else {
            return $this->convertHtmlToPdfCli($htmlContent, $xmlElement,
                    $settings, 0, 'Portrait');
        }
    }

    /**
     * Converts HTML to PDF. Depends on wkhtmltopdf to be installed as well as
     * xvfb so xvfb-run can give it an X session.
     * @param $htmlContent string HTML content.
     * @param $xmlElement XML model for the HTML.
     * @param $settings settings for creating the document.
     * @param $pageOffset page offset for page numbers (zero based).
     * @param $orientation 'Landscape' or 'Portrait'.
     * @return string filename of the PDF file or null if conversion failed.
     */
    private function convertHtmlToPdfCli($htmlContent, $xmlElement, $settings,
            $pageOffset, $orientation) {
        $uniq = self::createUniqueId();
        $pdfFilename = TMP_DIR . "/". $uniq . ".pdf";
        $footerArgs = $this->createFooterArgs($xmlElement, $settings);
        $marginArgs = "-B 25mm -L 25mm -R 25mm -T 25mm";
        $pageArgs = "--page-size Letter --zoom 1.15";
        $orientationArg = "-O " . $orientation;
        if ($orientation != 'Landscape' && $orientation != 'Portrait') {
            // Set to portrait if it's not Landscape or Portrait
            $orientationArg = "-O Portrait";
        }
        /*
        If your content is going to have tables and a footer, make sure you
        have these styles in your HTML before calling wkhtmltopdf:
            thead { display: table-header-group; }
            tfoot { display: table-row-group; }
            tr { page-break-inside: avoid; }
        otherwise you end up with overlapping text if a table spans a page.
        */
        $command = \h4cc\WKHTMLToPDF\WKHTMLToPDF::PATH . " --print-media-type "
                . $orientationArg . " --page-offset " . $pageOffset . " "
                . $pageArgs . " " . $marginArgs . " " . $footerArgs
                . " - " . $pdfFilename;
        $this->recordLog("Running " . $command);
        $handle = popen($command, "w");
        if ($handle == FALSE) {
            $this->recordLog("popen failed");
        } else {
            $result = fwrite($handle, $htmlContent);
            pclose($handle);
            if ($result == FALSE) {
                $this->recordLog("Failed to pipe HTML contents");
            }
        }
        if (file_exists($pdfFilename)) {
            $pdfContent = file_get_contents($pdfFilename);
            if ($pdfContent === FALSE) {
                $this->recordLog("Could not read " . $pdfFilename);
                return null;
            }
            return $pdfFilename;
        } else {
            $this->recordLog("wkhtmltopdf failed");
        }
        return null;
    }

    /**
     * Creates a document.
     * @param settings associative array with settings for creating the document.
     * @return associative array with key "document" associated with the filename of
     * the document that was generated and key "filename" with the recommended
     * filename for the document. Returns null if something went wrong.
     */
    public function createDocument($settings) {
        if ($this->validateSettings($settings) == FALSE) {
            return null;
        }
        $format = $settings["outputFormat"];
        try {
            $rawXml = NULL;
            $realXml = NULL;
            if (isset($settings["rawxml"])) {
                $rawXml = file_get_contents($settings["rawxml"]);
                if ($rawXml === FALSE) {
                    $rawXml = NULL;
                }
            }
            if (isset($settings["realxml"])) {
                $realXml = file_get_contents($settings["realxml"]);
                if ($realXml == FALSE) {
                    $realXml = NULL;
                }
            }
            if ($rawXml == NULL && $realXml == NULL) {
                $rawXml = $this->requestRawXml($settings);
            }
            if ($rawXml == null && $realXml == NULL) {
                return null;
            }
            $docFilename = null;
            if (DEBUG_XML && $format == "rawxml" && $rawXml != NULL) {
                $uniq = self::createUniqueId();
                $docFilename = TMP_DIR . "/" . $uniq . ".xml";
                $written = file_put_contents($docFilename, $rawXml);
                if ($written == FALSE) {
                    $this->recordLog("Failed to write raw XML");
                    return null;
                }
            }
            if ($docFilename == null) {
                if ($realXml == NULL) {
                    $realXml = $this->transformXml($rawXml);
                }
                if ($realXml == null) {
                    $this->recordLog("Failed to transform XML");
                    return null;
                }
                if (DEBUG_XML && $format == "xml") {
                    $uniq = self::createUniqueId();
                    $docFilename = TMP_DIR . "/" . $uniq . ".xml";
                    $written = file_put_contents($docFilename, $realXml);
                    if ($written == FALSE) {
                        $this->recordLog("Failed to write transformed XML");
                        return null;
                    }
                }
            }
            $name = "";
            if ($docFilename == null) {
                $element = simplexml_load_string($realXml);
                if ($element == FALSE) {
                    $this->recordLog("Failed to create simple XML");
                    return null;
                }
                $this->recordLog("Transformed XML loaded");
                $this->stripTags($element);
                $this->recordLog("Tags stripped");
                if ($element == null) {
                    return null;
                }
                $html = $this->createHtml($element, $settings);
                if ($html == null) {
                    return null;
                }
                if ($format == "docx") {
                    $docFilename = $this->convertHtmlToDocx($html, $element, $settings);
                    if ($docFilename == null) {
                        return null;
                    }
                } else if ($format == "html") {
                    $uniq = self::createUniqueId();
                    $docFilename = TMP_DIR . "/" . $uniq . ".html";
                    $written = file_put_contents($docFilename, $html);
                    if ($written == FALSE) {
                        $this->recordLog("Failed to write HTML content");
                        return null;
                    }
                } else if ($format == "pdf") {
                    $docFilename = $this->convertHtmlToPdf($html, $element, $settings);
                    if ($docFilename == null) {
                        return null;
                    }
                } else {
                    $this->recordLog("Unrecognized format " . $format);
                    return null;
                }
                $name = $element->name;
            }
            $recommendedFilename =
                    $this->recommendedFilename($name, $settings);
            $result = [ "document" => $docFilename, "filename" => $recommendedFilename ];
            $this->recordLog("Created: " . $docFilename . " -> ". $recommendedFilename);
            return $result;
        } catch (Exception $e) {
            $this->recordLog("Unexpected exception: " . $e->getMessage());
            return null;
        } catch (\Throwable $t) {
            $this->recordLog("Unexpected throwable: " . $t->getMessage());
            return null;
        }
    }

    /**
     * Create footer arguments for wkhtmltopdf.
     */
    private function createFooterArgs($simpleXmlElement, $settings) {
        // Look at $settings ecvFormat if generic footer isn't enough
        $name = $simpleXmlElement->name;
        $acadyear = $settings["academicYear"];
        $date = date('Y-m-d');
        $footerLeft = '--footer-left "' . $date . '"';
        $text = $this->createFooterText($simpleXmlElement, $settings);
        $ecvFormat = $settings["ecvFormat"];
        $footerCenter = '--footer-center "' . $text . '"';
        $footerRight = '--footer-right "[page]"';
        // font size is in pts
        $footerMainArgs = '--footer-font-name "Arial" --footer-font-size 11 --footer-spacing 5';
        $footerArgs = $footerMainArgs . " " . $footerLeft . " " . $footerCenter . " " . $footerRight;
        return $footerArgs;
    }

    /**
     * Create a footer for a DOCX document.
     * Format is Date, name/academic year, page number.
     */
    private function createFooterDocx($docx, $simpleXmlElement, $settings) {
        $date = new WordFragment($docx, 'defaultFooter');
        $font = 'Arial';
        $fontSize = '11';   // Font size is in pts
        $date->addDateAndHour(array(
            'dateFormat' => 'yyyy-MM-dd',
            'font' => $font,
            'fontSize' => $fontSize,
            'textAlign' => 'left'
        ));

        $name = $simpleXmlElement->name;
        $text = $this->createFooterText($simpleXmlElement, $settings);
        $label = new WordFragment($docx, 'defaultFooter');
        $label->addText($text, array(
            'font' => $font,
            'fontSize' => $fontSize,
            'textAlign' => 'center'
        ));

        $pageNumber = new WordFragment($docx, 'defaultFooter');
        $pageNumber->addPageNumber('numerical', array(
            'textAlign' => 'right',
            'font' => $font,
            'fontSize' => $fontSize
        ));

        $values = array(array(
            array('value' => $date, 'vAlign' => 'center'),
            array('value' => $label, 'vAlign' => 'center'),
            array('value' => $pageNumber, 'vAlign' => 'center')
        ));
        // Column widths are in twentieths of a pt.
        // Make date and page number the same width so label is centered.
        $params = array(
            'border' => 'nil',
            'columnWidths' => array(1400,5400,1400)
        );
        $footerTable = new WordFragment($docx, 'defaultFooter');
        $footerTable->addTable($values, $params);
        return $footerTable;
    }

    /**
     * Creates the footer text which is the portion between the date and page
     * number.
     */
    private function createFooterText($simpleXmlElement, $settings) {
        $name = $simpleXmlElement->name;
        $ecvFormat = $settings["ecvFormat"];
        if ($ecvFormat == "far") {
            $acadyear = $settings["academicYear"];
            if ($acadyear) {
                $text = $name . ', FAR ' . $acadyear;
            } else {
                $text = $name;
            }
        } else {    // must be ecv
            $text = $name . ', MacEwan CV';
        }
        return $text;
    }

    /**
     * Create HTML given XML element.
     * @param $simpleXmlElement SimpleXmlElement XML element with all required data.
     * @param $settings settings from request.
     */
    private function createHtml($simpleXmlElement, $settings) {
        // Specify our Twig templates location
        $loader = new Twig_Loader_Filesystem(__DIR__.'/templates');
        $twig = new Twig_Environment($loader, array(
            'debug' => true
        ));
        $twig->addExtension(new Twig_Extension_Debug());
        TwigFunctions::install($twig);
        $ecvFormat = $settings["ecvFormat"];
        if ($ecvFormat == null) {
            $ecvFormat = "ecv";
        }
        $ecvFormat = strtolower($ecvFormat);
        $tpl = $twig->loadTemplate('poke.html.twig');
        $html = $tpl->render(array(
            'ecv' => $simpleXmlElement,
            'ecvFormat' => $ecvFormat
        ));
        $html = PokePrintHandler::printHandler($ecvFormat, $html);
        return $html;
    }

    /**
     * Creates a unique ID for transactions, temp files, etc.
     */
    public static function createUniqueId() {
        try {
            $uuid = Uuid::uuid4()->toString();
            return $uuid;
        } catch (UnsatisfiedDependencyException $e) {
            // Hard exit since PHP doesn't have the right dependencies
            $this->recordLog("Exception creating unique ID: " . $e->getMessage());
            exit(-1);
        }
    }

    /**
     * Generates separate HTML sections for the HTML document for FAR format.
     * FAR format has a table partway through the document that is in
     * landscape mode while the rest of the document should be in portrait mode.
     * @param $htmlContent HTML document.
     * @return $filename of the document.
     */
    private function farPrint($htmlContent, $xmlElement, $settings) {
        $parser = new HtmlDomParser();
        $dom = $parser->str_get_html($htmlContent);

        // First element is the CSS which is needed on every section
        $sections = $dom->find('section[id=section]');		
        $htmlCss = array_shift($sections)->outertext;

        $outputFormat = $settings['outputFormat'];
        $sectionCount = 0;
        $uniq = self::createUniqueId();
        $sectionBase = TMP_DIR . "/" . $uniq . ".section.";
        $tmpFilenames = array();

		$currentPDFPageNumber = 0;  // Only applicable for PDF
        foreach ($sections as $section) {
            $currentContent = $section->outertext;
            if (trim($section->innertext) != false) {
                $htmlString = $htmlCss . $currentContent;
                if ($outputFormat == 'docx') {
                    $tmpdocx = new Phpdocx\Create\CreateDocx();
                    $tmpdocx->modifyPageLayout('letter', $this->marginSettings);
                                
                    if ($sectionCount == 1) {
                        $tmpdocx->modifyPageLayout('letter-landscape');
                    }	
                    $tmpdocx->embedHTML($htmlString, array('removeLineBreaks'=>True));
                    
                    // add footer
                    $footer = $this->createFooterDocx($tmpdocx, $xmlElement, $settings);
                    $tmpdocx->addFooter(array('footer' => $footer));
                    
                    //output to tmp file
                    $tmpFilename = $sectionBase . $sectionCount . '.docx';
                    $this->recordLog("Section docx file " . $tmpFilename);
                    $tmpdocx->createDocx($tmpFilename);
                    array_push($tmpFilenames, $tmpFilename);
                } else if ($outputFormat = 'pdf') {
                    if ($sectionCount == 1) {
                        $orientation = 'Landscape';
                    } else {
                        $orientation = 'Portrait';
                    }
                    $pdfFilename = $this->convertHtmlToPdfCli($htmlString,
                            $xmlElement, $settings, $currentPDFPageNumber,
                            $orientation);
                    $currentPDFPageNumber = $currentPDFPageNumber +
                            PokePrintHandler::getNumPagesPdf($pdfFilename);
                    array_push($tmpFilenames, $pdfFilename);
                }
                $sectionCount++;
            }
        }
        if ($outputFormat == 'docx') {
            $docxFilename = TMP_DIR . "/" . $uniq . ".docx";
            $mergeFirstPage = array_shift($tmpFilenames);
            $merge = new Phpdocx\Utilities\MultiMerge();
            $this->recordLog("Merging to " . $docxFilename);
            $merge->mergeDocx($mergeFirstPage, $tmpFilenames, $docxFilename, array('mergeType' => 0));
            unlink($mergeFirstPage);
            foreach ($tmpFilenames as $tmpFilename){
                unlink($tmpFilename);
            }
            return $docxFilename;
        } else if ($outputFormat == 'pdf') {
            $pdfFilename = TMP_DIR . "/" . $uniq . ".pdf";
            $pdf = new PDFMerger();
            foreach ($tmpFilenames as $tmpFilename){
                $pdf->addPDF($tmpFilename);
            }
            $this->recordLog("Merging to " . $pdfFilename);
            $pdf->merge('file', $pdfFilename);
            
            //remove tmp files
            foreach ($tmpFilenames as $tmpFilename){
                unlink($tmpFilename);
            }
            return $pdfFilename;
        } else {
            $this->recordLog("Unsupported FAR output format: ". $outputFormat);
            return null;
        }
    }

    /**
     * Returns a recommended filename for the PDF file.
     * For a FAR, the full name is incorporated into the filename, otherwise
     * the filename is simply the employee ID with a .pdf extension.
     * @param fullName full name of the person.
     */
    private function recommendedFilename($fullName, $settings) {
        $format = strtolower($settings["outputFormat"]);
        if ($format == "docx") {
            $extension = ".docx";
        } else if ($format == "html") {
            $extension = ".html";
        } else if ($format == "pdf") {
            $extension = ".pdf";
        } else if ($format == "rawxml") {
            return $settings["employeeId"] . "-raw.xml";
        } else if ($format == "xml") {
            return $settings["employeeId"] . ".xml";
        } else {
            $extension = "";
        }
        if (strtolower($settings["ecvFormat"]) == "far") {
            $acadyear = $settings["academicYear"];
            $filename = "Faculty_Annual_Report-" . $acadyear . "-" . $fullName . $extension;
            $filename = preg_replace("/\s/", "_", $filename);
            return $filename;
        } else {
            return $settings["employeeId"] . $extension;
        }
    }

    /** Record a log. */
    public function recordLog($message) {
        $log = "POKE - " . $this->transactionId . ": " . $message;
        syslog(LOG_INFO, $log);
    }

    /**
     * Requests the raw XML from PeopleSoft.
     */
    private function requestRawXml($settings) {
        $employeeId = $settings["employeeId"];
        if ($employeeId == null) {
            $this->recordLog("No employee ID");
            return null;
        }
        $ecvtype = $settings["ecvType"];
        if ($ecvtype == null) {
            $ecvtype = "self";
        } else {
            $ecvtype = strtolower($ecvtype);
        }
        if ($ecvtype == "self") {
            $uri = "get/" . $employeeId;
        } else if ($ecvtype == "far") {
            $acadYear = $settings["academicYear"];
            $uri = "far/" . $employeeId . "/" . $acadYear;
        } else {
            $this->recordLog("Unknown ecvtype " . $ecvtype);
            return null;
        }
        $client = new GuzzleHttp\Client(array('base_uri' => PS_BASE_URL));
        $clientOptions = array(
            'auth' => [ PS_USERNAME, PS_PASSWORD ],
            'headers' => [ 'Accept' => 'text/xml' ],
            'verify' => false
        );
        try {
            $this->recordLog("Retrieving XML from " . PS_BASE_URL . $uri);
            $response = $client->request('GET', $uri, $clientOptions);
            $rawXml = $response->getBody()->getContents();
            if ($rawXml === FALSE || strlen($rawXml) == 0) {
                $this->recordLog("Failed to get XML from response " . $url);
                return null;
            }
        } catch (ClientException $e) {  // 400 level errors
            $this->recordLog("ClientException: " . Psr7\str($e->getRequest()));
            $this->recordLog("ClientException (response): "
                    . Psr7\str($e->getResponse()));
            return null;
        } catch (RequestException $e) {
            $this->recordLog("RequestException: " . Psr7\str($e->getRequest()));
            if ($e->hasResponse()) {
                $this->recordLog("RequestException (response): "
                        . Psr7\str($e->getResponse()));
            }
            return null;
        } catch (ServerException $e) {  // 500 level errors
            $this->recordLog("ServerException " . Psr7\str($e->getRequest()));
            return null;
        } catch (TransferException $e) {
            $this->recordLog("TransferException " . $e->getMessage());
            return null;
        } catch (Exception $e) {
            $this->recordLog("Unexpected exception ". $e->getMessage());
            return null;
        }
        $this->recordLog("Retrieved XML from " . $uri);
        return $rawXml;
    }

    /**
     * Strip tags from the SimpleXmlElement nodes to take away any unwanted
     * formatting before creating the PDF.
     */
    private function stripTags($simpleXmlElement) {
        // Strip description, reflections, url, deliverables (based on XSL cdata section)
        $elements = array('description', 'reflections', 'url', 'deliverables');
        foreach ($elements as $element) {
            $xpath = $simpleXmlElement->xpath('//' . $element);
            array_walk($xpath, function(&$node) {
                if (count($node)) {
                    return;
                }
                $allowedTags = TwigFunctions::allowedTags();
                $node[0] = strip_tags($node[0], $allowedTags);
            });
        }
    }

    /**
     * Transforms the raw XML to document friendly XML.
     */
    private function transformXml($rawXml) {
        $xmlObj = new DOMDocument();
        $success = $xmlObj->loadXML($rawXml);
        if ($success == FALSE) {
            $this->recordLog("Failed to load XML: " . $rawXml);
            return null;
        }
        $this->recordLog("DOM created");
        $xslDoc = new DOMDocument();
        $success = $xslDoc->load(ECV_XSL_FILE);
        if ($success == FALSE) {
            $this->recordLog("Failed to load XSL");
            return null;
        }
        $this->recordLog("XSL loaded");
        $proc = new XSLTProcessor();
        $success = $proc->importStylesheet($xslDoc);
        if ($success == FALSE) {
            $this->recordLog("Failed to import XSL");
            return null;
        }
        $this->recordLog("XSLT Processor created");
        $realXml = $proc->transformToXml($xmlObj);
        if ($realXml == FALSE) {
            $this->recordLog("Failed to transform XML");
            return null;
        }
        $this->recordLog("Transformation complete");
        return $realXml;
    }

    /**
     * Validate settings returning TRUE for valid or FALSE for invalid.
     */
    private function validateSettings($settings) {
        if ($settings["employeeId"] == null) {
            $this->recordLog("No emplid");
            return FALSE;
        }
        $ecvType = $settings["ecvType"];
        if ($ecvType == null) {
            $ecvType = "self";
        } else {
            $ecvType = strtolower(trim($ecvType));
        }
        $validTypes = array("self", "far");
        if (in_array($ecvType, $validTypes) == FALSE) {
            $this->recordLog("Unsupported ecv type " . $ecvType);
            return FALSE;
        }
        if ($ecvType == "far") {
            $year = $settings["academicYear"];
            if ($year == null || trim($year) == "") {
                // FAR needs academic year
                $this->recordLog("Missing academic year for FAR");
                return FALSE;
            }
        }
        $ecvFormat = $settings["ecvFormat"];
        // Null format has a default so it's valid
        if ($ecvFormat != null) {
            $ecvFormat = strtolower(trim($ecvFormat));
            $validFormats = array("ecv", "far");
            if (in_array($ecvFormat, $validFormats) == FALSE) {
                $this->recordLog("Unsupported format " . $ecvFormat);
                return FALSE;
            }
            if ($ecvFormat == "far") {
                $year = $settings["academicYear"];
                if ($year == null || trim($year) == "") {
                    // FAR needs academic year
                    $this->recordLog("Missing academic year");
                    return FALSE;
                }
            }
        }
        $outputFormat = $settings["outputFormat"];
        if ($outputFormat == null) {
            $this->recordLog("Missing outputformat");
            return FALSE;
        }
        $validFormats = array('pdf', 'docx', 'html');
        if (DEBUG_XML) {
            array_push($validFormats, "rawxml", "xml");
        }
        if (in_array($outputFormat, $validFormats) == FALSE) {
            $this->recordLog("Unsupported outputformat " . $outputFormat);
            return FALSE;
        }
        return TRUE;
    }
}
?>