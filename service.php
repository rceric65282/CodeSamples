<?php

require_once("vendor/autoload.php");
require_once("config.php");
require_once("poke.php");
require_once("poke-req-handler.php");

/**
 * Service endpoint for calling Poke API.
 */
class PokeService {
    private $handler;
    private $poke;

    function __construct() {
        openlog("POKE", LOG_INFO, LOG_USER);
        $this->handler = new PokeRequestHandler();
        $this->poke = new Poke();
    }

    /**
     * Handles the incoming request.
     */
    public function handleRequest() {
        /*
         * Increase PHP timeout in case retrieving from PeopleSoft and
         * transforming the data take a while.
         */
        set_time_limit(1800);
        $this->poke->recordLog("Transaction started");
        try {
            $credentials = $this->handler->getBasicAuthCredentials();
            if ($credentials == null) {
                $this->handler->setStatusForbidden();
                $this->poke->recordLog("No basic auth");
                return FALSE;
            }
            // Check basic auth
            if ($credentials['username'] != AUTH_USERNAME ||
                    $credentials['password'] != AUTH_PASSWORD) {
                $this->handler->setStatusForbidden();
                $this->poke->recordLog("Bad basic auth");
                return FALSE;
            }
   
            $settings = $this->settingsFromRequest();
    
            $requestMethod = $this->handler->getRequestMethod();
            if ($requestMethod == 'GET' || $requestMethod == 'POST') {
                $result = $this->poke->createDocument($settings);
                if ($result == null) {
                    $this->handler->setStatusInternalError();
                    return;
                }
                $filename = $result["filename"];
                $docFilename = $result["document"];
                if ($docFilename == null || $filename == null) {
                    $this->poke->recordLog("Didn't get a document file and a filename.");
                    $this->handler->setStatusInternalError();
                    return;
                }
                $format = $settings["outputFormat"];
                if ($format == "docx") {
                    $contentType = PokeRequestHandler::DOCX_CONTENT_TYPE;
                } else if ($format == "html") {
                    $contentType = PokeRequestHandler::HTML_CONTENT_TYPE;
                } else if ($format == "xml" || $format == "rawxml") {
                    $contentType = PokeRequestHandler::XML_CONTENT_TYPE;
                } else if ($format == "pdf") {
                    $contentType = PokeRequestHandler::PDF_CONTENT_TYPE;
                } else {
                    // Don't know what format, so just make it a binary stream
                    $contentType = "application/octet-stream";
                }
                $sent = $this->handler->sendContent($docFilename, $contentType, $filename);
                if ($sent == FALSE) {
                    $this->poke->recordLog("Couldn't send document");
                    $this->handler->setStatusInternalError();
                }
                unlink($docFilename);
            } else {
                $this->poke->recordLog("Unsupported request method " . $requestMethod);
                $this->handler->setStatusBadRequest();
                return FALSE;
            }
            return TRUE;
        } catch (Exception $e) {
            $this->poke->recordLog("Unexpected exception: " . $e->getMessage());
            $this->handler->setStatusInternalError();
            return FALSE;
        } finally {
            $this->poke->recordLog("Transaction complete");
        }
    }

    /**
     * Return settings gathered from the request.
     */
    private function settingsFromRequest() {
        $settings = array();
        $settings["academicYear"] = $this->handler->getAcademicYear();
        $settings["employeeId"] = $this->handler->getEmployeeId();
        $settings["ecvFormat"] = $this->handler->getEcvFormat();
        $settings["ecvType"] = $this->handler->getEcvType();
        $settings["operatorId"] = $this->handler->getOperatorId();
        $settings["outputFormat"] = $this->handler->getOutputFormat();
        $settings["useContentDisposition"] =
                $this->handler->getUseContentDisposition();
        return $settings;
    }
}

$service = new PokeService();
$service->handleRequest();

?>