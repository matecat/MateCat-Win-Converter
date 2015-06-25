<?php

class ProtocolException extends Exception {
}

class SocketException extends Exception {
    public function __construct() {
        $message = "Socket error " . socket_last_error() . ": " . socket_strerror( socket_last_error() );
        parent::__construct($message, 0, null);
    }
}

class BadFileTypeException extends Exception {
    public function __construct($fileType) {
        $message = "Conversion server does not support the file type $fileType";
        parent::__construct($message, 0, null);
    }
}

class FileTooBigException extends Exception {
    public function __construct($fileSize) {
        $message = "Files greater than 2^31 bytes are not supported: provided file is " . $fileSize . " bytes long";
        parent::__construct($message, 0, null);
    }
}

class BadFileSizeException extends Exception {
    public function __construct($fileSize) {
        $message = "Conversion server received a file size smaller or equal to 0";
        parent::__construct($message, 0, null);
    }
}

class BrokenFileException extends Exception {
    public function __construct() {
        $message = "Conversion server could not convert the provided file";
        parent::__construct($message, 0, null);
    }
}

class ConvertedFileTooBig extends Exception {
    public function __construct() {
        $message = "Conversion server could not send back the converted file because it's too big";
        parent::__construct($message, 0, null);
    }
}

class InternalServerError extends Exception {
    public function __construct() {
        $message = "See the conversion server's log for further informations";
        parent::__construct($message, 0, null);
    }
}
