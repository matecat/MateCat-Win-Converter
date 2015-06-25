<?php

include "Exceptions.php";

class LegacyOfficeConverter {

    private static $MAX_INT_32 = 2147483648;
    private static $SOCKET_READ_TIMEOUT = 30; // in seconds
    private static $SOCKET_WRITE_TIMEOUT = 10; // in seconds

    private static $fileTypeByExtension = array(
        'doc' => 1,
        'xls' => 2,
        'ppt' => 3,
        'dot' => 4, // Legacy Word Template
        'xlt' => 5, // Legacy Excel Template
        'pot' => 6, // Legacy PowerPoint Template
        'pps' => 7, // Legacy PowerPoint slideshow
        'rtf' => 8,
    );

    private $host;
    private $port;

    public function __construct($host, $port) {
        $this->host = $host;
        $this->port = $port;
    }

    /**
     * Returns the file type constant associated with the extension of the provided
     * file, or -1 if the extension is not supported.
     */
    public static function getFileType($fileName) {
        $extension = strtolower(pathinfo($fileName, PATHINFO_EXTENSION));
        $fileType = @self::$fileTypeByExtension[$extension];
        if ($fileType != null) {
            return $fileType;
        } else {
            return -1;
        }
    }

    /**
     * Perform the file conversion on the remote server and returns the bytes of
     * the converted file, ready to be persisted (eg. with file_put_contents).
     * To obtain the correct fileType from a file name, use the getFileType
     * function of this class.
     * NOTE: Because of the communication protocol's design, you cannot send
     * files greater than 2.147.483.648 bytes, the max value of a signed 32
     * bits integer.
     */
    public function convert($fileType, $bytes) {
        // Check if file is too big
        $fileSize = strlen($bytes);
        if ($fileSize > self::$MAX_INT_32) {
            throw new FileTooBigException($fileSize);
        }

        // Create and configure the socket
        $socket = @socket_create(AF_INET, SOCK_STREAM, getprotobyname('tcp'));
        socket_set_block($socket);
        socket_set_option($socket, SOL_SOCKET, SO_RCVTIMEO, array('sec' => self::$SOCKET_READ_TIMEOUT, 'usec' => 0));
        socket_set_option($socket, SOL_SOCKET, SO_SNDTIMEO, array('sec' => self::$SOCKET_WRITE_TIMEOUT, 'usec' => 0));

        // Connect to the server
        $ok = @socket_connect($socket, $this->host, $this->port);
        if (!is_resource($socket) || $ok === false) {
            throw new SocketException();
        }

        // Start communicating
        try {
            // Send the file type
            self::writeInt($socket, $fileType);
            // Send the file size
            self::writeInt($socket, $fileSize);
            // Send the entire file
            self::write($socket, $bytes);
            // Read the server's status code
            $statusCode = self::readInt($socket);
            if ($statusCode != 0) {
                // If status code is not 0 something bad happened
                switch ($statusCode) {
                    case 1:
                        throw new BadFileTypeException($fileType);
                    case 2:
                        throw new BadFileSizeException($fileSize);
                    case 3:
                        throw new BrokenFileException();
                    case 4:
                        throw new ConvertedFileTooBig();
                    case 5:
                        throw new InternalServerError();
                    default:
                        throw new Exception("Conversion server sent an unknown status code: " . $statusCode);
                }
            }
            // Read the converted file size
            $convertedSize = self::readInt($socket);
            // Read the entire converted file
            $bytes = self::read($socket, $convertedSize);
            // Ok, everything went fine!
            socket_close($socket);
            return $bytes;

        } catch (Exception $e) {
            // A finalize would be better, but this code runs on PHP 5.4
            socket_close($socket);
            throw $e;
        }
    }

    private static function read($socket, $length) {
        $data = '';
        $remaining = $length;

        do {
            $bytesRead = socket_read($socket, $remaining, PHP_BINARY_READ);
            if ($bytesRead === false) {
                throw new SocketException();
            } elseif ($bytesRead === '') {
                throw new ProtocolException("Input ended unexpectedly: stream ended after " . strlen($data) . " bytes, expecting " . $length . " bytes");
            }
            $data .= $bytesRead;
            $remaining = $length - strlen($data);
        } while ($remaining > 0);

        return $data;
    }

    private static function write($socket, $bytes) {
        $bytesWritten = socket_write($socket, $bytes);
        if ($bytesWritten === false) {
            throw new SocketException();
        }
        return $bytesWritten;
    }

    private static function writeInt($socket, $intValue) {
        return self::write($socket, pack('N', $intValue));
    }

    private static function readInt($socket) {
        $intBytes = self::read($socket, 4);
        return current(unpack('N', $intBytes));
    }

}
